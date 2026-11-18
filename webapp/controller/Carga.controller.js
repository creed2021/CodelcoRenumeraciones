sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], (Controller, JSONModel, Filter, FilterOperator, MessageToast, MessageBox) => {
    "use strict";

    // Columnas obligatorias del archivo (según DFT)
    const COLUMNAS = ["RUT", "CCNOMINA", "IMPORTE", "MONEDA"];

    return Controller.extend("codelco.portaldescuentos.controller.Carga", {

        onInit() {
            this.getView().setModel(new JSONModel(this._estadoInicial()), "carga");
            this.getOwnerComponent().getRouter()
                .getRoute("RouteCarga")
                .attachPatternMatched(this._onRouteMatched, this);
        },

        _onRouteMatched() {
            const oComponent = this.getOwnerComponent();
            oComponent.getModel("mock").dataLoaded().then(() => {
                oComponent.recalcularPeriodo();
                this._limpiar();

                // Si no se puede cargar (período cerrado o ya enviado) se vuelve al Home
                const oApp = oComponent.getModel("app");
                if (!oApp.getProperty("/puedeCargar")) {
                    MessageBox.warning(oApp.getProperty("/motivoNoCarga"), {
                        onClose: () => oComponent.getRouter().navTo("RouteHome", {}, true)
                    });
                }
            });
        },

        _estadoInicial() {
            return {
                nombreArchivo: "", tamanio: "", textoLectura: "", estructuraOk: false,
                filas: [], todasLasFilas: [], conformes: 0, observados: 0,
                checks: [], filtro: "todos"
            };
        },

        _limpiar() {
            this.getView().getModel("carga").setData(this._estadoInicial());
            const oFU = this.byId("fuExcel");
            if (oFU) { oFU.clear(); }
            const oBinding = this.byId("tblPreview").getBinding("items");
            if (oBinding) { oBinding.filter([]); }
        },

        // ================= Lectura del Excel =================

        onTipoInvalido() {
            MessageBox.error("El archivo debe ser Excel (.xlsx o .xls).");
        },

        onArchivoSeleccionado(oEvent) {
            if (!window.XLSX) {
                MessageBox.warning("La librería Excel todavía se está cargando. Espere un momento y vuelva a intentar.");
                return;
            }
            const oFile = oEvent.getParameter("files")?.[0];
            if (!oFile) {
                return;
            }

            const oReader = new FileReader();
            oReader.onload = (e) => {
                try {
                    const oLibro = window.XLSX.read(new Uint8Array(e.target.result), { type: "array" });
                    const oHoja = oLibro.Sheets[oLibro.SheetNames[0]];
                    // header:1 → arreglo de filas, para poder leer los encabezados aunque no haya datos
                    const aFilas = window.XLSX.utils.sheet_to_json(oHoja, { header: 1, defval: "", raw: true });
                    this._procesarArchivo(oFile, aFilas);
                } catch (err) {
                    console.error(err);
                    MessageBox.error("No se pudo leer el archivo. Verifique que sea un Excel válido.");
                    this._limpiar();
                }
            };
            oReader.readAsArrayBuffer(oFile);
        },

        _procesarArchivo(oFile, aFilas) {
            const oMock = this.getOwnerComponent().getModel("mock");
            const iMax = oMock.getProperty("/parametros/maxRegistros");

            // ---------- 1) Validación de estructura (si falla, no se carga nada) ----------
            if (!aFilas.length) {
                return this._errorEstructura("El archivo está vacío.");
            }
            const aEncabezados = aFilas[0].map(h => String(h).trim().toUpperCase());
            const aFaltantes = COLUMNAS.filter(c => !aEncabezados.includes(c));
            if (aFaltantes.length) {
                return this._errorEstructura(
                    `El Excel no tiene el formato correcto. Faltan las columnas: ${aFaltantes.join(", ")}.\n\n` +
                    "Descargue la plantilla desde la pantalla de inicio.");
            }
            const aDatos = aFilas.slice(1).filter(f => f.some(v => String(v).trim() !== ""));
            if (!aDatos.length) {
                return this._errorEstructura("El archivo no contiene registros.");
            }
            if (aDatos.length > iMax) {
                return this._errorEstructura(`El archivo supera el máximo de ${iMax.toLocaleString("es-CL")} registros.`);
            }
            const aExtra = aEncabezados.filter(h => h && !COLUMNAS.includes(h));
            if (aExtra.length) {
                MessageToast.show(`Se ignorarán las columnas: ${aExtra.join(", ")}`);
            }

            // ---------- 2) Validación fila por fila ----------
            const idx = c => aEncabezados.indexOf(c);
            const aRegistros = aDatos.map((f, i) => this._validarFila({
                linea: i + 1,
                rut: String(f[idx("RUT")]).trim(),
                ccnomina: String(f[idx("CCNOMINA")]).trim(),
                importe: f[idx("IMPORTE")],
                moneda: String(f[idx("MONEDA")]).trim().toUpperCase()
            }));
            this._marcarDuplicados(aRegistros);

            // ---------- 3) Resultado ----------
            const iObs = aRegistros.filter(r => r.estado === "Observado").length;
            const oCarga = this.getView().getModel("carga");
            oCarga.setData({
                nombreArchivo: oFile.name,
                tamanio: `${Math.max(1, Math.round(oFile.size / 1024))} KB · leído ${new Date().toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit" })}`,
                textoLectura: `Archivo leído correctamente · ${aRegistros.length} filas detectadas`,
                estructuraOk: true,
                filas: aRegistros,
                todasLasFilas: aRegistros,
                conformes: aRegistros.length - iObs,
                observados: iObs,
                checks: this._armarChecks(aRegistros),
                filtro: "todos"
            });
        },

        _errorEstructura(sMensaje) {
            MessageBox.error(sMensaje, { title: "Formato de archivo incorrecto" });
            this._limpiar();
        },

        /** Valida una fila. Devuelve el registro con estado Conforme/Observado y mensaje. */
        _validarFila(r) {
            const oMock = this.getOwnerComponent().getModel("mock");
            const aConceptos = (oMock.getProperty("/conceptosAutorizados") || []).map(c => c.ccnomina);
            const aMonedas = oMock.getProperty("/parametros/monedas") || [];

            const oImporte = this._parseImporte(r.importe);
            r.importeNumero = oImporte.valor;
            r.importeTexto = oImporte.valido
                ? oImporte.valor.toLocaleString("es-CL", { maximumFractionDigits: 2 })
                : String(r.importe ?? "");

            let sError = "";
            if (!r.rut) {
                sError = "RUT vacío";
            } else if (!this._rutValido(r.rut)) {
                sError = "RUT inválido";
            } else if (!aConceptos.includes(r.ccnomina)) {
                sError = "Concepto no autorizado";
            } else if (r.importe === "" || r.importe === null || r.importe === undefined) {
                sError = "Importe vacío";
            } else if (!oImporte.valido) {
                sError = "Formato importe erróneo";
            } else if (oImporte.valor <= 0) {
                sError = "Importe debe ser mayor a cero";
            } else if (!aMonedas.includes(r.moneda)) {
                sError = "Moneda no admitida";
            }

            r.rut = this._formatearRut(r.rut);
            r.estado = sError ? "Observado" : "Conforme";
            r.mensaje = sError;
            return r;
        },

        /** Duplicados: mismo RUT + concepto, salvo conceptos que admiten múltiples líneas (Cajas) */
        _marcarDuplicados(aRegistros) {
            const aMultiples = (this.getOwnerComponent().getModel("mock").getProperty("/conceptosAutorizados") || [])
                .filter(c => c.multiplesLineas).map(c => c.ccnomina);
            const mVistos = {};
            aRegistros.forEach(r => {
                if (r.estado !== "Conforme" || aMultiples.includes(r.ccnomina)) { return; }
                const sClave = `${r.rut}|${r.ccnomina}`;
                if (mVistos[sClave]) {
                    r.estado = "Observado";
                    r.mensaje = "Registro duplicado archivo";
                }
                mVistos[sClave] = true;
            });
        },

        _armarChecks(aRegistros) {
            const contar = aMsgs => aRegistros.filter(r => aMsgs.includes(r.mensaje)).length;
            const check = (titulo, aMsgs, sOk) => {
                const n = contar(aMsgs);
                return { titulo, ok: n === 0, detalle: n === 0 ? sOk : `${n} registro(s) con observación` };
            };
            return [
                { titulo: "Estructura y tipos de dato", ok: true, detalle: `${COLUMNAS.length} columnas · ${COLUMNAS.join(", ")}` },
                check("RUT del trabajador", ["RUT vacío", "RUT inválido"], "Todos los RUT son válidos"),
                check("Conceptos autorizados", ["Concepto no autorizado"], "Todos los conceptos están habilitados"),
                check("Importes y moneda", ["Importe vacío", "Formato importe erróneo", "Importe debe ser mayor a cero", "Moneda no admitida"], "Importes numéricos en CLP o UF"),
                check("Duplicados", ["Registro duplicado archivo"], "Sin líneas duplicadas")
            ];
        },

        // ================= Utilidades RUT / importe =================

        _rutValido(sRut) {
            const s = String(sRut).replace(/\./g, "").replace(/\s/g, "").toUpperCase();
            const m = s.match(/^(\d{1,8})-?([\dK])$/);
            if (!m) { return false; }
            return this._calcularDV(m[1]) === m[2];
        },

        _calcularDV(sCuerpo) {
            let iSuma = 0, iMult = 2;
            for (let i = sCuerpo.length - 1; i >= 0; i--) {
                iSuma += Number(sCuerpo[i]) * iMult;
                iMult = iMult === 7 ? 2 : iMult + 1;
            }
            const r = 11 - (iSuma % 11);
            return r === 11 ? "0" : r === 10 ? "K" : String(r);
        },

        /** Normaliza a 12.345.678-9 (si es válido) */
        _formatearRut(sRut) {
            const s = String(sRut).replace(/\./g, "").replace(/\s/g, "").toUpperCase();
            const m = s.match(/^(\d{1,8})-?([\dK])$/);
            if (!m) { return String(sRut); }
            return `${Number(m[1]).toLocaleString("es-CL")}-${m[2]}`;
        },

        /** Igual que en Carrefour, pero indica si el valor es válido */
        _parseImporte(v) {
            if (typeof v === "number") { return { valido: !isNaN(v), valor: v }; }
            let s = String(v ?? "").trim().replace(/\s/g, "");
            if (!s) { return { valido: false, valor: 0 }; }
            if (s.includes(",")) { s = s.replace(/\./g, "").replace(",", "."); }
            if (!/^-?\d+(\.\d+)?$/.test(s)) { return { valido: false, valor: 0 }; }
            return { valido: true, valor: parseFloat(s) };
        },

        // ================= Acciones =================

        onFiltroPreview(oEvent) {
            const sKey = oEvent.getParameter("item").getKey();
            const oBinding = this.byId("tblPreview").getBinding("items");
            oBinding.filter(sKey === "obs" ? [new Filter("estado", FilterOperator.EQ, "Observado")] : []);
        },

        onQuitar() {
            this._limpiar();
        },

        onCancelar() {
            this._limpiar();
            this.getOwnerComponent().getRouter().navTo("RouteHome", {}, true);
        },

        onConfirmar() {
            const oDatos = this.getView().getModel("carga").getData();
            const iObs = oDatos.observados;
            const sMsg = iObs > 0
                ? `Se enviarán ${oDatos.todasLasFilas.length} registros. ${iObs} con observación quedarán rechazados.\n\n` +
                  "Una vez confirmado no podrá volver a cargar archivo en este período. ¿Desea continuar?"
                : `Se enviarán ${oDatos.todasLasFilas.length} registros.\n\n` +
                  "Una vez confirmado no podrá volver a cargar archivo en este período. ¿Desea continuar?";

            MessageBox.confirm(sMsg, {
                title: "Confirmar envío",
                actions: [MessageBox.Action.OK, MessageBox.Action.CANCEL],
                emphasizedAction: MessageBox.Action.OK,
                onClose: (sAccion) => {
                    if (sAccion === MessageBox.Action.OK) {
                        this._guardarCarga(oDatos);
                    }
                }
            });
        },

        /** 🧪 Mock: guarda la carga en el modelo (después será una acción del CAP) */
        _guardarCarga(oDatos) {
            const oComponent = this.getOwnerComponent();
            const oPeriodo = oComponent.getModel("app").getProperty("/periodo");
            const aDetalle = oDatos.todasLasFilas.map(r => ({
                linea: r.linea, rut: r.rut, ccnomina: r.ccnomina,
                importe: r.importeNumero, importeTexto: r.importeTexto, moneda: r.moneda,
                estado: r.estado, mensaje: r.mensaje
            }));

            oComponent.agregarCarga({
                id: `${oPeriodo.anio}-${String(oPeriodo.mes).padStart(2, "0")}`,
                periodoMes: oPeriodo.mes,
                periodoAnio: oPeriodo.anio,
                periodoTexto: oPeriodo.texto,
                nombreArchivo: oDatos.nombreArchivo,
                fechaCarga: oComponent.formatearFechaHora(oComponent.getFechaHoy()),
                estado: "Enviado · pendiente de proceso",
                estadoState: "Information",
                procesada: false,
                totalRegistros: aDetalle.length,
                totalConformes: oDatos.conformes,
                totalObservados: oDatos.observados,
                totalCargados: 0,
                totalRechazados: 0,
                detalle: aDetalle
            });

            this._limpiar();
            MessageToast.show("Archivo enviado correctamente");
            oComponent.getRouter().navTo("RouteHome", {}, true);
        }
    });
});