sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], (Controller, MessageToast, MessageBox) => {
    "use strict";

    return Controller.extend("codelco.portaldescuentos.controller.Home", {

        onInit() {
            // Cada vez que se entra al Home (también al volver de Carga) se recalcula todo
            this.getOwnerComponent().getRouter()
                .getRoute("RouteHome")
                .attachPatternMatched(this._onRouteMatched, this);
        },

        _onRouteMatched() {
            const oComponent = this.getOwnerComponent();
            oComponent.getModel("mock").dataLoaded().then(() => {
                oComponent.recalcularPeriodo();
                this._armarDatosHome();
            });
        },

        /** Subtítulo, resumen de motivos e historial (todo calculado, nada fijo) */
        _armarDatosHome() {
            const oApp = this.getOwnerComponent().getModel("app");
            const oMock = this.getOwnerComponent().getModel("mock");
            const oPeriodo = oApp.getProperty("/periodo");
            const oCargaActual = oApp.getProperty("/cargaActual");
            const aCargas = oMock.getProperty("/cargas") || [];

            // Subtítulo según estado del período
            const sSubtitulo = oPeriodo.abierto
                ? `Carga habilitada hasta el ${oPeriodo.cierre} a las 23:59 h`
                : `La carga del período ${oPeriodo.texto} venció el ${oPeriodo.cierre}. ` +
                  `Próximo período: del ${oPeriodo.proximaApertura} al ${oPeriodo.proximoCierre}.`;
            oApp.setProperty("/subtitulo", sSubtitulo);

            // Resumen de la carga actual (KPIs + motivos agrupados)
            this._armarResumen(oCargaActual);

            // Historial: todas las cargas excepto la del período actual, más recientes primero
            const aHistorial = aCargas
                .filter(c => !(c.periodoMes === oPeriodo.mes && c.periodoAnio === oPeriodo.anio))
                .sort((a, b) => (b.periodoAnio * 100 + b.periodoMes) - (a.periodoAnio * 100 + a.periodoMes));
            oApp.setProperty("/historial", aHistorial);
        },

        /**
         * Arma KPIs y tabla de motivos.
         * - Antes del proceso SAP: conformes / observados (validación en BTP).
         * - Después del proceso SAP: cargados / rechazados.
         */
        _armarResumen(oCarga) {
            const oApp = this.getOwnerComponent().getModel("app");
            const aCatalogo = this.getOwnerComponent().getModel("mock").getProperty("/motivos") || [];

            if (!oCarga) {
                oApp.setProperty("/resumen", {});
                oApp.setProperty("/resumenMotivos", []);
                return;
            }

            const bProcesada = !!oCarga.procesada;
            const aDetalle = oCarga.detalle || [];
            const sEstadoError = bProcesada ? "Rechazado" : "Observado";
            const aErrores = aDetalle.filter(d => d.estado === sEstadoError);

            // Importe total de los registros válidos, separado por moneda (CLP y UF no se suman)
            const sumar = sMoneda => aDetalle
                .filter(d => d.estado !== sEstadoError && d.moneda === sMoneda)
                .reduce((s, d) => s + (Number(d.importe) || 0), 0);
            const fCLP = sumar("CLP");
            const fUF = sumar("UF");

            oApp.setProperty("/resumen", {
                subtitulo: bProcesada
                    ? `${oCarga.totalRegistros} registros informados · procesados el ${oCarga.fechaProceso}`
                    : `${oCarga.totalRegistros} registros informados · enviado el ${oCarga.fechaCarga} · pendiente de proceso en SAP`,
                informados: oCarga.totalRegistros,
                ok: bProcesada ? oCarga.totalCargados : oCarga.totalConformes,
                okLabel: bProcesada ? "Procesados y cargados al IT0015" : "Conformes · pendientes de proceso",
                error: aErrores.length,
                errorLabel: bProcesada
                    ? `Rechazados · se informan en ${this.getOwnerComponent().getModel("app").getProperty("/periodo/siguienteTexto")}`
                    : "Con observación · quedarán rechazados",
                importeCLP: Math.round(fCLP).toLocaleString("es-CL"),
                importeUF: fUF ? fUF.toLocaleString("es-CL", { maximumFractionDigits: 2 }) : ""
            });

            // Agrupar por motivo
            const mMotivos = {};
            aErrores.forEach(d => { mMotivos[d.mensaje] = (mMotivos[d.mensaje] || 0) + 1; });
            const iMax = Math.max(1, ...Object.values(mMotivos));
            const aMotivos = Object.keys(mMotivos)
                .map(m => ({
                    motivo: m,
                    cantidad: mMotivos[m],
                    porcentaje: Math.round((mMotivos[m] / iMax) * 100),
                    accion: (aCatalogo.find(c => c.mensaje === m) || {}).accion || "Revise el registro y vuelva a informarlo."
                }))
                .sort((a, b) => b.cantidad - a.cantidad);
            oApp.setProperty("/resumenMotivos", aMotivos);
        },

        /** Exporta el resumen de motivos a Excel */
        onExportarResumen() {
            if (!window.XLSX) {
                MessageBox.warning("La librería Excel todavía se está cargando.");
                return;
            }
            const oApp = this.getOwnerComponent().getModel("app");
            const aFilas = (oApp.getProperty("/resumenMotivos") || []).map(m => ({
                "Motivo de rechazo": m.motivo,
                "Registros": m.cantidad,
                "Qué debe corregir": m.accion
            }));
            const oHoja = window.XLSX.utils.json_to_sheet(aFilas);
            const oLibro = window.XLSX.utils.book_new();
            window.XLSX.utils.book_append_sheet(oLibro, oHoja, "Resumen");
            window.XLSX.writeFile(oLibro, `resumen_carga_${oApp.getProperty("/periodo/texto").replace("/", "-")}.xlsx`);
        },

        /** Genera un Excel vacío con las columnas de la plantilla (usa xlsx.safe.js) */
        onDescargarPlantilla() {
            if (!window.XLSX || !window.XLSX.writeFile) {
                MessageBox.warning("La librería Excel todavía se está cargando. Intente nuevamente en unos segundos.");
                return;
            }
            const XLSX = window.XLSX;
            const aColumnas = this.getOwnerComponent().getModel("mock").getProperty("/parametros/columnasPlantilla");

            const oHoja = XLSX.utils.aoa_to_sheet([aColumnas]);
            oHoja["!cols"] = aColumnas.map(() => ({ wch: 16 }));
            const oLibro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(oLibro, oHoja, "Descuentos");
            XLSX.writeFile(oLibro, "plantilla_descuentos.xlsx");

            MessageToast.show("Plantilla descargada");
        },

        // ================= 🧪 Herramientas de prueba (mock) =================

        onSimularProcesoSAP() {
            if (this.getOwnerComponent().simularProcesoSAP()) {
                this._armarDatosHome();
                MessageToast.show("Proceso SAP simulado");
            }
        },

        onLimpiarPruebas() {
            MessageBox.confirm("Se borrarán todas las cargas de prueba guardadas en este navegador. ¿Continuar?", {
                onClose: (sAccion) => {
                    if (sAccion === MessageBox.Action.OK) {
                        this.getOwnerComponent().limpiarCargas();
                        this._armarDatosHome();
                        MessageToast.show("Datos de prueba eliminados");
                    }
                }
            });
        },

        onSubirArchivo() {
            this.getOwnerComponent().getRouter().navTo("RouteCarga");
        },

        onVerDetalleActual() {
            const oPeriodo = this.getOwnerComponent().getModel("app").getProperty("/periodo");
            this._navConsulta(oPeriodo.mes, oPeriodo.anio);
        },

        onVerHistorial(oEvent) {
            const oCarga = oEvent.getSource().getBindingContext("app").getObject();
            this._navConsulta(oCarga.periodoMes, oCarga.periodoAnio);
        },

        _navConsulta(iMes, iAnio) {
            // Formato de la ruta: consulta/2026-11
            this.getOwnerComponent().getRouter().navTo("RouteConsulta", {
                periodo: `${iAnio}-${String(iMes).padStart(2, "0")}`
            });
        }
    });
});