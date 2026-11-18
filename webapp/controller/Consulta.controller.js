sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/m/MessageToast"
], (Controller, JSONModel, MessageBox, MessageToast) => {
    "use strict";

    return Controller.extend("codelco.portaldescuentos.controller.Consulta", {

        onInit() {
            this.getView().setModel(new JSONModel({}), "consulta");
            this.getOwnerComponent().getRouter()
                .getRoute("RouteConsulta")
                .attachPatternMatched(this._onRouteMatched, this);
        },

        /** Ruta: consulta/{periodo}  →  periodo = "2026-11" */
        _onRouteMatched(oEvent) {
            const sPeriodo = oEvent.getParameter("arguments").periodo;
            const [sAnio, sMes] = sPeriodo.split("-");
            const oComponent = this.getOwnerComponent();

            oComponent.getModel("mock").dataLoaded().then(() => {
                const oMock = oComponent.getModel("mock");
                const oCarga = (oMock.getProperty("/cargas") || []).find(c =>
                    c.periodoAnio === Number(sAnio) && c.periodoMes === Number(sMes)) || null;
                this._cargar(oCarga, `${sMes}/${sAnio}`);
            });
        },

        _cargar(oCarga, sPeriodoTexto) {
            const oMock = this.getOwnerComponent().getModel("mock");
            const aCatalogo = oMock.getProperty("/motivos") || [];
            const aConceptos = oMock.getProperty("/conceptosAutorizados") || [];

            // Agregar "acción sugerida" a cada registro
            const aDetalle = (oCarga?.detalle || []).map(d => Object.assign({}, d, {
                accion: d.mensaje && d.estado !== "Cargado"
                    ? ((aCatalogo.find(c => c.mensaje === d.mensaje) || {}).accion || "Revise el registro.")
                    : "—"
            }));

            const bProcesada = !!oCarga?.procesada;
            const sEstadoError = bProcesada ? "Rechazado" : "Observado";
            const iTotal = aDetalle.length;
            const iError = aDetalle.filter(d => d.estado === sEstadoError).length;
            const iOk = iTotal - iError;
            const iMotivos = new Set(aDetalle.filter(d => d.estado === sEstadoError).map(d => d.mensaje)).size;

            // Estados posibles para el filtro según la etapa
            const aEstados = [{ key: "", text: "Todos" }].concat(bProcesada
                ? [{ key: "Cargado", text: "Cargado" }, { key: "Rechazado", text: "Rechazado" }]
                : [{ key: "Conforme", text: "Conforme" }, { key: "Observado", text: "Observado" }]);

            this.getView().getModel("consulta").setData({
                carga: oCarga,
                periodoTexto: sPeriodoTexto,
                subtitulo: oCarga
                    ? `${oCarga.nombreArchivo} · enviado el ${oCarga.fechaCarga} · ${oCarga.estado}`
                    : "",
                detalleCompleto: aDetalle,
                registros: aDetalle,
                cantidadFiltrada: aDetalle.length,
                estados: aEstados,
                conceptos: [{ key: "", text: "Todos" }].concat(aConceptos.map(c => ({ key: c.ccnomina, text: `${c.ccnomina} · ${c.descripcion}` }))),
                filtros: { estado: "", concepto: "", rut: "" },
                kpi: {
                    total: iTotal,
                    ok: iOk,
                    okPct: iTotal ? (Math.round((iOk / iTotal) * 1000) / 10).toLocaleString("es-CL") : 0,
                    okLabel: bProcesada ? "Procesados" : "Conformes · pendientes de proceso",
                    error: iError,
                    errorTitulo: bProcesada ? "A informar el mes siguiente" : "Con observación",
                    errorLabel: bProcesada ? "Rechazados" : "Quedarán rechazados",
                    motivos: iMotivos
                }
            });
        },

        // ================= Filtros =================

        onFiltrar() {
            const oModel = this.getView().getModel("consulta");
            const oF = oModel.getProperty("/filtros");
            const sRut = (oF.rut || "").replace(/\./g, "").toUpperCase();

            const aFiltrados = (oModel.getProperty("/detalleCompleto") || []).filter(d =>
                (!oF.estado || d.estado === oF.estado) &&
                (!oF.concepto || d.ccnomina === oF.concepto) &&
                (!sRut || String(d.rut).replace(/\./g, "").toUpperCase().includes(sRut)));

            oModel.setProperty("/registros", aFiltrados);
            oModel.setProperty("/cantidadFiltrada", aFiltrados.length);
        },

        onLimpiarFiltros() {
            this.getView().getModel("consulta").setProperty("/filtros", { estado: "", concepto: "", rut: "" });
            this.onFiltrar();
        },

        // ================= Exportar =================

        onExportar() {
            this._exportar(this.getView().getModel("consulta").getProperty("/registros"), "registros");
        },

        onExportarRechazados() {
            const oModel = this.getView().getModel("consulta");
            const sEstado = oModel.getProperty("/carga/procesada") ? "Rechazado" : "Observado";
            const aRech = (oModel.getProperty("/detalleCompleto") || []).filter(d => d.estado === sEstado);
            this._exportar(aRech, "rechazados");
        },

        _exportar(aRegistros, sSufijo) {
            if (!window.XLSX) {
                MessageBox.warning("La librería Excel todavía se está cargando.");
                return;
            }
            if (!aRegistros || !aRegistros.length) {
                MessageToast.show("No hay registros para exportar");
                return;
            }
            const aFilas = aRegistros.map(d => ({
                "Línea": d.linea,
                "RUT": d.rut,
                "CCNOMINA": d.ccnomina,
                "IMPORTE": d.importe,
                "MONEDA": d.moneda,
                "Estado": d.estado,
                "Motivo": d.mensaje || "",
                "Acción sugerida": d.accion === "—" ? "" : d.accion
            }));
            const oHoja = window.XLSX.utils.json_to_sheet(aFilas);
            const oLibro = window.XLSX.utils.book_new();
            window.XLSX.utils.book_append_sheet(oLibro, oHoja, "Registros");
            const sPeriodo = this.getView().getModel("consulta").getProperty("/periodoTexto").replace("/", "-");
            window.XLSX.writeFile(oLibro, `${sSufijo}_${sPeriodo}.xlsx`);
        },

        onVolver() {
            this.getOwnerComponent().getRouter().navTo("RouteHome", {}, true);
        }
    });
});