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

            // Resumen de motivos de rechazo de la carga actual (agrupado por mensaje)
            const mMotivos = {};
            (oCargaActual?.detalle || [])
                .filter(d => d.estado === "Rechazado")
                .forEach(d => { mMotivos[d.mensaje] = (mMotivos[d.mensaje] || 0) + 1; });
            const aMotivos = Object.keys(mMotivos)
                .map(m => ({ motivo: m, cantidad: mMotivos[m] }))
                .sort((a, b) => b.cantidad - a.cantidad);
            oApp.setProperty("/resumenMotivos", aMotivos);

            // Historial: todas las cargas excepto la del período actual, más recientes primero
            const aHistorial = aCargas
                .filter(c => !(c.periodoMes === oPeriodo.mes && c.periodoAnio === oPeriodo.anio))
                .sort((a, b) => (b.periodoAnio * 100 + b.periodoMes) - (a.periodoAnio * 100 + a.periodoMes));
            oApp.setProperty("/historial", aHistorial);
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