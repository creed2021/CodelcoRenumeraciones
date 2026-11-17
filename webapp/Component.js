sap.ui.define([
    "sap/ui/core/UIComponent",
    "sap/ui/model/json/JSONModel",
    "codelco/portaldescuentos/model/models",
    "codelco/portaldescuentos/model/periodo"
], (UIComponent, JSONModel, models, periodo) => {
    "use strict";

    return UIComponent.extend("codelco.portaldescuentos.Component", {
        metadata: {
            manifest: "json",
            interfaces: ["sap.ui.core.IAsyncContentCreation"]
        },

        init() {
            // 🔹 Cargar XLSX solo si no existe (mismo esquema que Carrefour, funciona también fuera de index.html)
            this._cargarXLSX();

            UIComponent.prototype.init.apply(this, arguments);

            // Modelo de dispositivo (generado por el asistente)
            this.setModel(models.createDeviceModel(), "device");

            // Modelo "app": estado calculado del período (no se guarda, se recalcula)
            this.setModel(new JSONModel({}), "app");

            // Cuando el mock termina de cargar, calculamos el período actual
            const oMock = this.getModel("mock");
            oMock.dataLoaded().then(() => {
                this.recalcularPeriodo();
            });

            this.getRouter().initialize();
        },

        /**
         * Calcula período actual, fecha límite y estado (abierto/cerrado).
         * Para simular otra fecha: agregar ?fecha=2026-11-10 a la URL.
         */
        recalcularPeriodo() {
            const oMock = this.getModel("mock");
            const oParams = oMock.getProperty("/parametros");
            const aCargas = oMock.getProperty("/cargas") || [];

            const oHoy = this._getFechaHoy();
            const oPeriodo = periodo.calcular(oHoy, oParams.diaCierre);

            // ¿Ya hay carga para el período actual?
            const oCargaActual = aCargas.find(c =>
                c.periodoMes === oPeriodo.mes && c.periodoAnio === oPeriodo.anio) || null;

            this.getModel("app").setData({
                hoy: oHoy,
                periodo: oPeriodo,
                cargaActual: oCargaActual,
                puedeCargar: oPeriodo.abierto && !oCargaActual,
                motivoNoCarga: !oPeriodo.abierto
                    ? "El período de carga está cerrado."
                    : (oCargaActual ? "Ya envió un archivo para este período." : "")
            });
        },

        _getFechaHoy() {
            const sFecha = new URLSearchParams(window.location.search).get("fecha");
            if (sFecha) {
                const oFecha = new Date(sFecha + "T12:00:00");
                if (!isNaN(oFecha)) {
                    return oFecha;
                }
            }
            return new Date();
        },

        _cargarXLSX() {
            if (window.XLSX) {
                console.info("[Component] XLSX ya estaba disponible globalmente.");
                return;
            }
            const sPath = sap.ui.require.toUrl("codelco/portaldescuentos/thirdparty/xlsx.safe.js");
            const oScript = document.createElement("script");
            oScript.src = sPath;
            oScript.async = false;
            oScript.onload = () => console.info("[Component] ✅ Librería XLSX cargada correctamente (CSP-safe).");
            oScript.onerror = (e) => console.error("[Component] ⚠️ Error cargando xlsx.safe.js", e);
            document.head.appendChild(oScript);
        }
    });
});