sap.ui.define([
    "sap/ui/core/UIComponent",
    "sap/ui/model/json/JSONModel",
    "codelco/portaldescuentos/model/models",
    "codelco/portaldescuentos/model/periodo"
], (UIComponent, JSONModel, models, periodo) => {
    "use strict";

    // Clave del almacenamiento local del navegador (solo para pruebas con mock)
    const LS_CARGAS = "codelco.portaldescuentos.cargas";

    return UIComponent.extend("codelco.portaldescuentos.Component", {
        metadata: {
            manifest: "json",
            interfaces: ["sap.ui.core.IAsyncContentCreation"]
        },

        init() {
            // 🔹 Cargar XLSX solo si no existe (mismo esquema que Carrefour)
            this._cargarXLSX();

            UIComponent.prototype.init.apply(this, arguments);

            this.setModel(models.createDeviceModel(), "device");

            // Modelo "app": estado calculado (no se guarda, se recalcula)
            this.setModel(new JSONModel({}), "app");

            // Cuando el mock termina de cargar: recuperar cargas guardadas y calcular período
            const oMock = this.getModel("mock");
            oMock.dataLoaded().then(() => {
                oMock.setProperty("/cargas", this._leerCargasGuardadas());
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

            const oHoy = this.getFechaHoy();
            const oPeriodo = periodo.calcular(oHoy, oParams.diaCierre);

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

        getFechaHoy() {
            const sFecha = new URLSearchParams(window.location.search).get("fecha");
            if (sFecha) {
                const oFecha = new Date(sFecha + "T" + new Date().toTimeString().slice(0, 8));
                if (!isNaN(oFecha)) {
                    return oFecha;
                }
            }
            return new Date();
        },

        // ================= Cargas (mock + localStorage) =================

        /** Agrega una carga nueva y la persiste en el navegador */
        agregarCarga(oCarga) {
            const oMock = this.getModel("mock");
            const aCargas = oMock.getProperty("/cargas") || [];
            aCargas.push(oCarga);
            oMock.setProperty("/cargas", aCargas);
            this._guardarCargas();
            this.recalcularPeriodo();
        },

        /**
         * 🧪 Simula el proceso masivo en SAP sobre la carga del período actual:
         * - Observados  → Rechazado (mismo motivo)
         * - Conformes   → Cargado, salvo RUT no existe / no activo según el mock
         */
        simularProcesoSAP() {
            const oMock = this.getModel("mock");
            const oCarga = this.getModel("app").getProperty("/cargaActual");
            if (!oCarga || oCarga.procesada) {
                return false;
            }
            const aVigentes = oMock.getProperty("/trabajadores/vigentes") || [];
            const aInactivos = oMock.getProperty("/trabajadores/inactivos") || [];

            oCarga.detalle.forEach(d => {
                if (d.estado === "Observado") {
                    d.estado = "Rechazado";
                } else if (aInactivos.includes(d.rut)) {
                    d.estado = "Rechazado";
                    d.mensaje = "Rut no activo";
                } else if (!aVigentes.includes(d.rut)) {
                    d.estado = "Rechazado";
                    d.mensaje = "Rut no existe";
                } else {
                    d.estado = "Cargado";
                    d.mensaje = "Cargado en IT0015";
                }
            });

            oCarga.procesada = true;
            oCarga.fechaProceso = this.formatearFechaHora(this.getFechaHoy());
            oCarga.totalCargados = oCarga.detalle.filter(d => d.estado === "Cargado").length;
            oCarga.totalRechazados = oCarga.detalle.filter(d => d.estado === "Rechazado").length;
            oCarga.estado = oCarga.totalRechazados > 0 ? "Procesado con rechazos" : "Procesado";
            oCarga.estadoState = oCarga.totalRechazados > 0 ? "Warning" : "Success";

            this._guardarCargas();
            this.recalcularPeriodo();
            return true;
        },

        /** 🧪 Borra todas las cargas de prueba */
        limpiarCargas() {
            this.getModel("mock").setProperty("/cargas", []);
            try { window.localStorage.removeItem(LS_CARGAS); } catch (e) { /* sin storage */ }
            this.recalcularPeriodo();
        },

        formatearFechaHora(oFecha) {
            const p = n => String(n).padStart(2, "0");
            return `${p(oFecha.getDate())}/${p(oFecha.getMonth() + 1)}/${oFecha.getFullYear()} ` +
                   `${p(oFecha.getHours())}:${p(oFecha.getMinutes())}`;
        },

        _guardarCargas() {
            try {
                window.localStorage.setItem(LS_CARGAS,
                    JSON.stringify(this.getModel("mock").getProperty("/cargas") || []));
            } catch (e) { /* storage no disponible: solo queda en memoria */ }
        },

        _leerCargasGuardadas() {
            try {
                return JSON.parse(window.localStorage.getItem(LS_CARGAS)) || [];
            } catch (e) {
                return [];
            }
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