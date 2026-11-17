sap.ui.define([], () => {
    "use strict";

    const pad = n => String(n).padStart(2, "0");
    const fmt = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;

    return {
        /**
         * Período = mes y año de la fecha "hoy".
         * Ventana de carga: día 1 hasta el día de cierre (ej. 15) a las 23:59.
         */
        calcular(oHoy, iDiaCierre) {
            const iMes = oHoy.getMonth() + 1;
            const iAnio = oHoy.getFullYear();

            const oApertura = new Date(iAnio, iMes - 1, 1, 0, 0, 0);
            const oCierre = new Date(iAnio, iMes - 1, iDiaCierre, 23, 59, 59);
            const bAbierto = oHoy >= oApertura && oHoy <= oCierre;

            const iMsDia = 24 * 60 * 60 * 1000;
            const iDiasRestantes = bAbierto ? Math.ceil((oCierre - oHoy) / iMsDia) : 0;
            const iTotal = oCierre - oApertura;
            const iAvance = Math.min(100, Math.max(0, Math.round(((oHoy - oApertura) / iTotal) * 100)));

            // Próximo período (para el mensaje cuando está cerrado)
            const oProxApertura = new Date(iAnio, iMes, 1);
            const oProxCierre = new Date(oProxApertura.getFullYear(), oProxApertura.getMonth(), iDiaCierre);

            return {
                mes: iMes,
                anio: iAnio,
                texto: `${pad(iMes)}/${iAnio}`,
                apertura: fmt(oApertura),
                cierre: fmt(oCierre),
                abierto: bAbierto,
                estadoTexto: bAbierto ? "Abierto" : "Cerrado",
                estadoState: bAbierto ? "Success" : "Error",
                diasRestantes: iDiasRestantes,
                avance: iAvance,
                proximaApertura: fmt(oProxApertura),
                proximoCierre: fmt(oProxCierre)
            };
        }
    };
});
