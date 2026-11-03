namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Valores canónicos de las columnas ESTADO (C) e INFORME (H) de la hoja CONSULTAS.
    /// </summary>
    public static class EstadoComision
    {
        /// <summary> Informe ya verificado: no se consulta disco ni se modifica la celda. </summary>
        public const string Verificado = "VERIFICADO";

        /// <summary> Reintegro registrado: se protege igual que VERIFICADO. </summary>
        public const string Reintegro = "REINTEGRO";

        /// <summary> Comisión revocada. </summary>
        public const string Revocada = "REVOCADA";

        /// <summary> Comisión rechazada. </summary>
        public const string Rechazada = "RECHAZADA";

        /// <summary> Comisión anulada. </summary>
        public const string Anulada = "ANULADA";

        /// <summary> Marca de ausencia de soporte físico de la solicitud. </summary>
        public const string SinSolicitud = "SIN SOLICITUD";

        /// <summary> Variante heredada de la marca de ausencia de soporte. </summary>
        public const string FaltaSolicitud = "FALTA SOLICITUD";

        /// <summary> Hay informe (INF) pero falta el certificado de cumplimiento (CUM). </summary>
        public const string FaltaCum = "FALTA CUM";

        /// <summary> Hay certificado (CUM) pero falta el informe (INF). </summary>
        public const string FaltaInf = "FALTA INF";

        /// <summary> El soporte no pudo validarse (cédula no coincide, archivo sin texto o ERR). </summary>
        public const string Error = "ERROR";

        /// <summary> No se encontró ningún soporte de informe para la solicitud. </summary>
        public const string SinInforme = "SIN INFORME";

        /// <summary> Pareja INF + CUM completa (estado interno de la auditoría, no se escribe). </summary>
        public const string ParOk = "PAR OK";
    }
}
