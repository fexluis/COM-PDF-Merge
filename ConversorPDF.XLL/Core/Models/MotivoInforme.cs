namespace ConversorPDF.XLL.Core.Models
{
    /// <summary> Motivos posibles del resultado del proceso de búsqueda de informes. </summary>
    public static class MotivoInforme
    {
        /// <summary> El archivo ya posee el nombre estandarizado correcto. </summary>
        public const string YaCorrecto = "YA_CORRECTO";

        /// <summary> El archivo fue renombrado con éxito. </summary>
        public const string Renombrado = "RENOMBRADO";

        /// <summary> El archivo fue renombrado solo en simulación (sin tocar disco). </summary>
        public const string Simulado = "SIMULADO";

        /// <summary> El nombre del archivo no inicia con un número de solicitud válido. </summary>
        public const string PrefijoInvalido = "PREFIJO_INVALIDO";

        /// <summary> La solicitud no figura en CONSULTAS dentro del rango solicitado. </summary>
        public const string SinComision = "SIN_COMISION_EN_CONSULTAS";

        /// <summary> La solicitud existe en CONSULTAS pero sin cédula o sin nombre. </summary>
        public const string SinDatos = "SIN_DATOS_EN_CONSULTAS";

        /// <summary> El PDF no tiene capa de texto (escaneado) o no pudo abrirse. </summary>
        public const string SinTexto = "SIN_TEXTO";

        /// <summary> No se pudo determinar el tipo (no es INF ni CUM) o se detectaron ambos. </summary>
        public const string SinTipo = "SIN_TIPO";

        /// <summary> La cédula hallada en el PDF no coincide con la de CONSULTAS ni con la del nombre. </summary>
        public const string CedulaNoCoincide = "CEDULA_NO_COINCIDE";

        /// <summary> La cédula esperada no aparece en el contenido del PDF. </summary>
        public const string CedulaNoHallada = "CEDULA_NO_HALLADA";

        /// <summary> El PDF contiene más de una cédula candidata distinta (ambiguo). </summary>
        public const string CedulaAmbigua = "CEDULA_AMBIGUA";

        /// <summary> El destino existía con contenido distinto; se generó un nombre con sufijo. </summary>
        public const string Colision = "COLISION";

        /// <summary> El destino existía con contenido idéntico; el origen se archivó en _Duplicados. </summary>
        public const string Duplicado = "DUPLICADO";

        /// <summary> Pareja INF + CUM unificada en un archivo UNI. </summary>
        public const string Unificado = "UNIFICADO";

        /// <summary> Archivo INF o CUM sin pareja (falta el complemento). </summary>
        public const string SinPareja = "SIN_PAREJA";

        /// <summary> La unificación falló; los originales se conservaron. </summary>
        public const string ErrorMerge = "ERROR_MERGE";

        /// <summary> Error de entrada/salida al mover o renombrar. </summary>
        public const string ErrorIo = "ERROR_IO";
    }
}
