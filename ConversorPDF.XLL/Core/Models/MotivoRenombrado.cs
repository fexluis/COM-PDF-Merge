namespace ConversorPDF.XLL.Core.Models
{
    /// <summary> Motivos posibles del resultado del renombrado de un PDF. </summary>
    public static class MotivoRenombrado
    {
        /// <summary> El archivo fue renombrado con éxito. </summary>
        public const string Renombrado = "RENOMBRADO";

        /// <summary> El archivo ya posee el nombre estandarizado correcto. </summary>
        public const string YaCorrecto = "YA_CORRECTO";

        /// <summary> El archivo destino ya existe en la carpeta y no se sobrescribe. </summary>
        public const string Colision = "COLISION";

        /// <summary> El archivo destino ya existía con contenido idéntico; el origen se archivó en _Duplicados. </summary>
        public const string Duplicado = "DUPLICADO";

        /// <summary> Comisión con múltiples candidatos sin resolución concluyente por metadatos ni OCR. </summary>
        public const string Ambigua = "AMBIGUA";

        /// <summary> El número de comisión no se encuentra en la hoja CONSULTAS dentro del rango. </summary>
        public const string SinComision = "SIN_COMISION_EN_CONSULTAS";

        /// <summary> La comisión existe en CONSULTAS pero carece de cédula o nombre completos. </summary>
        public const string SinDatos = "SIN_DATOS_EN_CONSULTAS";

        /// <summary> El nombre del archivo no inicia con un prefijo numérico de comisión válido. </summary>
        public const string PrefijoInvalido = "PREFIJO_INVALIDO";

        /// <summary> Ocurrió una excepción de entrada/salida al intentar mover el archivo. </summary>
        public const string ErrorIo = "ERROR_IO";
    }
}

