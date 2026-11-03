namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Nivel de confianza de la cédula detectada dentro del contenido de un PDF.
    /// Solo <see cref="Media"/> y <see cref="Alta"/> se consideran verificación válida.
    /// </summary>
    public enum ConfianzaCedula
    {
        /// <summary> No se halló cédula, o el PDF no tiene texto. </summary>
        Ninguna = 0,

        /// <summary> Heurística libre (número suelto de 6-10 dígitos). Nunca valida por sí sola. </summary>
        Baja = 1,

        /// <summary> Coincidencia exacta y única con una cédula de CONSULTAS. </summary>
        Media = 2,

        /// <summary> Cédula anclada a etiqueta ("C.C.", "Identificación") o al patrón cédula + correo. </summary>
        Alta = 3
    }
}
