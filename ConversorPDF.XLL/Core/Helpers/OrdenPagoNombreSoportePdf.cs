namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// DTO que almacena los datos extraídos del nombre base de un archivo PDF de soporte de Orden de Pago.
    /// Soporta patrones completos (con categoría) y patrones fallback (sin categoría).
    /// </summary>
    internal sealed class OrdenPagoNombreSoportePdf
    {
        /// <summary>
        /// Número de Orden de Pago extraído (mínimo 4 dígitos).
        /// </summary>
        public string Orden { get; set; }

        /// <summary>
        /// Número de Obligación extraído.
        /// </summary>
        public string Obligacion { get; set; }

        /// <summary>
        /// Código de categoría de 3 letras (ej. VIA, ARL, NOM) o null si no aplica/no está en catálogo.
        /// </summary>
        public string Categoria { get; set; }

        /// <summary>
        /// Razón social o descripción del tercero extraída del soporte.
        /// </summary>
        public string Tercero { get; set; }

        /// <summary>
        /// Indica si el soporte incluye una categoría válida del catálogo de 3 letras.
        /// </summary>
        public bool TieneCategoria
        {
            get { return !string.IsNullOrEmpty(Categoria); }
        }
    }
}
