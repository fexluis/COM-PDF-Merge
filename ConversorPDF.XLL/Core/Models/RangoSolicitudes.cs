namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Rango de comisiones detectado en la columna B de la hoja CONSULTAS.
    /// </summary>
    public class RangoSolicitudes
    {
        /// <summary> Primer registro válido de la columna B. </summary>
        public int Inicio { get; set; }

        /// <summary> Último registro válido de la columna B. </summary>
        public int Fin { get; set; }

        /// <summary> True si se encontró al menos un registro válido. </summary>
        public bool HayDatos { get; set; }
    }
}
