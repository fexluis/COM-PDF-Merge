namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// DTO que encapsula los campos de datos leídos de una hoja de Excel tipo Orden de Pago Presupuestal.
    /// </summary>
    internal sealed class OrdenPagoDatosHoja
    {
        /// <summary>
        /// Nombre de la hoja de cálculo en Excel.
        /// </summary>
        public string NombreHoja { get; set; }

        /// <summary>
        /// Número de Orden de Pago leídos de la celda B15.
        /// </summary>
        public string Orden { get; set; }

        /// <summary>
        /// Fecha de pago leídos de la celda I15 o J15 (según la posición del rótulo).
        /// </summary>
        public string FechaPago { get; set; }

        /// <summary>
        /// Número de Obligación leídos de X16 o U16 (según la posición del rótulo).
        /// </summary>
        public string Obligacion { get; set; }

        /// <summary>
        /// Razón social / nombre del tercero leídos de J30 o I30 (según la posición del rótulo).
        /// </summary>
        public string Tercero { get; set; }

        /// <summary>
        /// Determina si la hoja contiene al menos un número de Orden de Pago válido.
        /// </summary>
        public bool EsValida
        {
            get { return !string.IsNullOrWhiteSpace(Orden); }
        }
    }
}
