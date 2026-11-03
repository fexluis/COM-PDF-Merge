namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Fila relevante de la hoja CONSULTAS para la búsqueda de informes:
    /// número de solicitud, cédula y nombre del comisionado.
    /// </summary>
    public class InformeSolicitud
    {
        /// <summary> Número de solicitud / comisión (columna B). </summary>
        public string Numero { get; set; }

        /// <summary> Cédula del comisionado en solo dígitos (columna J). </summary>
        public string Cedula { get; set; }

        /// <summary> Nombre completo del comisionado tal como figura en CONSULTAS (columna K). </summary>
        public string Nombre { get; set; }
    }
}
