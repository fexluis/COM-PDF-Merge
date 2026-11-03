namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Información de un PDF de informes ya interpretado, sea a partir de su
    /// nombre estandarizado o del resultado del renombrado.
    /// </summary>
    public class InformePdf
    {
        /// <summary> Número de solicitud / comisión. </summary>
        public string NumeroSolicitud { get; set; }

        /// <summary> Cédula usada en el nombre (solo dígitos). </summary>
        public string Cedula { get; set; }

        /// <summary> Tipo del informe (ver <see cref="InformeTipo"/>). </summary>
        public string Tipo { get; set; }

        /// <summary> Nombre del comisionado tal como quedó en el archivo (máx. 12 caracteres). </summary>
        public string NombreComisionado { get; set; }

        /// <summary> Ruta completa del archivo en disco. </summary>
        public string RutaOrigen { get; set; }

        /// <summary> Nombre del archivo con extensión. </summary>
        public string NombreArchivo { get; set; }

        /// <summary> Cantidad de páginas (0 si no se consultó). </summary>
        public int Paginas { get; set; }

        /// <summary> True si el tipo participa en la formación de parejas INF + CUM. </summary>
        public bool EsParValido
        {
            get { return InformeTipo.EsParValido(Tipo); }
        }
    }
}
