namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Representa un elemento canónico o personalizado del catálogo de terminales aéreas en memoria.
    /// Utilizado para el cálculo y resolución de topes de transporte aéreo (maxvalor).
    /// </summary>
    public class TerminalCatalogoItem
    {
        /// <summary>
        /// Identificador único de la terminal en el catálogo SQLite.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Nombre descriptivo canónico de la terminal o aeropuerto (ej. 'ANTIOQUIA (Rionegro)', 'OTRO').
        /// </summary>
        public string Terminal { get; set; }

        /// <summary>
        /// Valor monetario máximo autorizado para la terminal aérea (tope oficial).
        /// </summary>
        public decimal MaxValor { get; set; }
    }
}
