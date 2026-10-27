namespace ConversorPDF.XLL.VisTimeline.Domain.Models
{
    /// <summary>
    /// Representa un item en vis-timeline.
    /// Los nombres de propiedades son lowercase para coincidir exactamente
    /// con la API de vis-timeline al serializar a JSON.
    /// </summary>
    public class TimelineItem
    {
        /// <summary>Identificador único (IdSesion).</summary>
        public object id { get; set; }

        /// <summary>
        /// Texto visible dentro de la barra del item.
        /// Truncado a 40 chars para items estrechos.
        /// </summary>
        public string content { get; set; }

        /// <summary>
        /// Tooltip al hacer hover. Acepta HTML.
        /// </summary>
        public string title { get; set; }

        /// <summary>Fecha de inicio. Formato: "yyyy-MM-dd"</summary>
        public string start { get; set; }

        /// <summary>
        /// Fecha de fin EXCLUSIVA (FechaFinal + 1 día).
        /// Solo se incluye si type = "range".
        /// </summary>
        public string end { get; set; }

        /// <summary>Referencia al id de TimelineGroup (SedeCiudad).</summary>
        public string group { get; set; }

        /// <summary>"range" para multi-día, "box" para un solo día.</summary>
        public string type { get; set; }

        /// <summary>
        /// Clase CSS inyectada dinámicamente para colorear el item.
        /// Ej: "sede-cali", "sede-bogota"
        /// </summary>
        public string className { get; set; }
    }
}
