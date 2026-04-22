namespace ConversorPDF.XLL.Core.Models.Academic
{
    public class Contenido
    {
        public int Id { get; set; }
        public int IdMicrocurriculo { get; set; }
        public string Unidad { get; set; }
        public string ContenidoTexto { get; set; } // Renamed to avoid collision with class name
    }
}
