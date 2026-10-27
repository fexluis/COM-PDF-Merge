namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class GastoTerrestreEntity
    {
        public int Id { get; set; }
        public string Territorial { get; set; }
        public string CodigoTrayecto { get; set; }
        public string Trayecto { get; set; }
        public decimal? Valor { get; set; }
        public string Subunidad { get; set; }
    }
}
