namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Programa
    {
        public int Id { get; set; }
        public string NombrePrograma { get; set; }
        public string Nombre { get => NombrePrograma; set => NombrePrograma = value; }
        public string Abreviatura { get; set; }
        public string CodPrograma { get => Abreviatura; set => Abreviatura = value; }
        public int Pensum { get; set; }
        public int Actual { get; set; }
        public int Activo { get; set; }
        public string Modalidad { get; set; }
    }
}
