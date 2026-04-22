namespace ConversorPDF.XLL.Core.Models.Academic
{
    public class Asignatura
    {
        public int Id { get; set; }
        public int IdPrograma { get; set; }
        public string Semestre { get; set; }
        public string Código { get; set; }
        public string Abreviatura { get; set; }
        public string Área { get; set; }
        public string NombreAsignatura { get; set; }
        public int Creditos { get; set; }
        public string TipoAsignatura { get; set; }
        public string TipoCredito { get; set; }
        public string Prerequisito { get; set; }
    }
}
