namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Asignatura
    {
        public int Id { get; set; }
        public int IdPrograma { get; set; }
        public string Semestre { get; set; }
        public int NumNivel { get; set; }
        public string CodUnidad { get; set; }
        public string CodMateria { get; set; }
        public int UniTeorica { get; set; }
        public string Metodologia { get; set; }
        public string Abreviatura { get; set; }
        public string Nucleo { get; set; }
        public string Nombre { get; set; }
        public int Creditos { get; set; }
        public int IntHoraria { get; set; }
        public string TipoAsignatura { get; set; }
        public string TipoCredito { get; set; }
        public string Prerequisito { get; set; }

        // Alias y propiedades de conveniencia para compatibilidad con el AG y consultas analíticas
        public int Horas { get => IntHoraria; set => IntHoraria = value; }
        public double HorasSemanales { get; set; } = 4.0;
        public int IdSede { get; set; }
        public string Codigo { get => CodMateria; set => CodMateria = value; }
        public string Modalidad { get; set; } = "Presencial";
        public string Subunidad { get; set; } = string.Empty;
    }
}
