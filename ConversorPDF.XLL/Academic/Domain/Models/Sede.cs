namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Sede
    {
        public int Id { get; set; }
        public int? IdDivipola { get; set; }
        public string Departamento { get; set; }
        public string Ciudad { get; set; }
        public string DepartamentoSiif { get; set; }
        public string CiudadSiif { get; set; }
        public string Subsede { get; set; }
        public int Categoria { get; set; }
        public decimal Transporte { get; set; }
        
        /// <summary>
        /// Indica si la sede está activa (1) o inactiva (0).
        /// </summary>
        public int Activo { get; set; } = 1;
        
        public string Subunidad { get; set; }

        /// <summary>
        /// Nombre amigable de la sede combinando Ciudad y Subsede si aplica.
        /// </summary>
        public string Nombre => string.IsNullOrWhiteSpace(Subsede) ? (Ciudad ?? string.Empty) : $"{Ciudad} ({Subsede})";
    }
}
