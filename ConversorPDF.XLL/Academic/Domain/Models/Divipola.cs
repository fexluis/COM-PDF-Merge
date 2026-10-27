namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa un registro de la División Político-Administrativa (Divipola).
    /// </summary>
    public class Divipola
    {
        /// <summary> Identificador único de la división política. </summary>
        public int Id { get; set; }
        /// <summary> Nombre del departamento. </summary>
        public string Departamento { get; set; }
        /// <summary> Código DANE del departamento. </summary>
        public string DepartamentoCodigo { get; set; }
        /// <summary> Nombre normalizado del departamento para búsquedas. </summary>
        public string DepartamentoNormalizado { get; set; }
        /// <summary> Nombre del municipio. </summary>
        public string Municipio { get; set; }
        /// <summary> Código DANE del municipio. </summary>
        public string MunicipioCodigo { get; set; }
        /// <summary> Nombre normalizado del municipio para búsquedas. </summary>
        public string MunicipioNormalizado { get; set; }
        /// <summary> Categoría del municipio según clasificación nacional. </summary>
        public int MunicipioCategoria { get; set; }
    }
}
