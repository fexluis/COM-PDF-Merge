namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Objeto genérico para elementos de búsqueda (ComboBoxes).
    /// </summary>
    public class LookupItem
    {
        public int Id { get; set; }
        public string DisplayName { get; set; }
        public string Subunidad { get; set; }
        public string Cedula { get; set; }

        public override string ToString() => DisplayName;
    }
}
