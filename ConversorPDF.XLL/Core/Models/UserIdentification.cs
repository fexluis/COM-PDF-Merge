namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Representa los datos de identificación de un usuario para ser insertados en documentos Excel.
    /// </summary>
    public class UserIdentification
    {
        public string Initials { get; }
        public string FullName { get; }

        public UserIdentification(string initials, string fullName)
        {
            Initials = initials;
            FullName = fullName;
        }
    }
}
