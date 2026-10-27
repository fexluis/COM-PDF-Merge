namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Permiso
    {
        public string Perfil { get; set; }
        /// <summary>
        /// Lista de IDs de botones permitidos, separados por comas.
        /// Ejemplo: "ConvertirPDF,GenerarFilaSeleccionada"
        /// </summary>
        public string Botones { get; set; }
    }
}
