using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Usuario
    {
        public string Id { get; set; }
        public string Nombre { get; set; }
        public string NombreUsuario { get; set; }
        public string Subunidad { get; set; }
        public string Perfil { get; set; }
        public int Activo { get; set; } = 1;
        /// <summary> Cadena Base64 con la imagen de la firma del usuario, almacenada directamente en la columna firma de la tabla usuarios. </summary>
        public string Firma { get; set; }
    }
}
