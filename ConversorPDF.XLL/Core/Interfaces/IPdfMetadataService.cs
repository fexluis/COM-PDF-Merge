using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de gestión de metadatos en documentos PDF.
    /// </summary>
    public interface IPdfMetadataService
    {
        /// <summary>
        /// Guarda los metadatos proporcionados en el archivo PDF especificado.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF.</param>
        /// <param name="metadata">Objeto con la información de metadatos.</param>
        void GuardarMetadatos(string rutaPdf, Metadata metadata);

        /// <summary>
        /// Lee los metadatos de un archivo PDF existente.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF.</param>
        /// <returns>Objeto Metadata con la información extraída.</returns>
        Metadata LeerMetadatos(string rutaPdf);

        /// <summary>
        /// Obtiene un resumen formateado de los metadatos del PDF.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF.</param>
        /// <returns>String con la información legible de los metadatos.</returns>
        string ObtenerResumenMetadatos(string rutaPdf);
        
        /// <summary>
        /// Actualiza un campo específico de metadatos en un archivo PDF.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF.</param>
        /// <param name="campo">Nombre del campo a actualizar (ej: "estado", "titulo", "autor").</param>
        /// <param name="valor">Nuevo valor para el campo.</param>
        void ActualizarMetadato(string rutaPdf, string campo, string valor);
    }
}
