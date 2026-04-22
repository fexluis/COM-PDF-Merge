using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Advanced;
using System;
using System.IO;
using System.Text;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio para la gestión de metadatos en documentos PDF utilizando la librería PDFSharp.
    /// </summary>
    public class PdfMetadataService : IPdfMetadataService
    {
        // Nombres de campos con prefijo "meta" (Constantes internas del servicio)
        private const string META_FECHA = "/metaFecha";
        private const string META_DOC = "/metaDocumento";
        private const string META_CDP = "/metaCDP";
        private const string META_RP = "/metaRP";
        private const string META_OB = "/metaOB";
        private const string META_OP = "/metaOP";
        private const string META_TIPO = "/metaTipo";
        private const string META_COMISION = "/metaComision";

        /// <summary>
        /// Guarda el objeto <see cref="Metadata"/> en un archivo PDF mediante el uso de campos personalizados.
        /// Si el archivo no existe, crea uno nuevo y le inserta los datos.
        /// </summary>
        /// <param name="rutaPdf">Ruta física del archivo PDF.</param>
        /// <param name="metadata">Información a inyectar.</param>
        public void GuardarMetadatos(string rutaPdf, Metadata metadata)
        {
            if (string.IsNullOrEmpty(rutaPdf)) throw new ArgumentNullException(nameof(rutaPdf));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));

            PdfDocument document;

            if (File.Exists(rutaPdf))
            {
                document = PdfReader.Open(rutaPdf, PdfDocumentOpenMode.Modify);
            }
            else
            {
                document = new PdfDocument();
                document.AddPage();
            }

            try
            {
                // Metadatos estándar (Info Dictionary)
                document.Info.Title = metadata.Titulo ?? string.Empty;
                document.Info.Author = metadata.Autor ?? string.Empty;
                document.Info.Subject = metadata.Asunto ?? string.Empty;
                document.Info.Keywords = metadata.PalabrasClave ?? string.Empty;

                // Guardar campos personalizados directamente en el diccionario Info (Más estable en PDFSharp 1.50)
                document.Info.Elements.SetString(META_FECHA, metadata.Fecha?.ToString("yyyy-MM-dd") ?? string.Empty);
                document.Info.Elements.SetString(META_DOC, metadata.MetaDocumento?.ToString() ?? string.Empty);
                document.Info.Elements.SetString(META_CDP, metadata.CDP ?? string.Empty);
                document.Info.Elements.SetString(META_RP, metadata.RP ?? string.Empty);
                document.Info.Elements.SetString(META_OB, metadata.OB ?? string.Empty);
                document.Info.Elements.SetString(META_OP, metadata.OP ?? string.Empty);
                document.Info.Elements.SetString(META_TIPO, metadata.Tipo ?? string.Empty);
                document.Info.Elements.SetString(META_COMISION, metadata.Comision ?? string.Empty);

                document.Save(rutaPdf);
            }
            finally
            {
                document.Close();
                document.Dispose();
            }
        }

        /// <summary>
        /// Recupera un objeto <see cref="Metadata"/> desde el diccionario personalizado de un PDF.
        /// </summary>
        /// <param name="rutaPdf">Ruta física del archivo PDF.</param>
        /// <returns>Objeto con la información extraída.</returns>
        public Metadata LeerMetadatos(string rutaPdf)
        {
            if (!File.Exists(rutaPdf))
                throw new FileNotFoundException("Archivo PDF no encontrado", rutaPdf);

            using (var document = PdfReader.Open(rutaPdf, PdfDocumentOpenMode.ReadOnly))
            {
                var metadata = new Metadata
                {
                    Titulo = document.Info.Title,
                    Autor = document.Info.Author,
                    Asunto = document.Info.Subject,
                    PalabrasClave = document.Info.Keywords
                };

                // Recuperar campos personalizados directamente del diccionario Info
                metadata.Fecha = ParseFecha(document.Info.Elements.GetString(META_FECHA));
                metadata.MetaDocumento = ParseMetaDocumento(document.Info.Elements.GetString(META_DOC));
                metadata.CDP = document.Info.Elements.GetString(META_CDP);
                metadata.RP = document.Info.Elements.GetString(META_RP);
                metadata.OB = document.Info.Elements.GetString(META_OB);
                metadata.OP = document.Info.Elements.GetString(META_OP);
                metadata.Tipo = document.Info.Elements.GetString(META_TIPO);
                metadata.Comision = document.Info.Elements.GetString(META_COMISION);

                return metadata;
            }
        }

        /// <summary>
        /// Genera una cadena legible con todos los metadatos almacenados para propósitos de auditoría o visualización.
        /// </summary>
        /// <param name="rutaPdf">Ruta física del archivo PDF.</param>
        /// <returns>Texto formateado.</returns>
        public string ObtenerResumenMetadatos(string rutaPdf)
        {
            var m = LeerMetadatos(rutaPdf);
            var sb = new StringBuilder();

            sb.AppendLine("📄 METADATOS DEL DOCUMENTO");
            sb.AppendLine("═══════════════════════════════════════");
            sb.AppendLine($"Clave:           {m.Titulo}");
            sb.AppendLine($"Autor:           {m.Autor}");
            sb.AppendLine($"Asunto:          {m.Asunto}");
            sb.AppendLine($"Palabras Clave:  {m.PalabrasClave}");
            sb.AppendLine("───────────────────────────────────────");
            sb.AppendLine($"metaFecha:       {(m.Fecha.HasValue ? m.Fecha.Value.ToString("yyyy-MM-dd") : "N/A")}");
            sb.AppendLine($"metaDocumento:   {m.MetaDocumento?.ToString() ?? "N/A"}");
            sb.AppendLine($"metaCDP:         {m.CDP ?? "N/A"}");
            sb.AppendLine($"metaRP:          {m.RP ?? "N/A"}");
            sb.AppendLine($"metaOB:          {m.OB ?? "N/A"}");
            sb.AppendLine($"metaOP:          {m.OP ?? "N/A"}");
            sb.AppendLine($"metaTipo:        {m.Tipo ?? "N/A"}");
            sb.AppendLine($"metaComision:    {m.Comision ?? "N/A"}");
            sb.AppendLine("═══════════════════════════════════════");

            return sb.ToString();
        }

        /// <summary>
        /// Actualiza un único campo de metadato sin necesidad de reconstruir manualmente todo el objeto.
        /// </summary>
        /// <param name="rutaPdf">Ruta física del archivo PDF.</param>
        /// <param name="campo">Identificador del campo (sensible a mayúsculas/minúsculas vía switch).</param>
        /// <param name="valor">Nuevo valor en formato texto.</param>
        public void ActualizarMetadato(string rutaPdf, string campo, string valor)
        {
            if (string.IsNullOrEmpty(rutaPdf)) throw new ArgumentNullException(nameof(rutaPdf));
            if (string.IsNullOrEmpty(campo)) throw new ArgumentNullException(nameof(campo));

            var metadata = LeerMetadatos(rutaPdf);
            bool modificado = true;

            switch (campo.ToLower())
            {
                case "estado":
                case "metadocumento":
                    metadata.MetaDocumento = ParseMetaDocumento(valor);
                    break;
                case "titulo":
                    metadata.Titulo = valor;
                    break;
                case "autor":
                    metadata.Autor = valor;
                    break;
                case "asunto":
                    metadata.Asunto = valor;
                    break;
                case "palabrasclave":
                    metadata.PalabrasClave = valor;
                    break;
                case "cdp":
                    metadata.CDP = valor;
                    break;
                case "rp":
                    metadata.RP = valor;
                    break;
                case "ob":
                    metadata.OB = valor;
                    break;
                case "op":
                    metadata.OP = valor;
                    break;
                case "tipo":
                    metadata.Tipo = valor;
                    break;
                case "comision":
                    metadata.Comision = valor;
                    break;
                case "fecha":
                    metadata.Fecha = ParseFecha(valor);
                    break;
                default:
                    modificado = false;
                    break;
            }

            if (modificado)
            {
                GuardarMetadatos(rutaPdf, metadata);
            }
        }

        private static TipoMetadata? ParseMetaDocumento(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return null;
            
            // Reemplazar guiones por guiones bajos para compatibilidad con el Enum
            valor = valor.Replace("-", "_");

            if (Enum.TryParse<TipoMetadata>(valor, true, out var resultado))
                return resultado;

            return null;
        }

        private static DateTime? ParseFecha(string fechaStr) =>
            DateTime.TryParse(fechaStr, out var fecha) ? fecha : (DateTime?)null;
    }
}
