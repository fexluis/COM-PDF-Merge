using System;
using System.IO;
using System.Reflection;

namespace ConversorPDF.XLL.VisTimeline.Infrastructure.Helpers
{
    public static class ResourceExtractor
    {
        /// <summary>
        /// Extrae un recurso embebido del assembly al disco si aún no existe.
        /// Devuelve la ruta absoluta del archivo extraído.
        /// </summary>
        public static string ExtractToDocuments(string resourceName, string outputDir, string outputFileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Nota: Podríamos verificar integridad con hash, pero por ahora File.Exists es suficiente caché.
            if (!File.Exists(outputPath))
            {
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                        throw new InvalidOperationException(
                            $"Recurso embebido no encontrado: {resourceName}. " +
                            "Verifique que el archivo tenga acción de compilación 'EmbeddedResource'.");

                    Directory.CreateDirectory(outputDir);
                    using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fs);
                    }
                }
            }

            return outputPath;
        }
    }
}
