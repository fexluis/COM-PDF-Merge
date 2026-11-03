using System;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace ConversorPDF.Tests.TestData
{
    /// <summary>
    /// Genera PDFs reales de una página para los tests de informes (sin depender de archivos
    /// almacenados en el repositorio).
    /// </summary>
    public static class PdfFixtureFactory
    {
        /// <summary>
        /// Crea un PDF de una página con las líneas indicadas.
        /// </summary>
        /// <param name="carpeta">Carpeta destino (se crea si no existe).</param>
        /// <param name="nombreArchivo">Nombre del archivo con extensión .pdf.</param>
        /// <param name="lineas">Líneas de texto a dibujar.</param>
        /// <returns>Ruta completa del archivo creado.</returns>
        public static string CrearPdf(string carpeta, string nombreArchivo, params string[] lineas)
        {
            if (string.IsNullOrWhiteSpace(carpeta)) throw new ArgumentNullException("carpeta");
            if (string.IsNullOrWhiteSpace(nombreArchivo)) throw new ArgumentNullException("nombreArchivo");

            Directory.CreateDirectory(carpeta);
            string ruta = Path.Combine(carpeta, nombreArchivo);

            using (var documento = new PdfDocument())
            {
                PdfPage pagina = documento.AddPage();

                using (XGraphics gfx = XGraphics.FromPdfPage(pagina))
                {
                    var fuente = new XFont("Arial", 12, XFontStyle.Regular);
                    double y = 60;

                    foreach (string linea in lineas ?? new string[0])
                    {
                        gfx.DrawString(linea, fuente, XBrushes.Black,
                            new XRect(40, y, pagina.Width.Point - 80, 20), XStringFormats.TopLeft);
                        y += 22;
                    }
                }

                documento.Save(ruta);
            }

            return ruta;
        }

        /// <summary>
        /// Crea una carpeta temporal única para un test.
        /// </summary>
        /// <param name="nombre">Prefijo descriptivo.</param>
        /// <returns>Ruta de la carpeta creada.</returns>
        public static string CrearCarpetaTemporal(string nombre)
        {
            string ruta = Path.Combine(Path.GetTempPath(), "InformesTests", nombre + "_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(ruta);
            return ruta;
        }

        /// <summary>
        /// Elimina una carpeta temporal ignorando errores.
        /// </summary>
        /// <param name="ruta">Carpeta a eliminar.</param>
        public static void EliminarCarpeta(string ruta)
        {
            try
            {
                if (!string.IsNullOrEmpty(ruta) && Directory.Exists(ruta)) Directory.Delete(ruta, true);
            }
            catch
            {
                // Los tests no deben fallar por limpieza.
            }
        }
    }
}
