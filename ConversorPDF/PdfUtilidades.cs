using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace ConversorPDF
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F1E2D3C4-B5A6-9E8D-7C6B-5A4B3C2D1E0F")]
    [ProgId("ConversorPDF.PdfUtilidades")]
    public class PdfUtilidades : IPdfUtilidades
    {
        public string ExtraerTodoTexto(string rutaPdf)
        {
            if (!System.IO.File.Exists(rutaPdf))
                throw new ArgumentException($"Archivo no encontrado: {rutaPdf}");

            var sb = new StringBuilder();

            using (PdfDocument document = PdfDocument.Open(rutaPdf))
            {
                foreach (Page page in document.GetPages())
                {
                    // Mejor opción: usa ContentOrderTextExtractor para orden lógico + newlines 
                    string textoPagina = ContentOrderTextExtractor.GetText(page);
                    sb.AppendLine(textoPagina);
                    sb.AppendLine("\r\n--- Fin de página ---\r\n");
                }
            }

            return sb.ToString();
        }

        public string ExtraerTextoPagina(string rutaPdf, int numeroPagina)
        {
            if (!System.IO.File.Exists(rutaPdf))
                throw new ArgumentException($"Archivo no encontrado: {rutaPdf}");

            using (PdfDocument document = PdfDocument.Open(rutaPdf))
            {
                if (numeroPagina < 1 || numeroPagina > document.NumberOfPages)
                    throw new ArgumentException("Número de página inválido");

                Page page = document.GetPage(numeroPagina); // 1-based en PdfPig 
                return ContentOrderTextExtractor.GetText(page);
            }
        }

        public string ExtraerPalabras(string rutaPdf, int numeroPagina)
        {
            if (!System.IO.File.Exists(rutaPdf))
                throw new ArgumentException($"Archivo no encontrado: {rutaPdf}");

            var sb = new StringBuilder();

            using (PdfDocument document = PdfDocument.Open(rutaPdf))
            {
                if (numeroPagina < 1 || numeroPagina > document.NumberOfPages)
                    throw new ArgumentException("Número de página inválido");

                Page page = document.GetPage(numeroPagina);
                
                // Usar NearestNeighbourWordExtractor para obtener palabras
                IEnumerable<Word> words = page.GetWords(NearestNeighbourWordExtractor.Instance);

                sb.AppendLine($"Página {numeroPagina} - {words.Count()} palabras encontradas:");
                sb.AppendLine("----------------------------------------");

                foreach (Word word in words)
                {
                    sb.AppendLine($"Palabra: '{word.Text}' | " 
                                + $"Pos: Left={word.BoundingBox.Left:F1}, " 
                                + $"Bottom={word.BoundingBox.Bottom:F1}, " 
                                + $"Width={word.BoundingBox.Width:F1}, " 
                                + $"Height={word.BoundingBox.Height:F1}");
                }
            }

            return sb.ToString();
        }

        public string ExtraerTextoEnRectangulo(string rutaPdf, int numeroPagina, double left, double bottom, double width, double height)
        {
            if (!System.IO.File.Exists(rutaPdf))
                throw new ArgumentException($"Archivo no encontrado: {rutaPdf}");

            var sb = new StringBuilder();

            using (PdfDocument document = PdfDocument.Open(rutaPdf))
            {
                if (numeroPagina < 1 || numeroPagina > document.NumberOfPages)
                    throw new ArgumentException("Número de página inválido");

                Page page = document.GetPage(numeroPagina);
                
                // Definir los límites del rectángulo (coordenadas PDF: Y crece hacia arriba)
                // PdfRectangle constructor: left, bottom, right, top
                var rect = new PdfRectangle(left, bottom, left + width, bottom + height);

                sb.AppendLine($"Página {numeroPagina} - Texto en región [L:{left:F1}, B:{bottom:F1}, W:{width:F1}, H:{height:F1}]:");
                sb.AppendLine("----------------------------------------");

                // Filtrar letras que caen dentro del rectángulo
                // Usamos StartBaseLine (punto de referencia) o GlyphRectangle (caja visual)
                var letrasEnRegion = page.Letters.Where(l => 
                {
                    // Check if StartBaseLine is inside rect
                    bool inside = l.StartBaseLine.X >= rect.Left && l.StartBaseLine.X <= rect.Right &&
                                  l.StartBaseLine.Y >= rect.Bottom && l.StartBaseLine.Y <= rect.Top;
                    
                    // Check intersection with GlyphRectangle
                    // Manual intersection check since PdfRectangle might not have Intersects in this version
                    bool intersects = !(l.GlyphRectangle.Left > rect.Right || 
                                      l.GlyphRectangle.Right < rect.Left || 
                                      l.GlyphRectangle.Top < rect.Bottom || 
                                      l.GlyphRectangle.Bottom > rect.Top);

                    return inside || intersects;
                }).OrderByDescending(l => l.StartBaseLine.Y) // Ordenar de arriba a abajo
                 .ThenBy(l => l.StartBaseLine.X);           // Luego de izquierda a derecha

                // Reconstruir el texto agrupando por líneas (simple grouping por Y)
                var letras = letrasEnRegion.ToList();
                if (letras.Count > 0)
                {
                    double lastY = letras[0].StartBaseLine.Y;

                    foreach (var letter in letras)
                    {
                        // Detectar nueva línea aproximada
                        if (Math.Abs(letter.StartBaseLine.Y - lastY) > letter.FontSize * 0.5)
                        {
                            sb.AppendLine();
                            lastY = letter.StartBaseLine.Y;
                        }
                        sb.Append(letter.Value);
                    }
                }
            }

            return sb.ToString();
        }
    }
}
