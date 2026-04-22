using ExcelDna.Integration;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Helpers;
namespace ConversorPDF.XLL
{
    /// <summary>
    /// Funciones UDF para extracción de texto PDF, disponibles como fórmulas en celdas Excel.
    /// Uso: =PDF.ExtraerTexto("C:\ruta\archivo.pdf")
    /// </summary>
    public static class PdfFunctions
    {
        /// <summary>
        /// Extrae todo el texto de un archivo PDF.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF en el disco.</param>
        /// <returns>Todo el texto extraído del documento.</returns>
        public static string ExtraerTexto(
            [ExcelArgument(Description = "Ruta completa al archivo PDF")] string rutaPdf)
        {
            try
            {
                return new PdfToTextoHelper().ExtraerTodoTexto(rutaPdf);
            }
            catch (System.Exception ex)
            {
                LoggerHelper.Error("Error en UDF PDF", ex);
                return $"ERROR: {ex.Message}";
            }
        }

        /// <summary>
        /// Extrae el texto de una página específica de un PDF.
        /// </summary>
        /// <param name="rutaPdf">Ruta completa al archivo PDF.</param>
        /// <param name="numeroPagina">Número de la página (empezando desde 1).</param>
        /// <returns>Texto contenido en la página especificada.</returns>
        public static string ExtraerTextoPagina(
            [ExcelArgument(Description = "Ruta completa al archivo PDF")]  string rutaPdf,
            [ExcelArgument(Description = "Número de página (empieza en 1)")] int numeroPagina)
        {
            try
            {
                return new PdfToTextoHelper().ExtraerTextoPagina(rutaPdf, numeroPagina);
            }
            catch (System.Exception ex)
            {
                LoggerHelper.Error("Error en UDF PDF", ex);
                return $"ERROR: {ex.Message}";
            }
        }
    }
}