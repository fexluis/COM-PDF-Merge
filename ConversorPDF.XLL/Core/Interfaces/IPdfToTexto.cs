using System.Runtime.InteropServices;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para la extracción de texto de archivos PDF (vía iText7 o similar).
    /// Expuesta vía COM para ser consumida desde VBA.
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("D8A1B2C3-4E5F-6A7B-8C9D-0E1F2A3B4C5D")]
    public interface IPdfToTexto
    {
        /// <summary> Extrae todo el texto contenido en el archivo PDF. </summary>
        string ExtraerTodoTexto(string rutaPdf);

        /// <summary> Obtiene la cantidad de páginas del PDF (0 si no pudo abrirse). </summary>
        int ContarPaginas(string rutaPdf);

        /// <summary> Extrae el texto de una página específica (1-based index). </summary>
        string ExtraerTextoPagina(string rutaPdf, int numeroPagina);

        /// <summary> Devuelve una lista detallada de palabras con sus coordenadas espaciales. </summary>
        string ExtraerPalabras(string rutaPdf, int numeroPagina);

        /// <summary> Extrae exclusivamente el texto dentro de un rectángulo delimitado por coordenadas. </summary>
        string ExtraerTextoEnRectangulo(string rutaPdf, int numeroPagina, double left, double bottom, double width, double height);
    }
}