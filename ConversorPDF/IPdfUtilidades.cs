using System.Runtime.InteropServices;

namespace ConversorPDF
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("D8A1B2C3-4E5F-6A7B-8C9D-0E1F2A3B4C5D")]
    public interface IPdfUtilidades
    {
        string ExtraerTodoTexto(string rutaPdf);
        string ExtraerTextoPagina(string rutaPdf, int numeroPagina); // 1-based
        string ExtraerPalabras(string rutaPdf, int numeroPagina); // Devuelve lista detallada con coordenadas
        string ExtraerTextoEnRectangulo(string rutaPdf, int numeroPagina, double left, double bottom, double width, double height); // Extrae texto en un área específica
    }
}
