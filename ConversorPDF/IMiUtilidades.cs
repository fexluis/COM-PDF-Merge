using System.Runtime.InteropServices;

namespace ConversorPDF
{
    [ComVisible(true)]
    [Guid("A5C7E123-4B6C-4F1D-8A7E-1234567890AB")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMiUtilidades
    {
        [DispId(1)]
        string Saludar(string nombre);

        [DispId(2)]
        double Sumar(double a, double b);

        [DispId(3)]
        void CombinarArchivos(object entradas, string salida);
    }
}
