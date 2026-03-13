using System.Runtime.InteropServices;

namespace ConversorPDF
{
    [ComVisible(true)]
    [Guid("A5C7E123-4B6C-4F1D-8A7E-1234567890AB")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IApp
    {
        [DispId(1)]
        string Saludar(string nombre);

        [DispId(2)]
        double Sumar(double a, double b);

        [DispId(3)]
        void CombinarArchivos(object entradas, string salida);

        [DispId(4)]
        string GetOneDriveLocalPath(string path);

        [DispId(5)]
        string GetCurrentOfficeUser();

        [DispId(6)]
        string GetFirmaBase64(string nombreFirma);

        [DispId(7)]
        string Base64ToTempFile(string base64String, string nombreBase);

        [DispId(8)]
        string FirmaToTempFile(string nombreFirma);

        [DispId(9)]
        string GenerarFirmaBase64Txt();

        [DispId(10)]
        int ConvertirObligacionContador(object activeWorkbook, string userFirma);

        [DispId(11)]
        int UnificarObligacionContador(object activeWorkbook, string docType);
    }
}
