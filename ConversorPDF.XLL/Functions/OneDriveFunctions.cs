using ExcelDna.Integration;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL
{
    /// <summary>
    /// Funciones UDF para resolución de rutas OneDrive/SharePoint y usuario Office.
    /// Uso: =ONEDRIVE.RutaLocal(A1)
    /// </summary>
    public static class OneDriveFunctions
    {
        [ExcelFunction(
            Name        = "ONEDRIVE.RutaLocal",
            Description = "Resuelve la ruta local de un archivo sincronizado en OneDrive o SharePoint",
            Category    = "ESAP - OneDrive")]
        public static string RutaLocal(
            [ExcelArgument(Description = "Ruta OneDrive/SharePoint del archivo")] string path)
        {
            try
            {
                return OneDriveHelper.GetOneDriveLocalPath(path);
            }
            catch (System.Exception ex)
            {
                LoggerHelper.Error("Error en UDF", ex);
                return $"ERROR: {ex.Message}";
            }
        }

        [ExcelFunction(
            Name        = "ONEDRIVE.UsuarioOffice",
            Description = "Retorna el email o nombre del usuario Office activo",
            Category    = "ESAP - OneDrive")]
        public static string UsuarioOffice()
        {
            try
            {
                return OneDriveHelper.GetCurrentOfficeUser();
            }
            catch (System.Exception ex)
            {
                LoggerHelper.Error("Error en UDF", ex);
                return $"ERROR: {ex.Message}";
            }
        }
    }
}