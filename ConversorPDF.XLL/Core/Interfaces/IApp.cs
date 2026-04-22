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
    [ComVisible(true)]
    [Guid("A5C7E123-4B6C-4F1D-8A7E-1234567890AB")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    /// <summary>
    /// Interfaz principal (Entry Point) expuesta para interacción COM, coordinando los servicios del Add-in.
    /// </summary>
    public interface IApp
    {
        /// <summary> Combina una lista de archivos PDF en uno solo. </summary>
        [DispId(3)]
        void CombinarArchivos(object entradas, string salida);

        /// <summary> Resuelve la ruta local de OneDrive para una ruta de servidor. </summary>
        [DispId(4)]
        string GetOneDriveLocalPath(string path);

        /// <summary> Obtiene el nombre del usuario actual de Office. </summary>
        [DispId(5)]
        string GetCurrentOfficeUser();

        /// <summary> Obtiene la firma del usuario desde la base de datos de firmas. </summary>
        [DispId(6)]
        string GetFirmaHelperUsuario();

        /// <summary> Crea un archivo temporal a partir de una cadena Base64. </summary>
        [DispId(7)]
        string CrearArchivoBase64Temporal(string base64String, string nombreBase);

        /// <summary> Obtiene la firma (Base64) del usuario de Windows actual. </summary>
        [DispId(8)]
        string GetFirmaUsuarioActual();

        /// <summary> Genera un archivo de texto con la firma en Base64 para depuración. </summary>
        [DispId(9)]
        string GenerarFirmaBase64Txt();

        /// <summary> Procesa la conversión de obligaciones presupuestales del contador a PDF. </summary>
        [DispId(10)]
        int ConvertirObligacionContador(object activeWorkbook);

        /// <summary> Unifica los documentos de obligación con sus soportes PDF. </summary>
        [DispId(11)]
        int UnificarObligacionContador(object activeWorkbook, string docType);

        /// <summary> Obtiene los datos de identificación vinculados al usuario. </summary>
        [DispId(12)]
        UserIdentification GetIdentification(string username);
    }
}