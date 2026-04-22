using System;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de procesamiento de obligaciones del contador.
    /// </summary>
    public interface IContadorService
    {
        /// <summary>
        /// Convierte las hojas de obligaciones en archivos PDF individuales con firma y metadatos.
        /// </summary>
        /// <param name="activeWorkbook">Libro de Excel activo.</param>
        /// <param name="userFirma">Firma en formato base64.</param>
        /// <param name="progressCallback">Callback para reporte de progreso.</param>
        /// <returns>Cantidad de archivos procesados exitosamente.</returns>
        int ConvertirObligacionContador(dynamic activeWorkbook, string userFirma, Action<int, int> progressCallback = null);

        /// <summary>
        /// Unifica los archivos de obligaciones con sus respectivos soportes PDF.
        /// </summary>
        /// <param name="activeWorkbook">Libro de Excel activo.</param>
        /// <param name="docType">Tipo de documento (ej. "OB").</param>
        /// <param name="progressCallback">Callback para reporte de progreso.</param>
        /// <returns>Cantidad de archivos unificados exitosamente.</returns>
        int UnificarObligacionContador(dynamic activeWorkbook, string docType, Action<int, int> progressCallback = null);
    }
}
