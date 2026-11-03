using System.Collections.Generic;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Models;
namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para la manipulación de datos y objetos de Microsoft Excel.
    /// Provee una capa de abstracción sobre Microsoft.Office.Interop.Excel.
    /// </summary>
    public interface IExcelRepository
    {
        /// <summary> Obtiene la instancia del libro de Excel activo. </summary>
        dynamic GetActiveWorkbook();
        
        /// <summary> Obtiene la instancia de la hoja de cálculo activa. </summary>
        dynamic GetActiveSheet();
        
        /// <summary> Lee el valor de una celda o rango específico. </summary>
        object GetCellValue(dynamic sheet, string range);
        
        /// <summary> Escribe un valor en una celda o rango específico. </summary>
        void SetCellValue(dynamic sheet, string range, object value);
        
        /// <summary> Obtiene una matriz de valores de un rango bidimensional. </summary>
        object[,] GetRangeValues(dynamic sheet, int startRow, int startCol, int endRow, int endCol);
        
        /// <summary> Obtiene una hoja por su nombre. </summary>
        dynamic GetWorksheet(string name);
        
        /// <summary> Verifica si una hoja existe en el libro activo. </summary>
        bool WorksheetExists(string name);
        
        /// <summary> Crea una copia de una hoja existente al final del libro. </summary>
        dynamic CreateWorksheetCopy(string sourceName, string newName);
        
        /// <summary> Elimina una hoja permanentemente. </summary>
        void DeleteWorksheet(string name);
        
        /// <summary> Obtiene todos los datos del área usada de una hoja como matriz C#. </summary>
        object[,] GetWorksheetData(string name);
        
        /// <summary> Guarda el libro de Excel en una ruta específica (Save Copy As). </summary>
        void SaveAs(dynamic workbook, string path);
        
        /// <summary> Exporta una hoja seleccionada a formato PDF con opciones de orientación. </summary>
        void ExportToPdf(dynamic sheet, string path, bool forcedPortrait = false);
        
        /// <summary> Controla el refresco visual de la aplicación Excel. </summary>
        void ScreenUpdating(bool enabled);
        
        /// <summary> Controla la visualización de cuadros de diálogo de alerta de Excel. </summary>
        void DisplayAlerts(bool enabled);
        
        /// <summary> Recupera el valor de una variable global almacenada en el entorno de Excel. </summary>
        string GetGlobalVariable(string varName);
        
        /// <summary> Guarda el libro activo en su ubicación original. </summary>
        void SaveActiveWorkbook();

        /// <summary> Obtiene la ruta de la carpeta de OneDrive configurada. </summary>
        string GetOneDriveFolderPath();
        
        /// <summary> Obtiene los valores de una columna específica dentro de una Tabla (ListObject). </summary>
        List<string> GetTableColumnValues(string tableName, string columnName);
        
        /// <summary> Obtiene el contenido completo de una tabla de Excel como matriz. </summary>
        object[,] GetTableData(string tableName);
        
        /// <summary> Actualiza los datos de una tabla de Excel con una nueva matriz. </summary>
        void SetTableData(string tableName, object[,] data);
    }
}
