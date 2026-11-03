using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;

namespace ConversorPDF.Tests.TestData
{
    /// <summary>
    /// Hoja de Excel falsa mínima: expone un indexador <c>Cells[fila, columna]</c>
    /// sobre la matriz de CONSULTAS para verificar las escrituras en la columna INFORME.
    /// </summary>
    public class HojaExcelFalsa
    {
        private readonly object[,] _data;

        /// <summary> Crea la hoja sobre la matriz indicada (1-based como Excel). </summary>
        /// <param name="data">Matriz de la hoja.</param>
        public HojaExcelFalsa(object[,] data)
        {
            _data = data;
        }

        /// <summary> Celdas de la hoja. </summary>
        public CeldasExcelFalsas Cells
        {
            get { return new CeldasExcelFalsas(_data); }
        }
    }

    /// <summary>
    /// Indexador de celdas sobre la matriz de la hoja falsa.
    /// </summary>
    public class CeldasExcelFalsas
    {
        private readonly object[,] _data;

        /// <summary> Crea el acceso a celdas sobre la matriz indicada. </summary>
        /// <param name="data">Matriz de la hoja.</param>
        public CeldasExcelFalsas(object[,] data)
        {
            _data = data;
        }

        /// <summary> Obtiene o establece el valor de una celda (1-based). </summary>
        /// <param name="fila">Fila de Excel.</param>
        /// <param name="columna">Columna de Excel.</param>
        public object this[int fila, int columna]
        {
            get { return _data[fila, columna]; }
            set { _data[fila, columna] = value; }
        }
    }

    /// <summary>
    /// Implementación falsa de <see cref="IExcelRepository"/> para los tests de informes.
    /// Solo implementa lo que usa el procesador; el resto lanza <see cref="NotImplementedException"/>.
    /// </summary>
    public class ExcelRepositoryFalsa : IExcelRepository
    {
        /// <summary> Datos de la hoja CONSULTAS (1-based). </summary>
        public object[,] Consultas { get; set; }

        /// <summary> Hoja devuelta por <see cref="GetWorksheet"/> (normalmente una HojaExcelFalsa). </summary>
        public object Hoja { get; set; }

        /// <summary> Cantidad de veces que se guardó el libro. </summary>
        public int VecesGuardado { get; private set; }

        /// <summary> Ruta devuelta por <see cref="GetOneDriveFolderPath"/>. </summary>
        public string CarpetaLibro { get; set; }

        /// <inheritdoc/>
        public object[,] GetWorksheetData(string name)
        {
            return string.Equals(name, ConstantsHelper.SHEET_CONSULTAS, StringComparison.OrdinalIgnoreCase)
                ? Consultas
                : null;
        }

        /// <inheritdoc/>
        public dynamic GetWorksheet(string name)
        {
            return Hoja;
        }

        /// <inheritdoc/>
        public void SaveActiveWorkbook()
        {
            VecesGuardado++;
        }

        /// <inheritdoc/>
        public string GetOneDriveFolderPath()
        {
            return CarpetaLibro;
        }

        /// <inheritdoc/>
        public dynamic GetActiveWorkbook() { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public dynamic GetActiveSheet() { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public object GetCellValue(dynamic sheet, string range) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void SetCellValue(dynamic sheet, string range, object value) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public object[,] GetRangeValues(dynamic sheet, int startRow, int startCol, int endRow, int endCol) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public bool WorksheetExists(string name) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public dynamic CreateWorksheetCopy(string sourceName, string newName) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void DeleteWorksheet(string name) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void SaveAs(dynamic workbook, string path) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void ExportToPdf(dynamic sheet, string path, bool forcedPortrait = false) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void ScreenUpdating(bool enabled) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void DisplayAlerts(bool enabled) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public string GetGlobalVariable(string varName) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public List<string> GetTableColumnValues(string tableName, string columnName) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public object[,] GetTableData(string tableName) { throw new NotImplementedException(); }

        /// <inheritdoc/>
        public void SetTableData(string tableName, object[,] data) { throw new NotImplementedException(); }
    }
}
