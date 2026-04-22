using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using ExcelDna.Integration;
using System.IO;
namespace ConversorPDF.XLL.Core.Repositories
{
    public class ExcelRepository : IExcelRepository
    {
        private readonly dynamic _excelApp;

        public ExcelRepository()
        {
            _excelApp = ExcelDnaUtil.Application;
        }

        public dynamic GetActiveWorkbook() => _excelApp.ActiveWorkbook;

        public dynamic GetActiveSheet() => _excelApp.ActiveSheet;

        public object GetCellValue(dynamic sheet, string range) => sheet.Range[range].Value;

        public void SetCellValue(dynamic sheet, string range, object value) => sheet.Range[range].Value = value;

        public object[,] GetRangeValues(dynamic sheet, int startRow, int startCol, int endRow, int endCol)
        {
            dynamic range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
            return (object[,])range.Value2;
        }

        public dynamic GetWorksheet(string name)
        {
            try { return GetActiveWorkbook().Worksheets[name]; }
            catch { return null; }
        }

        public bool WorksheetExists(string name)
        {
            foreach (dynamic sh in GetActiveWorkbook().Worksheets)
            {
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        public dynamic CreateWorksheetCopy(string sourceName, string newName)
        {
            dynamic wb = GetActiveWorkbook();
            dynamic wsTemplate = wb.Worksheets[sourceName];
            wsTemplate.Visible = -1; // xlSheetVisible
            wsTemplate.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            
            dynamic wsNueva = _excelApp.ActiveSheet;
            wsNueva.Name = newName;
            wsTemplate.Visible = 2; // xlSheetHidden
            return wsNueva;
        }

        public void DeleteWorksheet(string name)
        {
            try
            {
                dynamic ws = GetWorksheet(name);
                if (ws != null)
                {
                    DisplayAlerts(false);
                    ws.Delete();
                    DisplayAlerts(true);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"No se pudo eliminar la hoja '{name}': {ex.Message}");
                DisplayAlerts(true);
            }
        }

        public object[,] GetWorksheetData(string name)
        {
            try
            {
                dynamic ws = GetWorksheet(name);
                if (ws == null) return null;
                dynamic range = ws.UsedRange;
                return (object[,])range.Value2;
            }
            catch { return null; }
        }

        public void SaveAs(dynamic workbook, string path)
        {
            if (File.Exists(path)) File.Delete(path);
            workbook.SaveAs(path, 51 /*xlOpenXMLWorkbook*/);
        }

        public void ExportToPdf(dynamic sheet, string path, bool forcedPortrait = false)
        {
            if (forcedPortrait)
            {
                App.ConfigurarPagina(sheet, true);
            }
            sheet.ExportAsFixedFormat(0 /*xlTypePDF*/, path);
        }

        public void ScreenUpdating(bool enabled) => _excelApp.ScreenUpdating = enabled;

        public void DisplayAlerts(bool enabled) => _excelApp.DisplayAlerts = enabled;

        public string GetGlobalVariable(string varName)
        {
            try
            {
                dynamic nameObj = null;
                try { nameObj = _excelApp.Names.Item(varName); }
                catch { return ""; }

                if (nameObj != null)
                {
                    try
                    {
                        object oRef = nameObj.RefersToRange.Value;
                        return oRef != null ? oRef.ToString() : "";
                    }
                    catch
                    {
                        string formula = nameObj.Value;
                        if (string.IsNullOrEmpty(formula)) return "";
                        return formula.Trim('=', '\"');
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error obteniendo variable global {varName}: {ex.Message}");
                return "";
            }
        }

        public string GetOneDriveFolderPath()
        {
            try
            {
                string path = GetActiveWorkbook().FullName;
                string local = OneDriveHelper.GetOneDriveLocalPath(path);
                return Path.GetDirectoryName(local);
            }
            catch { return ""; }
        }

        public List<string> GetTableColumnValues(string tableName, string columnName)
        {
            var res = new List<string>();
            try
            {
                dynamic ws = GetActiveWorkbook().Worksheets["Dependencias"]; 
                dynamic tbl = ws.ListObjects[tableName];
                dynamic col = tbl.ListColumns[columnName].DataBodyRange;
                if (col != null)
                {
                    foreach (dynamic cell in col.Cells)
                    {
                        string val = cell.Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(val)) res.Add(val.Trim());
                    }
                }
            }
            catch { }
            return res;
        }

        public object[,] GetTableData(string tableName)
        {
            try
            {
                dynamic ws = GetActiveWorkbook().Worksheets["Dependencias"];
                dynamic tbl = ws.ListObjects[tableName];
                return (object[,])tbl.DataBodyRange.Value2;
            }
            catch { return null; }
        }

        public void SetTableData(string tableName, object[,] data)
        {
            try
            {
                dynamic ws = GetActiveWorkbook().Worksheets["Dependencias"];
                dynamic tbl = ws.ListObjects[tableName];
                tbl.DataBodyRange?.Delete();
                if (data != null && data.GetLength(0) > 0)
                {
                    dynamic startCell = tbl.HeaderRowRange.Cells[1].Offset[1, 0];
                    dynamic endCell = startCell.Offset[data.GetLength(0) - 1, data.GetLength(1) - 1];
                    ws.Range[startCell, endCell].Value2 = data;
                }
            }
            catch { }
        }
    }
}
