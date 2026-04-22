using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using ConversorPDF.XLL;
using ExcelDna.Integration;
namespace ConversorPDF.XLL.FormProgramador
{
    public static class ExcelHelper
    {
        public static dynamic GetApp()
        {
            return ExcelDnaUtil.Application;
        }

        public static dynamic GetConfigSheet()
        {
            dynamic app = GetApp();
            if (app == null || app.ActiveWorkbook == null) return null;
            try
            {

                dynamic ws = app.ActiveWorkbook.Sheets["Config"];
                if (ws != null) ws.Visible = 2; // xlSheetVeryHidden
                return ws;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"No se pudo obtener la hoja Config: {ex.Message}");
                return null;
            }
        }

        public static dynamic GetDependenciasSheet()
        {
            dynamic app = GetApp();
            if (app == null || app.ActiveWorkbook == null) return null;
            try
            {

                dynamic ws = app.ActiveWorkbook.Sheets["Dependencias"];

                if (ws != null) ws.Visible = 2; // xlSheetVeryHidden
                return ws;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"No se pudo obtener la hoja Dependencias: {ex.Message}");
                return null;
            }
        }

        private static string ParseDateObj(object obj)
        {
            if (obj == null) return "";
            if (obj is DateTime time)
            {
                return time.ToString("dd/MM/yyyy");
            }
            string s = obj.ToString();
            if (DateTime.TryParse(s, out DateTime dtResult))
            {
                return dtResult.ToString("dd/MM/yyyy");
            }
            return s;
        }

        public static string[] LeerDatosConfig()
        {
            dynamic ws = GetConfigSheet();
            if (ws == null) return new string[] { "", "", "" };

            try
            {
                object vB1 = ws.Range["B1"].Value;
                object vB2 = ws.Range["B2"].Value;
                object vB3 = ws.Range["B3"].Value;

                string f1 = ParseDateObj(vB1);
                string f2 = ParseDateObj(vB2);
                string cdp = vB3 != null ? vB3.ToString() : "";

                return new string[] { f1, f2, cdp };
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error leyendo datos de configuración en ExcelHelper", ex);
                return new string[] { "", "", "" };
            }
        }

        public static void EscribirDatosConfig(object fInicio, object fFin, string cdp)
        {
            dynamic ws = GetConfigSheet();
            if (ws == null) return;

            try
            {
                if (fInicio != null) ws.Range["B1"].Value = fInicio;
                if (fFin != null) ws.Range["B2"].Value = fFin;
                if (cdp != null) ws.Range["B3"].Value = cdp;
            }
            catch { }
        }

        public static List<string> LeerListaCDP()
        {
            List<string> lista = new List<string>();
            dynamic ws = GetDependenciasSheet();
            if (ws == null) return lista;

            try
            {
                dynamic table = ws.ListObjects["tblDependencias"];
                dynamic cdpCol = table.ListColumns["CDP"].DataBodyRange;


                if (cdpCol != null)
                {
                    foreach (dynamic cell in cdpCol.Cells)
                    {
                        object v = cell.Value;
                        string val = v?.ToString();
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            lista.Add(val.Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error leyendo lista de CDPs desde tblDependencias", ex);
            }


            return lista;
        }

        public static string[] BuscarDependenciaPorCDP(string cdp)
        {
            if (string.IsNullOrWhiteSpace(cdp)) return new string[] { "", "" };

            dynamic ws = GetDependenciasSheet();
            if (ws == null) return new string[] { "", "" };

            try
            {
                dynamic table = ws.ListObjects["tblDependencias"];
                dynamic dataRange = table.DataBodyRange;
                int cdpIdx = table.ListColumns["CDP"].Index;
                int depIdx = table.ListColumns["DEPENDENCIA"].Index;
                int descIdx = table.ListColumns["DEPENDENCIA DESCRIPCION"].Index;

                object[,] values = (object[,])dataRange.Value2;
                int rows = values.GetLength(0);

                for (int r = 1; r <= rows; r++)
                {
                    object v = values[r, cdpIdx];
                    string val = v != null ? v.ToString() : "";
                    if (val.Trim() == cdp.Trim())
                    {
                        string dep = values[r, depIdx]?.ToString() ?? "";
                        string desc = values[r, descIdx]?.ToString() ?? "";
                        return new string[] { dep, desc };
                    }
                }
            }
            catch { }
            return new string[] { "", "" };
        }

        public static void EjecutarMacro(string nombreMacro, object arg1, object arg2, object arg3)
        {
            dynamic app = GetApp();
            if (app != null)
            {
                try
                {
                    app.Run(nombreMacro, arg1, arg2, arg3);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"Error al ejecutar la macro {nombreMacro}", ex);
                    System.Windows.Forms.MessageBox.Show("Error al ejecutar la macro " + nombreMacro + ":\n" + ex.Message,

                        "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }


        public static string GetGlobalVariable(string varName)
        {
            dynamic app = GetApp();
            if (app == null) return "";
            try
            {
                // Intentar obtener el objeto Name de forma segura
                dynamic nameObj = null;
                try { nameObj = app.Names.Item(varName); }
                catch { return ""; } // El nombre no existe, retornar vacío sin loguear advertencia

                if (nameObj != null)
                {
                    // Un nombre puede referirse a un rango o a una constante/fórmula
                    try
                    {
                        object oRef = nameObj.RefersToRange.Value;
                        return oRef != null ? oRef.ToString() : "";
                    }
                    catch
                    {
                        // Si no es un rango, intentamos obtener el valor de la fórmula
                        string formula = nameObj.Value; // e.g. "=\"123\"" o "=TRUE"
                        if (string.IsNullOrEmpty(formula)) return "";
                        return formula.Trim('=', '\"');
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                // Solo logueamos errores que no sean el típico "no encontrado"
                LoggerHelper.Warning($"Error inesperado obteniendo variable global {varName}: {ex.Message}");
                return "";
            }
        }

        public static object[,] LeerConfigRows4to16()
        {
            dynamic ws = GetConfigSheet();
            if (ws == null) return null;
            try
            {
                return (object[,])ws.Range["B4:B16"].Value2;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error leyendo filas 4 a 16 de configuración", ex);
                return null;
            }
        }

        public static void EscribirConfigRows4to16(object[,] values)
        {
            dynamic ws = GetConfigSheet();
            if (ws == null) return;
            try
            {
                ws.Range["B4:B16"].Value2 = values;
            }
            catch { }
        }

        public static object[,] LeerTablaData(string tableName)
        {
            dynamic wsDep = GetDependenciasSheet();
            if (wsDep == null) return null;
            try
            {
                dynamic tbl = wsDep.ListObjects[tableName];
                if (tbl == null || tbl.DataBodyRange == null) return null;
                return (object[,])tbl.DataBodyRange.Value2;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error leyendo tabla {tableName}", ex);
                return null;
            }
        }

        public static void EscribirTablaData(string tableName, object[,] data)
        {
            dynamic wsDep = GetDependenciasSheet();
            if (wsDep == null) return;
            try
            {
                dynamic tbl = wsDep.ListObjects[tableName];
                if (tbl == null) return;

                // Limpiar datos actuales
                tbl.DataBodyRange?.Delete();

                // Insertar nuevos datos si los hay
                if (data != null && data.GetLength(0) > 0)
                {
                    int rows = data.GetLength(0);
                    int cols = data.GetLength(1);
                    dynamic startCell = tbl.HeaderRowRange.Cells[1].Offset[1, 0];
                    dynamic endCell = startCell.Offset[rows - 1, cols - 1];
                    dynamic targetRange = wsDep.Range[startCell, endCell];
                    targetRange.Value2 = data;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error escribiendo en tabla {tableName}", ex);
            }
        }
    }
}