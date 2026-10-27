using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using System.Text.RegularExpressions;
namespace ConversorPDF.XLL.Core.Helpers
{
    public static class FirmaUnificarObHelper
    {

        public static int UnificarObligacionContador(dynamic activeWorkbook, string docType, App util, Action<int, int> progressCallback = null)
        {
            if (activeWorkbook == null) return 0;

            dynamic activeSheet = activeWorkbook.ActiveSheet;
            if (activeSheet.Range["B14"].Value?.ToString() != "REGISTRO PRESUPUESTAL DE OBLIGACION.") return 0;

            if (docType != "OB")
            {
                LoggerHelper.Error("Parámetro docType no válido");
                throw new Exception("Parámetro docType no válido");
            }

            string localPath = FirmaPDFsHelper.ObtenerRutaLocal(activeWorkbook, util);
            if (string.IsNullOrEmpty(localPath)) return 0;

            string pathSoportes = Path.Combine(localPath, App.FolderNames.SOPORTES);
            string pathUnificados = Path.Combine(localPath, App.FolderNames.UNIFICADOS);


            if (!Directory.Exists(pathSoportes) || !ContienePdfsSeguro(pathSoportes))
            {
                System.Windows.MessageBox.Show(
                    "No hay archivos PDF en la carpeta de Soportes.\nVerifica que la carpeta exista y contenga archivos.",
                    "Advertencia - Sin Soportes",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return 0;
            }

            FirmaPDFsHelper.CrearCarpetaSiNoExiste(pathSoportes);
            FirmaPDFsHelper.CrearCarpetaSiNoExiste(pathUnificados);

            int filesCount = 0;

            int totalSheets = activeWorkbook.Worksheets.Count;
            int currentSheet = 0;

            foreach (dynamic ws in activeWorkbook.Worksheets)
            {
                currentSheet++;
                progressCallback?.Invoke(currentSheet, totalSheets);

                string searchOrder = "";
                string unifiedFileName = "";

                if (ws.Range["F16"].Value?.ToString() == "Estado:")
                {
                    searchOrder = ws.Range["C15"].Value?.ToString() ?? "";
                    string categoria = FirmaPDFsHelper.GetCategoria(ws.Range["C38"].Value?.ToString());
                    unifiedFileName = $"OB-{searchOrder}-{categoria}-{FirmaPDFsHelper.SanitizarNombre(ws.Range["I22"].Value?.ToString())}";
                }
                else
                {
                    searchOrder = ws.Range["D15"].Value?.ToString() ?? "";
                    dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                    if (rngObjeto != null)
                    {
                        string categoria = FirmaPDFsHelper.GetCategoria(ws.Range["D" + rngObjeto.Row].Value?.ToString());
                        unifiedFileName = $"OB-{searchOrder}-{categoria}-{FirmaPDFsHelper.SanitizarNombre(ws.Range["K22"].Value?.ToString())}";
                    }
                    else
                    {
                        unifiedFileName = $"OB-{searchOrder}-XXX-{FirmaPDFsHelper.SanitizarNombre(ws.Range["K22"].Value?.ToString())}";
                    }
                }

                unifiedFileName = (unifiedFileName.Length > 60 ? unifiedFileName.Substring(0, 60) : unifiedFileName) + ".pdf";

                string[] matchSoportes = BuscarArchivosPDF(pathSoportes, searchOrder);

                if (matchSoportes.Length > 0)
                {
                    string targetFile = Path.Combine(localPath, App.FolderNames.OBLIGACIONES, unifiedFileName);
                    string outputPdfPath = Path.Combine(pathUnificados, unifiedFileName);

                    int totalFiles = matchSoportes.Length + (File.Exists(targetFile) ? 1 : 0);
                    string[] listaArchivos = new string[totalFiles];


                    int idx = 0;
                    if (File.Exists(targetFile))
                    {
                        listaArchivos[idx++] = targetFile;
                    }

                    foreach (var sop in matchSoportes)
                    {
                        listaArchivos[idx++] = sop;
                    }

                    try
                    {
                        util.CombinarArchivos(listaArchivos, outputPdfPath);
                        filesCount++;
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error($"Error unificando PDF: {outputPdfPath}", ex);
                    }
                }
            }
            return filesCount;
        }

        /// <summary>
        /// Comprueba de forma segura y recursiva si existe al menos un archivo PDF en la carpeta o sus subcarpetas.
        /// Ignora carpetas sin permisos para evitar interrupciones.
        /// </summary>
        /// <param name="pathFolder">Ruta a evaluar.</param>
        /// <returns>Verdadero si encuentra al menos un PDF.</returns>
        private static bool ContienePdfsSeguro(string pathFolder)
        {
            var stack = new System.Collections.Generic.Stack<string>();
            stack.Push(pathFolder);

            while (stack.Count > 0)
            {
                string dir = stack.Pop();
                try
                {
                    if (Directory.GetFiles(dir, "*.pdf").Length > 0)
                        return true;

                    foreach (var subDir in Directory.GetDirectories(dir))
                    {
                        stack.Push(subDir);
                    }
                }
                catch (UnauthorizedAccessException) { /* Omitir carpeta sin permisos */ }
                catch (PathTooLongException) { /* Omitir carpeta inaccesible */ }
            }

            return false;
        }

        /// <summary>
        /// Busca de forma segura y recursiva archivos PDF de soporte que coincidan con el número de orden.
        /// Utiliza recorrido por pila para evitar excepciones de acceso en subcarpetas.
        /// </summary>
        /// <param name="pathFolder">Ruta raíz de la carpeta de soportes.</param>
        /// <param name="searchOrder">Número de obligación a buscar (coincidencia exacta aislada de otros dígitos).</param>
        /// <returns>Arreglo de rutas de archivos PDF que coinciden, ordenados alfabéticamente.</returns>
        private static string[] BuscarArchivosPDF(string pathFolder, string searchOrder)
        {
            if (!Directory.Exists(pathFolder) || string.IsNullOrEmpty(searchOrder))
            {
                return new string[0];
            }

            try
            {
                var matched = new System.Collections.Generic.List<string>();
                
                // Expresión regular mejorada para buscar el número aislado de otros dígitos numéricos.
                // Usar lookbehinds y lookaheads (?<!\d) y (?!\d) permite que coincida si está rodeado
                // de letras o guiones bajos (ej. Soporte_15_pago.pdf sí coincidirá con 15).
                string pattern = $@"(?<!\d){Regex.Escape(searchOrder)}(?!\d)";

                var stack = new System.Collections.Generic.Stack<string>();
                stack.Push(pathFolder);

                while (stack.Count > 0)
                {
                    string dir = stack.Pop();
                    try
                    {
                        foreach (var f in Directory.GetFiles(dir, "*.pdf"))
                        {
                            string fileName = Path.GetFileNameWithoutExtension(f);
                            
                            if (Regex.IsMatch(fileName, pattern, RegexOptions.IgnoreCase))
                            {
                                matched.Add(f);
                            }
                        }

                        foreach (var d in Directory.GetDirectories(dir))
                        {
                            stack.Push(d);
                        }
                    }
                    catch (UnauthorizedAccessException) { /* Omitir carpeta sin permisos */ }
                    catch (PathTooLongException) { /* Omitir carpeta inaccesible */ }
                }

                // Retornar los archivos ordenados para garantizar un orden consistente al unificar
                matched.Sort();
                return matched.ToArray();
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error buscando PDFs de soportes: {ex.Message}");
                return new string[0];
            }
        }
    }
}