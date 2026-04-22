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


            if (!Directory.Exists(pathSoportes) || Directory.GetFiles(pathSoportes, "*.pdf").Length == 0)
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
                    unifiedFileName = $"OB-{searchOrder}-{categoria}-{ws.Range["I22"].Value}";
                }
                else
                {
                    searchOrder = ws.Range["D15"].Value?.ToString() ?? "";
                    dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                    if (rngObjeto != null)
                    {
                        string categoria = FirmaPDFsHelper.GetCategoria(ws.Range["D" + rngObjeto.Row].Value?.ToString());
                        unifiedFileName = $"OB-{searchOrder}-{categoria}-{ws.Range["K22"].Value}";
                    }
                    else
                    {
                        unifiedFileName = $"OB-{searchOrder}-XXX-{ws.Range["K22"].Value}";
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

                    util.CombinarArchivos(listaArchivos, outputPdfPath);
                    filesCount++;
                }
            }
            return filesCount;
        }

        private static string[] BuscarArchivosPDF(string pathFolder, string searchOrder)
        {
            if (!Directory.Exists(pathFolder) || string.IsNullOrEmpty(searchOrder))
            {
                return new string[0];
            }

            try
            {
                var allPdfs = Directory.GetFiles(pathFolder, "*.pdf", SearchOption.TopDirectoryOnly);
                var matched = new System.Collections.Generic.List<string>();

                foreach (var pdf in allPdfs)
                {
                    if (Path.GetFileName(pdf).IndexOf(searchOrder, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        matched.Add(pdf);
                    }
                }

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