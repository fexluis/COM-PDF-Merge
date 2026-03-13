using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder; // Required for dynamic dispatch

namespace ConversorPDF
{
    public static class ContadorUtilidades
    {
        private const string APP_NAME = "ESAP - Contador - v1.4";
        private const string PATH_OBLIGACIONES = @"\Obligaciones\";

        public static int ConvertirObligacionContador(dynamic activeWorkbook, string userFirma, MiUtilidades util)
        {
            if (activeWorkbook == null) return 0;

            dynamic activeSheet = activeWorkbook.ActiveSheet;
            if (activeSheet.Range["B14"].Value?.ToString() != "REGISTRO PRESUPUESTAL DE OBLIGACION.")
            {
                return 0;
            }

            string localPath = ObtenerRutaLocal(activeWorkbook, util);
            if (string.IsNullOrEmpty(localPath)) return 0;

            CrearCarpetaSiNoExiste(localPath + PATH_OBLIGACIONES);

            string tempPath = util.FirmaToTempFile(userFirma);
            if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
            {
                throw new Exception("No se pudo crear el archivo temporal con la imagen.");
            }

            // Deshabilitar actualización de pantalla si excel application está disponible
            dynamic excelApp = activeWorkbook.Application;
            excelApp.ScreenUpdating = false;

            int filesCount = 0;

            foreach (dynamic ws in activeWorkbook.Worksheets)
            {
                dynamic rngFirma = ws.Range["A1:BH300"].Find("FIRMA(S) RESPONSABLE(S)");
                string firmaPos = (rngFirma != null) ? "J" + (rngFirma.Row - 3) : "J57";

                string categoria = "";
                string cellValue = "";

                if (ws.Range["F16"].Value?.ToString() == "Estado:")
                {
                    categoria = GetCategoria(ws.Range["C38"].Value?.ToString());
                    cellValue = $"OB-{ws.Range["C15"].Value}-{categoria}-{ws.Range["I22"].Value}";
                }
                else
                {
                    dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                    if (rngObjeto != null)
                    {
                        categoria = GetCategoria(ws.Range["D" + rngObjeto.Row].Value?.ToString());
                    }
                    else
                    {
                        categoria = "XXX";
                    }
                    cellValue = $"OB-{ws.Range["D15"].Value}-{categoria}-{ws.Range["K22"].Value}";
                }

                // Insertar Firma
                InsertarFirma(ws, firmaPos, tempPath);

                // Configurar Página
                int rowNum = (rngFirma != null) ? (int)rngFirma.Row : 60; // fallback a 60
                ConfigurarPagina(ws, rowNum, excelApp);

                // Exportar PDF
                string fileName = cellValue.Length > 60 ? cellValue.Substring(0, 60) : cellValue;
                string fullPdfPath = localPath + PATH_OBLIGACIONES + fileName + ".pdf";

                try
                {
                    // xlTypePDF = 0
                    ws.ExportAsFixedFormat(0, fullPdfPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error exportando PDF: " + ex.Message);
                }

                filesCount++;
            }

            try
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
            }
            catch { }

            excelApp.ScreenUpdating = true;

            return filesCount;
        }

        public static int UnificarObligacionContador(dynamic activeWorkbook, string docType, MiUtilidades util)
        {
            if (activeWorkbook == null) return 0;

            dynamic activeSheet = activeWorkbook.ActiveSheet;
            if (activeSheet.Range["B14"].Value?.ToString() != "REGISTRO PRESUPUESTAL DE OBLIGACION.") return 0;

            if (docType != "OB")
            {
                throw new Exception("Parámetro docType no válido");
            }

            string localPath = ObtenerRutaLocal(activeWorkbook, util);
            if (string.IsNullOrEmpty(localPath)) return 0;

            string pathSoportes = localPath + @"\S\";
            string pathUnificados = localPath + @"\Unificados\";
            
            if (!Directory.Exists(pathSoportes) || Directory.GetFiles(pathSoportes, "*.pdf").Length == 0)
            {
                throw new Exception("No hay archivos en la carpeta de Soportes!");
            }

            CrearCarpetaSiNoExiste(pathSoportes);
            CrearCarpetaSiNoExiste(pathUnificados);

            int filesCount = 0;

            foreach (dynamic ws in activeWorkbook.Worksheets)
            {
                string searchOrder = "";
                string unifiedFileName = "";

                if (ws.Range["F16"].Value?.ToString() == "Estado:")
                {
                    searchOrder = ws.Range["C15"].Value?.ToString() ?? "";
                    string categoria = GetCategoria(ws.Range["C38"].Value?.ToString());
                    unifiedFileName = $"OB-{searchOrder}-{categoria}-{ws.Range["I22"].Value}";
                }
                else
                {
                    searchOrder = ws.Range["D15"].Value?.ToString() ?? "";
                    dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                    if (rngObjeto != null)
                    {
                        string categoria = GetCategoria(ws.Range["D" + rngObjeto.Row].Value?.ToString());
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
                    string targetFile = localPath + PATH_OBLIGACIONES + unifiedFileName;
                    string outputPdfPath = pathUnificados + unifiedFileName;

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

        private static string GetCategoria(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return "XXX";

            string textoMayus = texto.ToUpperInvariant();

            if (textoMayus.Contains("VIAT")) return "VIA";
            if (textoMayus.Contains("ARL") || textoMayus.Contains("RIESGO")) return "ARL";
            if (textoMayus.Contains("HONOR")) return "HON";
            if (textoMayus.Contains("SEGURIDA") || textoMayus.Contains("SOCIAL")) return "SSO";
            if (textoMayus.Contains("NOMINA") || textoMayus.Contains("NÓMINA")) return "NOM";
            if (textoMayus.Contains("CESANT")) return "FNA";
            if (textoMayus.Contains("ARREND") || textoMayus.Contains("ARRIEND")) return "ARR";
            if (textoMayus.Contains("SERVICIO") || textoMayus.Contains("PUBLICO") || textoMayus.Contains("PÚBLICO")) return "SPU";
            if (textoMayus.Contains("ESTUDIANTE") || textoMayus.Contains("MONITOR") || textoMayus.Contains("PRACTICA") || textoMayus.Contains("ESTIMULO")) return "EST";
            if (textoMayus.Contains("4X") || textoMayus.Contains("4 X") || textoMayus.Contains("X1000") || textoMayus.Contains("X 1000") || textoMayus.Contains("POR MIL")) return "4PM";

            return "XXX";
        }

        private static string ObtenerRutaLocal(dynamic activeWorkbook, MiUtilidades util)
        {
            try
            {
                string path = util.GetOneDriveLocalPath(activeWorkbook.Path);
                if (string.IsNullOrEmpty(path)) return "";

                path = path.Replace(@"Escritorio\AÑO 2026 PAGADURIA\ALEX BOLETINES 2026", "_ALEX BOLETINES 2026");
                return path;
            }
            catch
            {
                return "";
            }
        }

        private static void CrearCarpetaSiNoExiste(string ruta)
        {
            if (!Directory.Exists(ruta))
            {
                try
                {
                    Directory.CreateDirectory(ruta);
                }
                catch { }
            }
        }

        private static void InsertarFirma(dynamic ws, string rangoAddress, string tempFile)
        {
            try
            {
                dynamic img = ws.Pictures.Insert(tempFile);
                dynamic targetRange = ws.Range[rangoAddress];
                
                img.Left = targetRange.Left;
                img.Top = targetRange.Top;
                img.Width = 60;
                img.Height = 60;
                img.Placement = 1;
                img.PrintObject = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al insertar imagen: " + ex.Message);
            }
        }

        private static void ConfigurarPagina(dynamic ws, int categoria, dynamic excelApp)
        {
            try
            {
                dynamic ps = ws.PageSetup;
                ps.LeftMargin = excelApp.InchesToPoints(0.1);
                ps.RightMargin = excelApp.InchesToPoints(0.1);
                ps.TopMargin = excelApp.InchesToPoints(0.1);
                ps.BottomMargin = excelApp.InchesToPoints(0.1);
                ps.HeaderMargin = excelApp.InchesToPoints(0.1);
                ps.FooterMargin = excelApp.InchesToPoints(0.1);
                ps.CenterHorizontally = true;
                ps.CenterVertically = true;
                ps.Zoom = false;
                ps.FitToPagesTall = 1;
                ps.FitToPagesWide = 1;

                if (categoria > 62)
                {
                    // xlPortrait = 1
                    ps.Orientation = 1;
                }
                else
                {
                    // xlLandscape = 2
                    ps.Orientation = 2;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Configurando página: " + ex.Message);
            }
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
            catch
            {
                return new string[0];
            }
        }
    }
}
