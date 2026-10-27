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
    public static class FirmaPDFsHelper
    {
        /// <summary>
        /// Limpia el prefijo 'data:image/...;base64,' de una cadena Base64 cruda.
        /// </summary>
        /// <param name="raw">Cadena Base64 cruda.</param>
        /// <returns>Cadena Base64 limpia.</returns>
        public static string LimpiarPrefijoBase64(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;
            int idx = raw.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
            return idx >= 0 ? raw.Substring(idx + 7) : raw;
        }

        public static int LoopSheetsFirmarAndSaveAsPDF(dynamic activeWorkbook, App util, Action<int, int> progressCallback = null)
        {
            if (activeWorkbook == null) return 0;

            string localPath = ObtenerRutaLocal(activeWorkbook, util);
            if (string.IsNullOrEmpty(localPath)) return 0;

            string tempPath = util.GetFirmaUsuarioActualInteractivo();
            if (string.IsNullOrEmpty(tempPath))
            {
                // FASE 1 / FASE 6: Feedback al usuario cuando cancela firma (usar WPF MessageBox para consistencia)
                LoggerHelper.Info("Proceso cancelado por el usuario o firma no disponible.");
                System.Windows.MessageBox.Show(
                    "El proceso de conversión a PDF fue cancelado porque no se configuró o no se encontró una firma válida.",
                    "Conversión Cancelada",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                    
                // FASE 6: Abortar proceso si usuario cancela (ruta temp vacía por cancelación o fallo validado)
                return 0;
            }
            if (!File.Exists(tempPath))
            {
                // FASE 2: Reintento defensivo si el archivo temporal desaparece (ej. cuarentena por Antivirus)
                LoggerHelper.Warning("Archivo temporal de firma desapareció, intentando generarlo de nuevo...");
                tempPath = util.GetFirmaUsuarioActualInteractivo(); // Segundo intento
                
                if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
                {
                    LoggerHelper.Error("Archivo temporal de firma no existe en disco tras un segundo intento de generación.");
                    throw new Exception("No se pudo obtener la firma del usuario. Verifique que tenga su firma configurada en la base de datos.");
                }
            }

            try
            {
                dynamic excelApp = activeWorkbook.Application;
                using (ExcelOptimizeHelper.Optimize(excelApp))
                {
                    int count = 0;
                    int totalSheets = activeWorkbook.Worksheets.Count;
                    int currentSheet = 0;

                    string carpetaDestino = DeterminarCarpetaDestino(activeWorkbook);
                    if (carpetaDestino == null) return 0;

                    foreach (dynamic ws in activeWorkbook.Worksheets)
                    {
                        currentSheet++;
                        progressCallback?.Invoke(currentSheet, totalSheets);

                        var info = ObtenerNombrePdfYCeldaFirma(ws, carpetaDestino);
                        string nombrePdf = info.Item1;
                        string celdaFirma = info.Item2;

                        if (nombrePdf == "0" || celdaFirma == "0") continue;

                        InsertarDatosIdentificacion(ws, carpetaDestino, util);
                        InsertarFirma(ws, celdaFirma, tempPath);
                        ConfigurarPagina(ws);

                        string fullPdfPath = Path.Combine(localPath, carpetaDestino, nombrePdf + ".pdf");
                        string dir = Path.GetDirectoryName(fullPdfPath);
                        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                        try
                        {
                            ws.ExportAsFixedFormat(0, fullPdfPath);
                            count++;
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Error($"Error exportando PDF: {fullPdfPath}", ex);
                        }
                    }

                    return count;
                }
            }
            finally
            {
                if (File.Exists(tempPath)) try { File.Delete(tempPath); } catch (Exception ex) { LoggerHelper.Warning($"No se pudo eliminar firma temporal: {ex.Message}"); }
            }
        }

        private static string DeterminarCarpetaDestino(dynamic activeWorkbook)
        {
            try
            {
                dynamic activeSheet = activeWorkbook.ActiveSheet;
                string a16 = activeSheet.Range["A16"].Value?.ToString() ?? string.Empty;
                string b14 = activeSheet.Range["B14"].Value?.ToString() ?? string.Empty;
                string a14 = activeSheet.Range["A14"].Value?.ToString() ?? string.Empty;

                if (a16 == "CERTIFICADO DE DISPONIBILIDAD PRESUPUESTAL") return App.FolderNames.CDPS;
                if (b14 == "REGISTRO PRESUPUESTAL DEL COMPROMISO") return App.FolderNames.RPS;
                if (a14 == "ORDEN DE PAGO PRESUPUESTAL") return App.FolderNames.BOLETINES;
                if (b14 == "REGISTRO PRESUPUESTAL DE OBLIGACION.") return App.FolderNames.OBLIGACIONES;
                if (a14.Contains("Comisión Servicio")) return App.FolderNames.VERIFICADOS;
                return null;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error determinando carpeta destino: {ex.Message}");
                return null;
            }
        }

        private static (string Nombre, string Celda) ObtenerNombrePdfYCeldaFirma(dynamic ws, string carpetaDestino)
        {
            dynamic celdaBusqueda = (ws.Range["A1:BH300"].Find("FIRMA(S) RESPONSABLE(S)") ?? ws.Range["A1:BH300"].Find("Firma Responsable")) ?? ws.Range["A1:BH300"].Find("firma", MatchCase: false);

            if (celdaBusqueda == null) return ("0", "0");

            int rowFirma = celdaBusqueda.Row - 3;
            string prefijo = "";
            string nombreOrden = "";
            string referenciaFirma = "";

            if (carpetaDestino == App.FolderNames.BOLETINES)
            {
                if (ws.Range["X16"].Value != null)
                {
                    nombreOrden = (ws.Range["J16"].Value?.ToString() == "Anulada") ? "Anulada" : ws.Range["J30"].Value?.ToString();
                    prefijo = ws.Range["J15"].Value?.ToString().Replace("-", "") + "_" + ws.Range["B15"].Value?.ToString() + "_" + ws.Range["X16"].Value?.ToString() + "_";
                    referenciaFirma = "Q" + rowFirma;
                }
                else
                {
                    nombreOrden = string.IsNullOrEmpty(ws.Range["I30"].Value?.ToString()) ? ws.Range["I31"].Value?.ToString() : ws.Range["I30"].Value?.ToString();
                    if (ws.Range["I16"].Value?.ToString() == "Anulada") nombreOrden = "Anulada";
                    prefijo = ws.Range["I15"].Value?.ToString().Replace("-", "") + "_" + ws.Range["B15"].Value?.ToString() + "_" + ws.Range["U16"].Value?.ToString() + "_";
                    referenciaFirma = "O53";
                }
            }
            else if (carpetaDestino == App.FolderNames.OBLIGACIONES)
            {
                string numero = "";
                string tercero = "";
                string objeto = "";

                if (ws.Range["F16"].Value?.ToString() == "Estado:")
                {
                    numero = ws.Range["C15"].Value?.ToString() ?? "";
                    objeto = ws.Range["C38"].Value?.ToString() ?? "";
                    tercero = ws.Range["I22"].Value?.ToString() ?? "";
                }
                else
                {
                    dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                    numero = ws.Range["D15"].Value?.ToString() ?? "";
                    objeto = ws.Range["D" + (rngObjeto?.Row ?? 1)].Value?.ToString() ?? "";
                    tercero = ws.Range["K22"].Value?.ToString() ?? "";
                }
                
                string categoria = GetCategoria(objeto);
                string terceroLimpio = SanitizarNombre(tercero);
                string finalNameOb = $"OB-{numero}-{categoria}-{terceroLimpio}";
                
                if (finalNameOb.Length > 60) finalNameOb = finalNameOb.Substring(0, 60);

                return (finalNameOb, "L" + rowFirma);
            }
            else if (carpetaDestino == App.FolderNames.CDPS)
            {
                prefijo = ws.Range["E18"].Value?.ToString().Replace("-", "") + "_CDP_" + ws.Range["B18"].Value + "_";
                referenciaFirma = "G" + (celdaBusqueda.Row - 4);
            }
            else if (carpetaDestino == App.FolderNames.RPS)
            {
                prefijo = ws.Range["I16"].Value?.ToString().Replace("-", "") + "_RP_" + ws.Range["F16"].Value + "_";
                nombreOrden = ws.Range["I20"].Value?.ToString();
                referenciaFirma = "M" + rowFirma;
            }
            else if (carpetaDestino == App.FolderNames.VERIFICADOS)
            {
                prefijo = ws.Range["N15"].Value?.ToString().Replace("-", "") + "_V_" + ws.Range["C15"].Value + "_";
                nombreOrden = ws.Range["A23"].Value?.ToString();
                referenciaFirma = "Q" + rowFirma;
            }

            string finalName = SanitizarNombre(prefijo + nombreOrden);
            return (finalName, referenciaFirma);
        }

        public static string SanitizarNombreCompleto(string nombre)
        {
            if (string.IsNullOrEmpty(nombre)) return "Archivo_" + DateTime.Now.Ticks;
            string n = nombre.Replace("  ", " ").Replace("-", " ");
            n = System.Globalization.CultureInfo.InvariantCulture.TextInfo
                    .ToTitleCase(n.ToLowerInvariant())
                    .Replace(" ", "");
            var partes = n.Split('_');
            for (int i = 0; i < partes.Length; i++) if (partes[i] == "") partes[i] = "DESCONOCIDO";
            return string.Join("_", partes);
        }

        public static string SanitizarNombre(string nombre)
        {
            string n = SanitizarNombreCompleto(nombre);
            return n.Length > 30 ? n.Substring(0, 30) : n;
        }

        private static void InsertarDatosIdentificacion(dynamic ws, string carpeta, App util)
        {
            var ident = util.GetIdentification(App.DEFAULT_USER);
            if (ident == null) return;

            switch (carpeta)
            {
                case App.FolderNames.VERIFICADOS:
                    ws.Range["AH3"].Value = ident.Initials;
                    ws.Range["AM3"].Value = ident.FullName;
                    break;
                case App.FolderNames.RPS:
                    ws.Range["U2"].Value = ident.Initials;
                    ws.Range["Y2"].Value = ident.FullName;
                    break;
                case App.FolderNames.CDPS:
                    ws.Range["N3"].Value = ident.Initials;
                    ws.Range["Q3"].Value = ident.FullName;
                    break;
            }
        }

        private static void InsertarFirma(dynamic ws, string celda, string path)
        {
            try
            {
                dynamic img = ws.Pictures.Insert(path);
                dynamic target = ws.Range[celda];
                img.Left = target.Left; img.Top = target.Top; img.Width = 60; img.Height = 60; img.Placement = 1;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error insertando firma en {celda}", ex);
            }
        }

        private static void ConfigurarPagina(dynamic ws, bool esVertical = false)
        {
            App.ConfigurarPagina(ws, esVertical);
        }

        public static string GetCategoria(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return "XXX";
            string textoMayus = texto.ToUpper();

            // Lógica simplificada con pares Clave-Valor implícitos en orden de prioridad
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

        /// <summary>
        /// Obtiene la ruta local sincronizada en OneDrive para el libro activo.
        /// Decisión D-9: Se eliminó el parche hardcodeado de reemplazo de directorio.
        /// </summary>
        /// <param name="activeWorkbook">Libro activo de Excel.</param>
        /// <param name="util">Instancia de la aplicación.</param>
        /// <returns>Ruta local o cadena vacía si ocurre un error.</returns>
        public static string ObtenerRutaLocal(dynamic activeWorkbook, App util)
        {
            try
            {
                string path = util.GetOneDriveLocalPath(activeWorkbook.Path);
                LoggerHelper.Info($"Ruta local resuelta: {path}");
                return path ?? "";
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error obteniendo ruta local: {ex.Message}");
                return "";
            }
        }

        public static void CrearCarpetaSiNoExiste(string path)
        {
            if (!string.IsNullOrEmpty(path) && !Directory.Exists(path)) Directory.CreateDirectory(path);
        }
    }
}