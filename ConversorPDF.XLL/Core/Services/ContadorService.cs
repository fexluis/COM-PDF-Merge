using System;
using System.IO;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Helpers;
using System.Globalization;

namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio para el procesamiento de obligaciones del contador.
    /// </summary>
    public class ContadorService : IContadorService
    {
        private readonly IUserService _userService;
        private readonly IPdfMetadataService _metadataService;
        private readonly IPdfService _pdfService;
        private readonly IFirmaExtractService _firmaExtractService;

        public ContadorService(
            IUserService userService, 
            IPdfMetadataService metadataService, 
            IPdfService pdfService,
            IFirmaExtractService firmaExtractService)
        {
            _userService = userService;
            _metadataService = metadataService;
            _pdfService = pdfService;
            _firmaExtractService = firmaExtractService;
        }


        /// <inheritdoc/>
        public int ConvertirObligacionContador(dynamic activeWorkbook, string userFirma, Action<int, int> progressCallback = null)
        {
            if (activeWorkbook == null) return 0;

            dynamic activeSheet = activeWorkbook.ActiveSheet;
            if (activeSheet.Range["B14"].Value?.ToString() != "REGISTRO PRESUPUESTAL DE OBLIGACION.")
            {
                return 0;
            }

            string localPath = ObtenerRutaLocal(activeWorkbook);
            if (string.IsNullOrEmpty(localPath)) return 0;

            string folderObligaciones = Path.Combine(localPath, "Obligaciones");
            if (!Directory.Exists(folderObligaciones)) Directory.CreateDirectory(folderObligaciones);

            string tempPath = Path.Combine(Path.GetTempPath(), $"Firma_{Guid.NewGuid():N}.png");
            try
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(userFirma))
                        throw new Exception("El usuario no tiene una firma configurada.");

                    string cleanBase64 = FirmaPDFsHelper.LimpiarPrefijoBase64(userFirma);
                    byte[] bytes = Convert.FromBase64String(cleanBase64);
                    File.WriteAllBytes(tempPath, bytes);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("No se pudo procesar la firma del usuario.", ex);
                    throw new Exception("No se pudo procesar la firma del usuario. Verifique que esté correctamente configurada en la base de datos.");
                }

                dynamic excelApp = activeWorkbook.Application;
                using (ExcelOptimizeHelper.Optimize(excelApp))
                {
                    int filesCount = 0;
                    int totalSheets = activeWorkbook.Worksheets.Count;
                    int currentSheet = 0;

                    foreach (dynamic ws in activeWorkbook.Worksheets)
                    {
                        currentSheet++;
                        progressCallback?.Invoke(currentSheet, totalSheets);

                        dynamic rngFirma = ws.Range["A1:BH300"].Find("FIRMA(S) RESPONSABLE(S)");
                        string firmaPos = (rngFirma != null) ? "J" + (rngFirma.Row - 3) : "J57";

                        var datos = _firmaExtractService.ObligacionData(ws);

                        InsertarFirma(ws, firmaPos, tempPath);
                        int rowNum = (rngFirma != null) ? (int)rngFirma.Row : 60;
                        ConfigurarPagina(ws, rowNum, excelApp);

                        string fileName = datos.xUnifiedFileName.Length > 60 ? datos.xUnifiedFileName.Substring(0, 60) : datos.xUnifiedFileName;
                        string fullPdfPath = Path.Combine(folderObligaciones, fileName + ".pdf");

                        try
                        {
                            ws.ExportAsFixedFormat(0, fullPdfPath);

                            // Inyectar Metadatos
                            DateTime? fechaParseada = null;
                            if (DateTime.TryParse(datos.xFechaStr, out DateTime result)) fechaParseada = result;

                            string fullUser = _userService.GetUsuarioWindows();
                            var parts = fullUser.Split('\\');
                            string cleanUser = parts.Length > 1 ? parts[1].ToUpper() : fullUser.ToUpper();

                            _metadataService.GuardarMetadatos(fullPdfPath, new Metadata
                            {
                                Titulo = fileName,
                                Autor = cleanUser,
                                Asunto = datos.xObjeto,
                                PalabrasClave = $"{datos.xIdentificacion}, {datos.xTercero}, {datos.xDep}",
                                Fecha = fechaParseada,
                                MetaDocumento = TipoMetadata.OB,
                                RP = datos.xRP,
                                OB = datos.xNumero,
                                Tipo = datos.xCategoria
                            });
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Error($"Error exportando PDF o guardando metadatos: {fullPdfPath}", ex);
                        }

                        filesCount++;
                    }

                    return filesCount;
                }
            }
            finally
            {
                if (File.Exists(tempPath)) try { File.Delete(tempPath); } catch { }
            }
        }


        /// <summary>
        /// Obtiene la ruta local sincronizada en OneDrive para el libro activo.
        /// Decisión D-9: Se eliminó el parche hardcodeado de reemplazo de directorio.
        /// </summary>
        private string ObtenerRutaLocal(dynamic activeWorkbook)
        {
            try
            {
                string path = OneDriveHelper.GetOneDriveLocalPath(activeWorkbook.Path);
                return path ?? "";
            }
            catch { return ""; }
        }

        /// <summary>
        /// Inserta una imagen de firma en la hoja de Excel en una posición específica.
        /// </summary>
        /// <param name="ws">Hoja de cálculo activa.</param>
        /// <param name="firmaPos">Dirección de celda (ej: "J57").</param>
        /// <param name="tempPath">Ruta al archivo de imagen temporal.</param>
        private void InsertarFirma(dynamic ws, string firmaPos, string tempPath)
        {
            try
            {
                dynamic img = ws.Pictures.Insert(tempPath);
                dynamic target = ws.Range[firmaPos];
                img.Left = target.Left; img.Top = target.Top; img.Width = 60; img.Height = 60; img.Placement = 1;
            }
            catch (Exception ex) { LoggerHelper.Error($"Error insertando firma en {firmaPos}", ex); }
        }

        /// <summary>
        /// Aplica configuraciones de impresión a la hoja de Excel para la exportación a PDF.
        /// </summary>
        /// <param name="ws">Hoja de cálculo.</param>
        /// <param name="rowNum">Número de fila límite para el área de impresión.</param>
        /// <param name="excelApp">Instancia de la aplicación Excel.</param>
        private void ConfigurarPagina(dynamic ws, int rowNum, dynamic excelApp)
        {
            try
            {
                dynamic ps = ws.PageSetup;
                double m = excelApp.InchesToPoints(0.1);
                ps.LeftMargin = m; ps.RightMargin = m; ps.TopMargin = m; ps.BottomMargin = m;
                ps.CenterHorizontally = true; ps.CenterVertically = true;
                ps.Zoom = false; ps.FitToPagesTall = 1; ps.FitToPagesWide = 1;
                ps.Orientation = rowNum > 62 ? 1 : 2;
            }
            catch { }
        }


    }
}
