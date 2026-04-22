using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ConversorPDF.XLL.FormProgramador;
using ConversorPDF.XLL.UI;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.FormAcademic;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Collections.Concurrent;

namespace ConversorPDF.XLL
{
    /// <summary>
    /// Controlador del Ribbon ESAP. Conecta los botones de Excel con la lógica C# existente.
    /// </summary>
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        // Caché concurrente para optimizar el load del XML y mitigar Cache MISS reiterados.
        private static readonly ConcurrentDictionary<string, string> _ribbonCache = new ConcurrentDictionary<string, string>();
        // ------------------------------------------------------------------
        // Carga del XML del Ribbon (embebido como recurso)
        // ------------------------------------------------------------------
        public override string GetCustomUI(string ribbonID)
        {
            try
            {
                // FIX-2.1: Implementar caché persistente a nivel de AppDomain para evitar 40+ Cache MISS
                const string cacheKey = "ConversorPDF.RibbonXML.Cache";
                string cachedXml = AppDomain.CurrentDomain.GetData(cacheKey) as string;
                
                if (!string.IsNullOrEmpty(cachedXml))
                {
                    // LoggerHelper.Info($"[PERF] [RibbonController] Cache HIT (AppDomain) para {ribbonID}");
                    return cachedXml;
                }

                var sw = System.Diagnostics.Stopwatch.StartNew();
                
                return _ribbonCache.GetOrAdd(ribbonID, id => {
                    LoggerHelper.Info($"[PERF] [RibbonController] Cache MISS para {id}. Cargando recurso...");
                    string xml = GetResourceText("ConversorPDF.XLL.Ribbon.Ribbon.xml");
                    
                    if (string.IsNullOrEmpty(xml))
                    {
                        LoggerHelper.Warning("GetCustomUI: El XML del Ribbon está vacío o no se encontró.");
                        return string.Empty;
                    }

                    // Persistir en el AppDomain para que otras instancias/AppDomains lo operen
                    AppDomain.CurrentDomain.SetData(cacheKey, xml);

                    sw.Stop();
                    LoggerHelper.Info($"[PERF] [RibbonController] XML cargado y persistido en {sw.ElapsedMilliseconds}ms para {id}");
                    return xml;
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en GetCustomUI", ex);
                return string.Empty;
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI) { }

        // ------------------------------------------------------------------
        // Control de Visibilidad dinámico
        // ------------------------------------------------------------------
        public bool GetButtonVisible(IRibbonControl control)
        {
            try
            {
                // Por defecto, nada es visible (Whitelist approach)
                bool visible = false;

                string usuario = App.DEFAULT_USER;
                if (string.IsNullOrEmpty(usuario)) return false;

                // PERFIL CONTADOR -------------------------------------------------------------------------
                if (usuario == "ALEXANDER.MONTANO")
                {
                    if (control.Id == "Bt1" || control.Id == "Bt11")
                        visible = true;
                }
                // PERFIL PAGADOR -------------------------------------------------------------------------
                if (usuario == "ANTOVILL")
                {
                    if (control.Id == "Bt1" || control.Id == "Bt11")
                        visible = true;
                }
                // PERFIL ADMINISTRADOR --------------------------------------------------------------------
                if (usuario == "LUIS.BECHARA" || usuario == "USUARIO")
                {
                    // Usuarios autorizados pueden ver los botones de unificación y procesos especiales
                    if (control.Id == "Bt1" || control.Id == "Bt2" || control.Id == "Bt3" || control.Id == "Bt4" || control.Id == "Bt5" || control.Id == "Bt6" || control.Id == "Bt7" || control.Id == "Bt8" || control.Id == "Bt9" || control.Id == "Bt10" || control.Id == "Bt11")
                        visible = true;
                }
                // PERFIL PAGADOR --------------------------------------------------------------------------

                // El botón de usuario (Btx) y el principal (Bt0) son visibles para todos los autorizados
                if (control.Id == "Bt0" || control.Id == "Btx")
                    visible = true;

                return visible;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error en GetButtonVisible para {control.Id}: {ex.Message}");
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Carga de imágenes personalizadas
        // ------------------------------------------------------------------
        public object GetImage(string imageId)
        {
            if (imageId == "esap")
            {
                return GetResourceImage("ConversorPDF.XLL.img.esap.bmp");
            }
            return null;
        }

        // ------------------------------------------------------------------
        // Botón: Convertir Obligaciones
        // ------------------------------------------------------------------        
        public void ConvertirPDF(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando ConvertirPDF");
            try
            {
                dynamic excelApp = ExcelDnaUtil.Application;
                dynamic wb = excelApp.ActiveWorkbook;
                dynamic ws = wb.ActiveSheet;
                var app = new App();

                bool isComisionado = app.Viaticos.ValidarFormatoComisionado(ws);
                LoggerHelper.Info($"isComisionado: {isComisionado}");


                if (isComisionado)
                {
                    LoggerHelper.Info("Abriendo FormProgramador en modo Individual");
                    using (var form = new ConversorPDF.XLL.FormProgramador.FormProgramador(TipoFormMode.verViaIndividual, wb, app.Viaticos))
                    {
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            LoggerHelper.Info($"Procesando Viáticos Individual: Inicio={form.SelectedFechaInicio:dd/MM/yyyy}, Fin={form.SelectedFechaFin:dd/MM/yyyy}, CDP={form.SelectedCDP}");
                            int vCount = 0;
                            using (var progress = new ProgressManager("Procesando Viáticos", 100))
                            {
                                app.Viaticos.ClearCache();
                                vCount = app.Viaticos.ProcesarIndividual(
                                    (Action<int, int>)((curr, total) => progress.AddProgress(100.0 / total)),
                                    form.SelectedFechaInicio,
                                    form.SelectedFechaFin,
                                    form.SelectedCDP);
                            }


                            if (vCount > 0)
                            {
                                MessageBox.Show(
                                    $"Procesamiento completado: {vCount} registro(s) procesado(s).",
                                    App.APP_NAME, MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                        }
                    }
                }
                }
                else
                {
                    if (MessageBox.Show(
                        "¿Desea convertir el archivo a PDF?",
                        App.APP_NAME,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        LoggerHelper.Info("Iniciando LoopSheetsFirmarAndSaveAsPDF");
                        int count = 0;
                        using (var progress = new ProgressManager("Convirtiendo a PDF", 100))
                        {
                            count = FirmaPDFsHelper.LoopSheetsFirmarAndSaveAsPDF(wb, app, (Action<int, int>)((curr, total) => progress.AddProgress(100.0 / total)));
                        }

                        if (count > 0)
                        {
                            MessageBox.Show(
                                $"Conversión finalizada: {count} archivo(s) generado(s).",
                                App.APP_NAME, MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en ConvertirPDF", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Unificar Soportes
        // ------------------------------------------------------------------
        public void UnificarObligacionContador(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando UnificarObligacionContador");
            if (MessageBox.Show(
                    "¿Desea unificar los soportes con las obligaciones?",
                    App.APP_NAME,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                var app = new App();
                app.Viaticos.ClearCache();
                dynamic wb = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook;

                int count = 0;
                using (var progress = new ProgressManager("Unificando PDF", 100))
                {
                    count = FirmaUnificarObHelper.UnificarObligacionContador(wb, "OB", app, (Action<int, int>)((curr, total) =>
                    {
                        progress.AddProgress(100.0 / total);
                    }));
                }

                LoggerHelper.Info($"UnificarObligacionContador count: {count}");
                if (count > 0)
                    MessageBox.Show(
                        $"Unificación completada: {count} archivo(s) procesados.",
                        App.APP_NAME, MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en UnificarObligacionContador", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Convertir Obligación (Contador)
        // ------------------------------------------------------------------
        public void OnConvertirObligacionContador(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando OnConvertirObligacionContador");
            if (MessageBox.Show(
                    "¿Desea convertir la hoja activa a PDF (Obligación)?",
                    App.APP_NAME,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                var app = new App();
                dynamic wb = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook;

                int count = 0;
                using (var progress = new ProgressManager("Convirtiendo Obligación", 100))
                {
                    // El mensaje de porcentaje se maneja internamente en ContadorService
                    count = app.ConvertirObligacionContador(wb);
                }

                LoggerHelper.Info($"ConvertirObligacionContador count: {count}");
                if (count > 0)
                    MessageBox.Show(
                        $"Conversión completada: {count} archivo(s) generado(s).",
                        App.APP_NAME, MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                else
                    MessageBox.Show(
                        "No se encontró el formato de Obligación en la hoja activa.",
                        App.APP_NAME, MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en OnConvertirObligacionContador", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Unificar Orden Pago
        // ------------------------------------------------------------------
        public void UnificarOrdenPago(IRibbonControl control)
        {
            MessageBox.Show("Función Unificar Orden de Pago (Próximamente)", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ------------------------------------------------------------------
        // Botón: Unificar Viáticos (RP)
        // ------------------------------------------------------------------
        public void ViaticosUnificarSoportes(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando UnificarViaticos");
            if (MessageBox.Show(
                    "¿Desea unificar los soportes de viáticos?",
                    App.APP_NAME,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                // ViaticosCommon.ClearCache();
                dynamic wb = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook;
                var app = new App();

                int count = 0;
                using (var progress = new ProgressManager("Unificando Viáticos", 100))
                {
                    count = app.Viaticos.UnificarSolicitudes(0, 0);
                }

                LoggerHelper.Info($"UnificarViaticos count: {count}");
                if (count > 0)
                    MessageBox.Show(
                        $"Unificación completada: {count} archivo(s) procesados.",
                        App.APP_NAME, MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                else
                    MessageBox.Show(
                        "No se encontraron archivos para unificar.",
                        App.APP_NAME, MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en UnificarViaticos", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Procesar Masivo Viáticos
        // ------------------------------------------------------------------
        public void ProcesarViaticosMasivo(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando ProcesarViaticosMasivo (vía FormProgramador)");
            try
            {
                var app = new App();
                app.Viaticos.ClearCache();
                dynamic excelApp = ExcelDnaUtil.Application;
                dynamic wb = excelApp.ActiveWorkbook;
                dynamic ws = wb.ActiveSheet;

                bool isComisionado = app.Viaticos.ValidarFormatoComisionado(ws);
                if (isComisionado)

                {
                    LoggerHelper.Info("Abriendo FormProgramador en modo Masivo");
                    using (var form = new ConversorPDF.XLL.FormProgramador.FormProgramador(TipoFormMode.verViaMasivo, wb, app.Viaticos))

                    {
                        form.ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show(
                        "Seleccione la hoja de Programación de viáticos.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en ProcesarViaticosMasivo", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Gestión Documental Viáticos
        // ------------------------------------------------------------------
        public void GestionDocumentalViaticos(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando GestionDocumentalViaticos");
            try
            {
                dynamic excelApp = ExcelDnaUtil.Application;
                dynamic wb = excelApp.ActiveWorkbook;
                dynamic ws = wb.ActiveSheet;

                if (ws.Name == "CONSULTAS")
                {
                    var app = new App();
                    using (var form = new ConversorPDF.XLL.FormProgramador.FormGestionDocumental(wb, app.Viaticos))
                    {
                        form.ShowDialog();
                    }
                }

                else
                {
                    MessageBox.Show(
                        "Seleccione la hoja de CONSULTAS para realizar la gestión documental.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en GestionDocumentalViaticos", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Generar Cumplidos (desde CONSULTAS)
        // ------------------------------------------------------------------
        public void ProcesarCumplidos(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando ProcesarCumplidos (vía FormProgramador)");
            try
            {
                var app = new App();
                app.Viaticos.ClearCache();
                dynamic excelApp = ExcelDnaUtil.Application;
                dynamic wb = excelApp.ActiveWorkbook;
                using (var form = new ConversorPDF.XLL.FormProgramador.FormProgramador(TipoFormMode.verViaCumplidos, wb, app.Viaticos))

                {
                    form.ShowDialog();
                }
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en ProcesarCumplidos", ex);
                MessageBox.Show(
                    $"Error al abrir el formulario de cumplidos: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Generar Fila Seleccionada
        // ------------------------------------------------------------------
        public void GenerarFilaSeleccionada(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando GenerarFilaSeleccionada");
            dynamic excelApp = ExcelDnaUtil.Application;
            try
            {
                using (ExcelOptimizeHelper.Optimize(excelApp))
                {
                    dynamic wb = excelApp.ActiveWorkbook;
                    dynamic ws = wb.ActiveSheet;
                    string sheetName = ws.Name;
                    var app = new App();

                    LoggerHelper.Info($"Hoja activa: {sheetName}");
                    int processed = 0;
                    using (var progress = new ProgressManager("Procesando Fila Seleccionada", 100))
                    {
                        processed = app.Viaticos.ProcesarFilaSeleccionada(
                            (Action<int, int>)((curr, total) => progress.AddProgress(100.0 / total)));
                    }

                    if (processed > 0)
                    {
                        MessageBox.Show(
                            $"Procesamiento completado: {processed} registro(s) procesado(s).",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        LoggerHelper.Info($"GenerarFilaSeleccionada finalizado exitosamente ({sheetName})");
                    }
                    else
                    {
                        MessageBox.Show(
                            "No se encontraron registros para procesar en la fila seleccionada o no se realizó ninguna acción.",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        LoggerHelper.Info("GenerarFilaSeleccionada finalizado sin procesar registros");
                    }
                }
            }
            catch (ViaticosFormatException ex)
            {
                LoggerHelper.Warning("Formato de libro no válido: " + ex.Message);
                MessageBox.Show(ex.Message, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en GenerarFilaSeleccionada", ex);
                MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }

        // ------------------------------------------------------------------
        // Botón: Generar Firma Base64
        // ------------------------------------------------------------------
        public void GenerarFirmaBase64(IRibbonControl control)
        {
            try
            {
                string outputTxt = FirmasGeneradorHelper.GenerarStringBase64DesdeImagen();
                if (!string.IsNullOrEmpty(outputTxt) && !outputTxt.StartsWith("ERROR"))
                    System.Windows.MessageBox.Show(
                        $"Firma generada y guardada en:\n{outputTxt}",
                        App.APP_NAME, System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Error: {ex.Message}", App.APP_NAME,
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Office365 UserManager (ahora usa Windows Identity)
        // ------------------------------------------------------------------
        public void Office365UserManager(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando Office365UserManager (Info Form)");
            try
            {
                using (var form = new FormInfo())
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en Office365UserManager", ex);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Buscar Informes
        // ------------------------------------------------------------------
        public void BuscarInformes(IRibbonControl control)
        {
            MessageBox.Show("Función Buscar Informes (Próximamente)", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ------------------------------------------------------------------
        // Botón: Leer Metadatos PDF
        // ------------------------------------------------------------------
        /// <summary>
        /// Callback del Ribbon para leer metadatos de archivos PDF seleccionados.
        /// </summary>
        /// <param name="control">Referencia al botón del Ribbon que dispara la acción.</param>
        public void LeerMetadatos(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando LeerMetadatos desde el Ribbon");
            try
            {
                PdfMetadataHelper.LeerMetadatosPDF();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en LeerMetadatos (Ribbon)", ex);
            }
        }

        /// <summary>
        /// Callback del Ribbon para abrir el generador de sesiones académicas.
        /// </summary>
        /// <param name="control">Referencia al botón del Ribbon.</param>
        public void SesionesAcademicas(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando SesionesAcademicas desde el Ribbon");
            try
            {
                // Mostramos el formulario de progreso mientras se inicializa el pesado motor académico
                using (var progress = new ProgressManager("Iniciando Sesiones Académicas...", 100))
                {
                    // El constructor del formulario ahora reporta su progreso al manager
                    using (var form = new FormSesionesAcademicas(progress))
                    {
                        // Cerramos el progreso justo antes de mostrar el formulario modal
                        progress.Close();
                        form.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en SesionesAcademicas (Ribbon)", ex);
            }
        }

        // ------------------------------------------------------------------
        // Helper: leer recurso embebido
        // ------------------------------------------------------------------
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return string.Empty;
                using (StreamReader reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        private static System.Drawing.Bitmap GetResourceImage(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                // Clonar la imagen para no mantener el stream abierto y evitar GDI+ leaks
                using (var originalImage = System.Drawing.Image.FromStream(stream))
                {
                    return new System.Drawing.Bitmap(originalImage);
                }
            }
        }
    }
}

