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
using ConversorPDF.XLL.Academic.UI.Forms;
using ConversorPDF.XLL.Ribbon;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL
{
    /// <summary>
    /// Controlador del Ribbon ESAP. Conecta los botones de Excel con la lógica C# existente.
    /// </summary>
    [ComVisible(true)]
    [Guid("08c7ef5e-8e08-408a-8442-010cd1d6654e")]
    [ProgId("ConversorPDF.XLL.RibbonController")]
    public class RibbonController : ExcelRibbon
    {
        static RibbonController()
        {
            LoggerHelper.Info("RibbonController: Constructor estático invocado.");
        }

        public RibbonController()
        {
            LoggerHelper.Info("RibbonController: Constructor de instancia invocado por Excel-DNA.");
        }

        private static readonly Lazy<string> _ribbonXml = new Lazy<string>(() =>
        {
            string xml = GetResourceText("ConversorPDF.XLL.Ribbon.Ribbon.xml");

            if (string.IsNullOrEmpty(xml))
            {
                LoggerHelper.Warning("GetCustomUI: El XML del Ribbon está vacío o no se encontró.");
                return string.Empty;
            }

            return xml;
        });

        // ------------------------------------------------------------------
        // Carga del XML del Ribbon (embebido como recurso)
        // ------------------------------------------------------------------
        /// <summary>
        /// Método principal que Excel invoca para solicitar el XML de la cinta.
        /// Devuelve el contenido de Ribbon.xml de forma segura y loguea si está vacío.
        /// </summary>
        public override string GetCustomUI(string ribbonID)
        {
            try
            {
                string xml = _ribbonXml.Value;
                if (string.IsNullOrEmpty(xml))
                {
                    LoggerHelper.Warning("GetCustomUI: XML del Ribbon vacío. La cinta no se mostrará.");
                }
                else
                {
                    LoggerHelper.Info($"GetCustomUI: Retornando XML válido de {xml.Length} caracteres para ribbonID={ribbonID}.");
                }
                return xml;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en GetCustomUI", ex);
                return string.Empty;
            }
        }

        private static IRibbonUI _ribbon;

        public void Ribbon_Load(IRibbonUI ribbonUI) 
        { 
            _ribbon = ribbonUI;
        }

        public static void InvalidateRibbon()
        {
            lock (_permisosLock)
            {
                // Fase 22: Asignar nueva instancia en lugar de .Clear() 
                // para evitar mutar referencias compartidas leídas concurrentemente
                _cachedPermisos = new System.Collections.Generic.HashSet<RibbonButtonId>();
                _cachedUsername = null;
                _lastProfileCheck = DateTime.MinValue;
                _lastProfileCheckFailed = false; // Fase 18
            }
            try
            {
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error invalidando Ribbon: {ex.Message}");
            }
        }

        private static System.Collections.Generic.HashSet<RibbonButtonId> _cachedPermisos = new System.Collections.Generic.HashSet<RibbonButtonId>();
        private static string _cachedUsername = null;
        private static DateTime _lastProfileCheck = DateTime.MinValue;
        private static bool _lastProfileCheckFailed = false; // Fase 18
        private static readonly object _permisosLock = new object();

        private System.Collections.Generic.HashSet<RibbonButtonId> ParsePermisos(string permisosStr)
        {
            var newPermisos = new System.Collections.Generic.HashSet<RibbonButtonId>();
            var perms = permisosStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in perms)
            {
                if (Enum.TryParse(p.Trim(), out RibbonButtonId buttonId))
                    newPermisos.Add(buttonId);
            }
            return newPermisos;
        }

        /// <summary>
        /// Obtiene los permisos de la base de datos de manera optimizada, manteniendo
        /// una caché por 5 minutos para evitar saturar SQLite durante el refresco del Ribbon.
        /// </summary>
        private System.Collections.Generic.HashSet<RibbonButtonId> GetCachedPermisos(string username, bool allowDbAccess = true)
        {
            // Fase 19: Fast path bajo lock
            lock (_permisosLock)
            {
                double cacheMinutes = _lastProfileCheckFailed ? 0.5 : 5;
                if (_cachedUsername == username && (DateTime.Now - _lastProfileCheck).TotalMinutes < cacheMinutes)
                    return _cachedPermisos;
            }

            // Si no se permite acceso a BD, intentar solo desde la cookie
            if (!allowDbAccess)
            {
                string cookiePermisos = CookiePermisosHelper.LoadCookie(username);
                var parsed = !string.IsNullOrEmpty(cookiePermisos)
                    ? ParsePermisos(cookiePermisos)
                    : new System.Collections.Generic.HashSet<RibbonButtonId>();

                lock (_permisosLock)
                {
                    // Re-validar: otro thread pudo haber actualizado con datos de BD
                    double cacheMinutes = _lastProfileCheckFailed ? 0.5 : 5;
                    if (_cachedUsername == username && (DateTime.Now - _lastProfileCheck).TotalMinutes < cacheMinutes)
                        return _cachedPermisos;

                    _cachedPermisos = parsed;
                    _cachedUsername = username;
                    _lastProfileCheck = DateTime.Now;
                    _lastProfileCheckFailed = true; // Cache corta para re-validar contra BD pronto
                }
                return parsed;
            }

            // Lectura DB fuera del lock
            var newPermisos = new System.Collections.Generic.HashSet<RibbonButtonId>();
            bool failed = false;
            try
            {
                // Fase 23: Usar GetContext() (Singleton ya migrado) en vez de GetContextForPath
                // que puede ejecutar migraciones costosas en el primer render del Ribbon.
                // GetContext() NO se envuelve en using — es un Singleton global, no descartable aquí.
                // Tampoco se llama MarkAsReadOnly() porque afectaría el Singleton global.
                var context = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext();
                var conn = context.Open();
                
                var repo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.AcademicRepository(conn);
                var result = repo.GetPermisosPorUsuario(username);

                if (result.IsSuccess)
                {
                    CookiePermisosHelper.SaveCookie(username, result.Value ?? string.Empty);
                    if (!string.IsNullOrEmpty(result.Value))
                    {
                        newPermisos = ParsePermisos(result.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Warning($"Error obteniendo permisos para Ribbon: {ex.Message}");
                failed = true;
                
                // Fallback a la cookie si falla la BD
                string cookiePermisos = CookiePermisosHelper.LoadCookie(username);
                if (!string.IsNullOrEmpty(cookiePermisos))
                {
                    newPermisos = ParsePermisos(cookiePermisos);
                    // Mantener failed = true para forzar un reintento a BD rápido (0.5 min)
                }
            }

            // Actualizar caché bajo lock
            lock (_permisosLock)
            {
                // Fase 20: Revalidar caché tras lectura DB
                double cacheMinutes = _lastProfileCheckFailed ? 0.5 : 5;
                if (_cachedUsername == username && (DateTime.Now - _lastProfileCheck).TotalMinutes < cacheMinutes)
                    return _cachedPermisos; // Otro thread ya actualizó

                _cachedPermisos = newPermisos;
                _cachedUsername = username;
                _lastProfileCheck = DateTime.Now;
                _lastProfileCheckFailed = failed; // Fase 19
                return _cachedPermisos;
            }
        }

        // ------------------------------------------------------------------
        // Control de Visibilidad dinámico
        // ------------------------------------------------------------------
        public bool GetButtonVisible(IRibbonControl control)
        {
            try
            {
                string usuario = App.DEFAULT_USER;
                if (string.IsNullOrEmpty(usuario)) return false; // Fase 16: fail-closed sin usuario

                bool dbReady = DatabaseBootstrapper.WasInitialized;
                var permisos = GetCachedPermisos(usuario, allowDbAccess: dbReady);

                if (!Enum.TryParse(control.Id, out RibbonButtonId buttonId))
                {
                    LoggerHelper.Warning($"ID de botón no reconocido: {control.Id}");
                    return false; // Fase 16: fail-closed para IDs no mapeados
                }

                // Fase 15: Si no hay permisos cargados (usuario sin perfil), ocultar por defecto para respetar seguridad
                if (permisos.Count == 0) return false;

                return permisos.Contains(buttonId);
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error en GetButtonVisible para {control.Id}: {ex.Message}");
                return false; // Fase 17: fail-closed en error crítico
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
                // Defensa en profundidad: verificar permiso en el backend aunque el botón ya esté oculto.
                // EnsureInitialized es idempotente; si la BD ya arrancó en background es un no-op.
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.ConvertirPDF))
                {
                    MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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
            // Defensa en profundidad: verificar permiso antes de mostrar el diálogo de confirmación.
            // EnsureInitialized se envuelve en su propio try/catch (separación de responsabilidades):
            // los errores de inicialización de BD son distintos a los errores de la operación de negocio.
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[UnificarObligacionContador] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.UnificarOB))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            // Defensa en profundidad: verificar permiso antes de mostrar el diálogo de confirmación.
            // EnsureInitialized envuelto en su propio try/catch (init ≠ operación de negocio).
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[OnConvertirObligacionContador] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.ConvertirObligacion))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            LoggerHelper.Info("Iniciando UnificarOrdenPago");

            // 1. Inicialización de base de datos en try/catch independiente
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[UnificarOrdenPago] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Verificación de permisos de usuario (RibbonButtonId.UnificarOP)
            if (!HasPermission(RibbonButtonId.UnificarOP))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 3. Diálogo de confirmación
            if (MessageBox.Show(
                    "¿Desea unificar los soportes con las órdenes de pago?",
                    App.APP_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            // 4. Ejecución del proceso de unificación con barra de progreso
            try
            {
                var app = new App();
                app.Viaticos.ClearCache();
                dynamic wb = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook;

                int count = 0;
                using (var progress = new ProgressManager("Unificando PDF", 100))
                {
                    count = UnificarOrdenPagoHelper.UnificarOrdenPago(wb, app, "OP",
                        (Action<int, int>)((curr, total) =>
                        {
                            progress.AddProgress(100.0 / total);
                        }));
                }

                LoggerHelper.Info($"UnificarOrdenPago count: {count}");
                if (count == -1)
                {
                    MessageBox.Show("No se pudo resolver la ruta local del libro. Guarde el archivo en OneDrive e intente de nuevo.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (count == -2)
                {
                    MessageBox.Show("No hay boletines ni soportes activos para unificar.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (count == -3)
                {
                    MessageBox.Show("El tipo de documento especificado no es válido.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (count == -4)
                {
                    MessageBox.Show("Error interno: parámetros nulos al invocar la unificación.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (count > 0)
                {
                    MessageBox.Show($"Unificación completada: {count} archivo(s) procesados.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No se encontraron archivos para unificar.",
                        App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en UnificarOrdenPago", ex);
                MessageBox.Show($"Error: {ex.Message}", App.APP_NAME,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Unificar Viáticos (RP)
        // ------------------------------------------------------------------
        public void ViaticosUnificarSoportes(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando UnificarViaticos");
            // Defensa en profundidad: verificar permiso antes de mostrar el diálogo de confirmación.
            // EnsureInitialized envuelto en su propio try/catch (init ≠ operación de negocio).
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[ViaticosUnificarSoportes] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.UnificarViaticos))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
                // Defensa en profundidad: verificar permiso en el backend.
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.MasivoViaticos))
                {
                    MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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
                // Defensa en profundidad: verificar permiso en el backend.
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.GestionDocumental))
                {
                    MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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
                // Defensa en profundidad: verificar permiso en el backend.
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.GenerarCumplidos))
                {
                    MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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
            // Defensa en profundidad: verificar permiso en el backend.
            // EnsureInitialized en su propio try/catch: si falla la init no debe afectar
            // al bloque finally { excelApp.ScreenUpdating = true } del cuerpo principal.
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[GenerarFilaSeleccionada] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.GenerarFilaSeleccionada))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            // Defensa en profundidad: verificar permiso en el backend.
            // ProcesarFirma es el RibbonButtonId correspondiente a este callback.
            // EnsureInitialized en su propio try/catch (errores de init ≠ errores de operación).
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[GenerarFirmaBase64] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.ProcesarFirma))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            // Defensa en profundidad: verificar permiso en el backend.
            // EnsureInitialized en su propio try/catch (errores de init ≠ errores de operación).
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
            }
            catch (Exception exInit)
            {
                LoggerHelper.Error("[LeerMetadatos] Error al inicializar la BD", exInit);
                MessageBox.Show($"No se pudo inicializar la base de datos.\n{exInit.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!HasPermission(RibbonButtonId.Metadatos))
            {
                MessageBox.Show("No tiene permisos para esta función.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
        /// Muestra un InputBox personalizado para solicitar el ID de vinculación.
        /// Devuelve el ID numérico, 0 si se deja en blanco/0 (Nuevo Registro), o null si el usuario cancela.
        /// </summary>
        private int? PromptForVinculacionId()
        {
            using (var form = new Form())
            {
                form.Width = 360;
                form.Height = 150;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.Text = "Sesiones Académicas";
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var lblInfo = new Label()
                {
                    Left = 20,
                    Top = 15,
                    Width = 310,
                    Height = 40,
                    Text = "Ingrese el ID de vinculación a revisar\n(Deje en blanco o escriba 0 para Nuevo Registro):"
                };
                var txtId = new TextBox() { Left = 20, Top = 60, Width = 150 };
                var btnOk = new Button() { Text = "Aceptar", Left = 180, Top = 58, Width = 70, DialogResult = DialogResult.OK };
                var btnCancel = new Button() { Text = "Cancelar", Left = 260, Top = 58, Width = 70, DialogResult = DialogResult.Cancel };

                form.Controls.Add(lblInfo);
                form.Controls.Add(txtId);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    string input = txtId.Text.Trim();
                    if (string.IsNullOrEmpty(input) || input == "0")
                    {
                        return 0; // Blanco o 0 es nuevo registro
                    }

                    if (int.TryParse(input, out int id) && id > 0)
                    {
                        return id;
                    }
                    else
                    {
                        MessageBox.Show("El ID ingresado no es válido. Debe ser un número entero mayor a 0.", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null; // Error de validación, cancela
                    }
                }

                return null; // Usuario canceló
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
                // Defensa en profundidad: verificar permiso antes de mostrar cualquier diálogo.
                // EnsureInitialized es idempotente; el hilo STA de AutoOpen puede haberla completado ya.
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.SesionesAcademicas))
                {
                    MessageBox.Show("No tiene permisos para Sesiones Académicas.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Solicitar ID de vinculación al usuario (solo si tiene permiso)
                int? idVinculacion = PromptForVinculacionId();
                if (idVinculacion == null)
                {
                    LoggerHelper.Info("El usuario canceló la apertura de Sesiones Académicas o ingresó un ID inválido.");
                    return;
                }

                using (var session = new AcademicDbSession())
                {
                    // Mostramos el formulario de progreso mientras se inicializa el pesado motor académico
                    using (var progress = new ProgressManager("Iniciando Sesiones Académicas...", 100))
                    {
                        // El constructor del formulario ahora reporta su progreso al manager
                        // Se pasa el ID ingresado por el usuario (0 = Nuevo, >0 = Revisar)
                        using (var form = new FormSesionesAcademicas(() => session.Connection, progress, idVinculacion.Value))
                        {
                            // Cerramos el progreso justo antes de mostrar el formulario modal
                            progress.Close();
                            form.ShowDialog();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en SesionesAcademicas (Ribbon)", ex);
            }
        }

        /// <summary>
        /// Abre el listado de vinculaciones académicas en el modo especificado (Producción o Staging AG).
        /// </summary>
        private void AbrirVinculacionesConModo(IRibbonControl control, AgModoDatos modo)
        {
            LoggerHelper.Info($"Iniciando VinculacionesAcademicas desde el Ribbon (Modo: {modo?.Nombre ?? "PRODUCCIÓN"})");
            try
            {
                if (!DatabaseBootstrapper.WasInitialized)
                {
                    // Mostrar progreso solo durante la primera carga pesada.
                    using (var progress = new ProgressManager("Iniciando módulo académico...", 4))
                    {
                        DatabaseBootstrapper.EnsureInitialized(step =>
                        {
                            if (step == "__CLOSE_PROGRESS__")
                            {
                                progress.CloseAndWait(500); // Esperar cierre real antes de SyncContext
                                return;
                            }
                            progress.AddProgress(1);
                        });
                    }
                }

                // Defensa en profundidad: verificar permiso después de que la BD esté lista.
                bool hasAccess = HasPermission(RibbonButtonId.Vinculaciones) || (modo?.EsModoAg == true && HasPermission(RibbonButtonId.AgHorariosAcademicos));
                if (!hasAccess)
                {
                    MessageBox.Show("No tiene permisos para Vinculaciones Académicas.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Pilar B: Usar patrón Shared Context para todo el ciclo de vida del formulario
                using (var session = new AcademicDbSession())
                {
                    session.MarkAsReadOnly(); // FormVinculaciones es solo lectura

                    // Pilar B: Usar conexión compartida de la sesión
                    var queryRepo = new VinculacionQueryRepository(() => session.Connection, modo);

                    var sedeRepo = new SedeLookupRepository(() => session.Connection);
                    var terceroRepo = new TerceroLookupRepository(() => session.Connection);
                    var periodoRepo = new PeriodoLookupRepository(() => session.Connection);
                    var programaRepo = new ProgramaLookupRepository(() => session.Connection);

                    var listService = new VinculacionListService(queryRepo, sedeRepo, terceroRepo, periodoRepo, programaRepo);

                    using (var form = new FormVinculaciones(listService, sedeRepo, terceroRepo, () => session.Connection, ExcelDnaUtil.Application, modo))
                    {
                        form.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en VinculacionesAcademicas (Ribbon)", ex);
                MessageBox.Show($"Error al abrir vinculaciones: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Callback del Ribbon para abrir el listado de vinculaciones académicas en modo Determinista (Producción).
        /// </summary>
        public void VinculacionesAcademicas(IRibbonControl control) => AbrirVinculacionesConModo(control, AgModoDatos.Determinista);

        /// <summary>
        /// Callback del Ribbon para abrir el listado de vinculaciones académicas en Modo AG (Staging).
        /// </summary>
        public void VinculacionesModoAg(IRibbonControl control) => AbrirVinculacionesConModo(control, AgModoDatos.Ag);

        /// <summary>
        /// Callback del Ribbon para generar un backup de la base de datos académica en SQLite.
        /// </summary>
        /// <param name="control">Referencia al botón del Ribbon.</param>
        public void BackupsAcademicas(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando BackupsAcademicas desde el Ribbon");
            try
            {
                DatabaseBootstrapper.EnsureInitialized(); // Asegura que el vault exista y esté migrado

                // Defensa en profundidad: verificar permiso antes de mostrar el SaveFileDialog.
                if (!HasPermission(RibbonButtonId.Backups))
                {
                    MessageBox.Show("No tiene permisos para generar backups.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (var sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Base de Datos SQLite (*.db)|*.db";
                    sfd.Title = "Guardar Backup de Sesiones Académicas";
                    sfd.FileName = $"Backup_AcademicData_{DateTime.Now:yyyyMMdd_HHmm}.db";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // Extraer el vault a un archivo temporal (TempDbPath)
                        // Usamos GetContextForPath para no afectar al SharedContext (Singleton)
                        using (var context = SqliteHelper.GetContextForPath(null))
                        {
                            context.Open(); // Extraer vault -> crea TempDbPath
                            context.MarkAsReadOnly(); // Prevenir cualquier persistencia accidental

                            // Realizar backup página a página asegurando consistencia del WAL
                            context.BackupTo(sfd.FileName);
                        }

                        LoggerHelper.Info($"Backup exportado exitosamente a: {sfd.FileName}");
                        MessageBox.Show("Backup exportado exitosamente.", "Backup Completado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LoggerHelper.Info("El usuario canceló la exportación del backup.");
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en BackupsAcademicas (Ribbon)", ex);
                MessageBox.Show($"Ocurrió un error al intentar crear el backup:\n{ex.Message}", "Error de Backup", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper Fase 20
        private bool HasPermission(RibbonButtonId buttonId)
        {
            try
            {
                string usuario = App.DEFAULT_USER;
                if (string.IsNullOrEmpty(usuario)) return false;
                // Fase 21: Reusar caché en vez de abrir contexto nuevo
                return GetCachedPermisos(usuario).Contains(buttonId);
            }
            catch (Exception ex) { LoggerHelper.Warning($"HasPermission: {ex.Message}"); }
            return false;
        }

        // ------------------------------------------------------------------
        // Botón: Gestión Usuarios
        // ------------------------------------------------------------------
        public void AbrirUsuarios(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando Gestion de Usuarios desde el Ribbon");
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.GestionUsuarios)) // Fase 20
                {
                    System.Windows.Forms.MessageBox.Show("No tiene permisos para gestionar usuarios.", App.APP_NAME, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }
                using (var session = new AcademicDbSession())
                {
                    using (var form = new FormUsuarios(() => session.Connection))
                    {
                        form.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en AbrirUsuarios (Ribbon)", ex);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Permisos (Gestión de Roles)
        // ------------------------------------------------------------------
        public void AbrirPermisos(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando AbrirPermisos desde el Ribbon");
            try
            {
                // Asegurar que la base de datos esté lista (y migrada) antes de intentar abrir
                // un formulario que potencialmente requiera leer de la tabla 'permisos' o 'usuarios'
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.Permisos)) // Fase 20
                {
                    System.Windows.Forms.MessageBox.Show("No tiene permisos para gestionar perfiles.", App.APP_NAME, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }
                using (var form = new FormPermisos())
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en AbrirPermisos", ex);
                MessageBox.Show($"Error: {ex.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Botón: Optimizar Horarios (Algoritmo Genético)
        // ------------------------------------------------------------------
        public void AbrirAgHorariosAcademicos(IRibbonControl control)
        {
            LoggerHelper.Info("Iniciando AbrirAgHorariosAcademicos desde el Ribbon");
            try
            {
                DatabaseBootstrapper.EnsureInitialized();
                if (!HasPermission(RibbonButtonId.AgHorariosAcademicos))
                {
                    MessageBox.Show("No tiene permisos para el Optimizador de Horarios.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (var session = new AcademicDbSession())
                {
                    using (var form = new AgFormGeneradorHorarios(() => session.Connection))
                    {
                        form.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en AbrirAgHorariosAcademicos (Ribbon)", ex);
                MessageBox.Show($"Error al abrir Optimizador de Horarios: {ex.Message}", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Helper: leer recurso embebido
        // ------------------------------------------------------------------
        /// <summary>
        /// Lee un archivo de recurso embebido dentro de la DLL.
        /// Retorna un string vacío y registra un warning si el recurso no existe.
        /// </summary>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    LoggerHelper.Warning($"GetCustomUI: Recurso '{resourceName}' no encontrado en el ensamblado.");
                    return string.Empty;
                }
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

