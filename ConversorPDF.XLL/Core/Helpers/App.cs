using System;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Enums;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
[assembly: AssemblyVersion(ConversorPDF.XLL.Core.Helpers.App.APP_VERSION)]
[assembly: AssemblyFileVersion(ConversorPDF.XLL.Core.Helpers.App.APP_VERSION)]

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Clase principal de la aplicación ConversorPDF.
    /// Proporciona la lógica central para la integración con Office y servicios de PDF.
    /// </summary>
    [ComVisible(true)]
    [Guid("4F97F332-3B6C-4B6E-8B7C-5AF9F4EEA4C9")]
    [ProgId("ConversorPDF.App")]
    [ClassInterface(ClassInterfaceType.None)]
    public class App : IApp
    {
        private readonly IUserService _userService;
        private readonly IPdfService _pdfService;
        private readonly IViaticosService _viaticosService;
        private readonly IContadorService _contadorService;

        public App() : this(
            new UserService(), 
            new PdfService(), 
            CreateViaticosService(),
            new ContadorService(new UserService(), new PdfMetadataService(), new PdfService(), new FirmaExtractService())) { }

        private static IViaticosService CreateViaticosService()
        {
            var excelRepo = new ExcelRepository();
            var userSvc = new UserService();
            var viatRepo = new ViaticosRepository(excelRepo);
            var metadataSvc = new PdfMetadataService();
            
            return new ViaticosService(
                excelRepo, 
                userSvc, 
                viatRepo,
                new ViaticosIndividualProcessor(excelRepo, viatRepo, userSvc, metadataSvc),
                new ViaticosMasivoProcessor(excelRepo, viatRepo),
                new ViaticosCumplidosProcessor(excelRepo, viatRepo, userSvc, metadataSvc),
                new ViaticosGestionProcessor(excelRepo, viatRepo, new PdfToTextoHelper(), metadataSvc)
            );
        }

        public App(IUserService userService, IPdfService pdfService, IViaticosService viaticosService, IContadorService contadorService)
        {
            _userService = userService;
            _pdfService = pdfService;
            _viaticosService = viaticosService;
            _contadorService = contadorService;
        }

        /// <summary>
        /// Acceso al servicio de procesamiento de viáticos.
        /// </summary>
        public IViaticosService Viaticos => _viaticosService;


        public const string APP_VERSION = "1.6.3.81";
        public const string APP_VIGENCIA = "2026";
        public const string APP_NAME = "ESAP - Conversor PDF v" + APP_VERSION;
        private static string _defaultUserCache;
        public static string DEFAULT_USER => _defaultUserCache ?? (_defaultUserCache = new App().GetUserName());

        /// <summary>
        /// Obtiene el nombre del usuario actual del sistema Windows.
        /// </summary>
        /// <returns>Nombre de usuario en mayúsculas.</returns>
        public string GetUserName()
        {
            try
            {
                string fullUser = _userService.GetUsuarioWindows();
                if (string.IsNullOrEmpty(fullUser)) return Environment.UserName?.ToUpper() ?? "UNKNOWN_USER";
                var parts = fullUser.Split('\\');
                return parts.Length > 1 ? parts[1].ToUpper() : fullUser.ToUpper();
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error obteniendo nombre de usuario: {ex.Message}. Usando fallback.");
                return Environment.UserName?.ToUpper() ?? "UNKNOWN_USER";
            }
       }

        /// <summary>
        /// Capitaliza la primera letra de cada palabra en una cadena.
        /// </summary>
        /// <param name="texto">Texto original.</param>
        /// <returns>Texto con formato Title Case.</returns>
        public static string CapitalizarPalabras(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return texto;

            TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(texto.ToLower());
        }


        public static class FolderNames
        {
            public const string CDPS = "CDPs";
            public const string RPS = "RPs";
            public const string BOLETINES = "Boletines";
            public const string OBLIGACIONES = "Obligaciones";
            public const string VERIFICADOS = "Verificados";
            public const string UNIFICADOS = "Unificados";
            public const string SOPORTES = "S";
            public const string AUTORIZADOS = "Autorizados";
            public const string SOLICITUDES = "Solicitudes";
            /// <summary>
            /// Carpeta histórica donde se archivan los boletines de Orden de Pago unificados.
            /// </summary>
            public const string OrdenesDePago = "OrdenesDePago";
        }

        /// <summary>
        /// Combina múltiples archivos PDF en uno solo.
        /// </summary>
        /// <param name="entradas">Array de rutas de archivos PDF de entrada.</param>
        /// <param name="salida">Ruta del archivo PDF resultante.</param>
        public void CombinarArchivos(object entradas, string salida)
        {
            try
            {
                string[] listaArchivos = null;

                if (entradas == null)
                    throw new COMException("El argumento 'entradas' es nulo.");

                if (entradas is string[] arrStr)
                {
                    listaArchivos = arrStr;
                }
                else if (entradas is object[] arrObj)
                {
                    listaArchivos = new string[arrObj.Length];
                    for (int i = 0; i < arrObj.Length; i++)
                        listaArchivos[i] = arrObj[i]?.ToString();
                }
                else
                {
                    // Intento de conversiÃ³n genÃ©rica si viene como Variant encapsulado (SAFEARRAY)
                    try
                    {
                        listaArchivos = (string[])entradas;
                    }
                    catch
                    {
                        throw new COMException("El argumento 'entradas' no se pudo convertir a Array de Strings. Tipo recibido: " + entradas.GetType().FullName);
                    }
                }

                if (listaArchivos == null || listaArchivos.Length == 0)
                    throw new COMException("No hay archivos de entrada.");

                _pdfService.CombinarPDFs(listaArchivos, salida);
            }
            catch (COMException)
            {
                throw;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error crÃ­tico al combinar PDFs", ex);
                throw new COMException("Error al combinar PDFs: " + ex.Message);
            }
        }

        /// <summary>
        /// Configura los parámetros de página de una hoja de Excel para su correcta impresión a PDF.
        /// </summary>
        /// <param name="ws">Objeto Worksheet de Excel.</param>
        /// <param name="esVertical">Si es verdadero, fuerza la orientación vertical.</param>
        public static void ConfigurarPagina(dynamic ws, bool esVertical = false)
        {
            try
            {
                dynamic excelApp = ws.Application;
                dynamic ps = ws.PageSetup;
                double margin = excelApp.InchesToPoints(0.2);

                ps.LeftMargin = margin;
                ps.RightMargin = margin;
                ps.TopMargin = margin;
                ps.BottomMargin = margin;
                ps.HeaderMargin = margin;
                ps.FooterMargin = margin;

                ps.PaperSize = 1;      // xlPaperLetter

                // OrientaciÃ³n dinÃ¡mica o forzada
                if (esVertical)
                {
                    ps.Orientation = 1; // xlPortrait
                }
                else
                {
                    int ultimaFila = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1;
                    ps.Orientation = (ultimaFila > 62) ? 1 : 2; // 1=xlPortrait, 2=xlLandscape
                }


                ps.CenterHorizontally = true;
                ps.CenterVertically = true;
                ps.Zoom = false;
                ps.FitToPagesTall = 1;
                ps.FitToPagesWide = 1;
            }
            catch (Exception ex)
            {

                LoggerHelper.Warning($"Error no critico en ConfigurarPagina: {ex.Message}");
            }
        }

        public string GetOneDriveLocalPath(string path)
        {
            return OneDriveHelper.GetOneDriveLocalPath(path);
        }

        public string GetCurrentOfficeUser()
        {
            return OneDriveHelper.GetCurrentOfficeUser();
        }

        public string GetFirmaHelperUsuario()
        {
            return _userService.GetFirmaBase64(GetUserName());
        }

        /// <summary>
        /// Crea un archivo temporal a partir de una cadena en formato Base64.
        /// </summary>
        /// <param name="base64String">Contenido del archivo en Base64.</param>
        /// <param name="nombreBase">Nombre base para el archivo temporal.</param>
        /// <returns>Ruta absoluta al archivo creado.</returns>
        public string CrearArchivoBase64Temporal(string base64String, string nombreBase = "Imagen")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(base64String))
                    return "";

                string cleanBase64 = FirmaPDFsHelper.LimpiarPrefijoBase64(base64String);

                byte[] bytes = Convert.FromBase64String(cleanBase64);

                string tempFolder = Path.GetTempPath();
                string cleanName = string.Join("_", nombreBase.Split(Path.GetInvalidFileNameChars()));
                string fileName = $"{cleanName}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 8)}.png";
                string fullPath = Path.Combine(tempFolder, fileName);

                File.WriteAllBytes(fullPath, bytes);

                if (File.Exists(fullPath))
                    return fullPath;

                return "";
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error en CrearArchivoBase64Temporal para {nombreBase}", ex);
                return "";
            }
        }

        public string GetFirmaUsuarioActual()
        {
            string base64Str = GetFirmaHelperUsuario();
            if (string.IsNullOrEmpty(base64Str)) return string.Empty;

            string nombreBase = "Firma_" + DEFAULT_USER;
            return CrearArchivoBase64Temporal(base64Str, nombreBase);
        }

        public void SaveFirmaHelperUsuario(string base64)
        {
            // FASE 13: Sanitizar/Normalizar el base64 eliminando prefijos HTML/UI antes de la base de datos
            string cleanBase64 = FirmaPDFsHelper.LimpiarPrefijoBase64(base64 ?? string.Empty);
            _userService.SaveFirmaBase64(GetUserName(), cleanBase64);
        }

        public string GetFirmaHelperUsuarioInteractivo()
        {
            string base64Str = GetFirmaHelperUsuario();
            if (string.IsNullOrEmpty(base64Str))
            {
                var result = System.Windows.MessageBox.Show(
                    "Su firma no está configurada o es inválida. ¿Desea seleccionar su firma ahora para continuar?",
                    "Firma no configurada",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Question,
                    System.Windows.MessageBoxResult.No,
                    System.Windows.MessageBoxOptions.DefaultDesktopOnly);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    string nuevaFirmaBase64 = FirmasGeneradorHelper.SeleccionarYConvertirFirmaBase64();
                    if (!string.IsNullOrEmpty(nuevaFirmaBase64) && !nuevaFirmaBase64.StartsWith("ERROR"))
                    {
                        try
                        {
                            SaveFirmaHelperUsuario(nuevaFirmaBase64);
                            // FASE 9: NO llamar ResetState() en producción — destruye estado CRDT activo.
                            // FASE 9+10: Retornar directamente la firma capturada, evitar round-trip a DB.
                            base64Str = nuevaFirmaBase64;
                            System.Windows.MessageBox.Show(
                                "Firma guardada correctamente. El proceso continuará.",
                                "Éxito",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Information,
                                System.Windows.MessageBoxResult.OK,
                                System.Windows.MessageBoxOptions.DefaultDesktopOnly);
                        }
                        catch (Exception ex)
                        {
                            // FASE 10: Capturar excepción de guardado para evitar flujo roto y dar feedback al usuario
                            LoggerHelper.Error("[App] Error al guardar firma interactiva.", ex);
                            System.Windows.MessageBox.Show(
                                "Error al guardar la firma. Intente nuevamente.",
                                "Error",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Error);
                            // base64Str permanece vacío — proceso aborta correctamente
                        }
                    }
                }
            }
            return base64Str;
        }

        public string GetFirmaUsuarioActualInteractivo()
        {
            string base64Str = GetFirmaHelperUsuarioInteractivo();
            if (string.IsNullOrEmpty(base64Str)) return string.Empty;

            string nombreBase = "Firma_" + DEFAULT_USER;
            return CrearArchivoBase64Temporal(base64Str, nombreBase);
        }

        /// <summary>
        /// Procesa las hojas del libro de Excel para firmar y guardar como PDF en el flujo de Obligaciones del Contador.
        /// </summary>
        /// <param name="activeWorkbook">Libro de Excel activo.</param>
        /// <returns>Cantidad de hojas procesadas.</returns>
        public int ConvertirObligacionContador(object activeWorkbook)
        {
            return _contadorService.ConvertirObligacionContador(activeWorkbook, GetFirmaHelperUsuarioInteractivo());
        }

        public int UnificarObligacionContador(object activeWorkbook, string docType)
        {
            return FirmaUnificarObHelper.UnificarObligacionContador(activeWorkbook, docType, this);
        }

        public string GenerarFirmaBase64Txt()
        {
            return FirmasGeneradorHelper.GenerarStringBase64DesdeImagen();
        }

        public static System.Drawing.Image GetResourceImage(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                return System.Drawing.Image.FromStream(stream);
            }
        }

        public UserIdentification GetIdentification(string username)
        {
            return _userService.GetIdentification(username);
        }
    }
}


