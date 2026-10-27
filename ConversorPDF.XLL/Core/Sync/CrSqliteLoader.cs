using System;
using System.Data.SQLite;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Gestiona el ciclo de vida de la extensión nativa CR-SQLite (crsqlite.dll).
    ///   CR-SQLite es una extensión SQLite que convierte tablas en CRDTs (Conflict-free
    ///   Replicated Data Types), permitiendo sincronización offline sin conflictos (LWW).
    ///
    /// PATRÓN:
    ///   Singleton lazy con double-check locking. La DLL se extrae del recurso embebido
    ///   una sola vez por proceso y se carga con LoadLibrary (igual que NativeLibraryPreloader
    ///   hace con SQLite.Interop.dll).
    ///
    /// USO:
    ///   1. Llamar Initialize() al arrancar el add-in (una sola vez en ExcelAddIn.AutoOpen).
    ///   2. Antes de usar crsql_* en cualquier conexión, llamar EnableOnConnection(conn).
    ///
    /// IMPORTANTE:
    ///   - La extensión debe habilitarse POR CONEXIÓN. No es global al proceso.
    ///   - Requiere PRAGMA enable_load_extension = 1 antes de load_extension().
    ///   - El archivo temp NO se borra mientras el proceso vive (la DLL puede estar en uso).
    /// </summary>
    internal static class CrSqliteLoader
    {
        // ====================================================================
        // CONSTANTES
        // ====================================================================

        /// <summary>
        /// Nombre del recurso embebido en el ensamblado.
        /// Debe coincidir con la ruta registrada en el .csproj como EmbeddedResource.
        /// </summary>
        private const string ResourceName =
            "ConversorPDF.XLL.Core.Sync.Native.crsqlite.dll";

        /// <summary>
        /// Prefijo del archivo temporal extraído.
        /// Se usa un nombre fijo (no GUID) para que si el proceso falla y reinicia,
        /// no se acumulen múltiples copias del DLL en %TEMP%.
        /// </summary>
        private const string TempDllFileName = "crsqlite_esap.dll";

        /// <summary>
        /// Punto de entrada de la extensión. Estándar de CR-SQLite.
        /// Se pasa a load_extension() como segundo argumento para ser explícitos.
        /// </summary>
        private const string ExtensionEntryPoint = "sqlite3_crsqlite_init";

        // ====================================================================
        // ESTADO INTERNO (thread-safe)
        // ====================================================================

        /// <summary>
        /// Ruta absoluta al DLL extraído en %TEMP%.
        /// null = no inicializado aún. Se setea una sola vez en Initialize().
        /// </summary>
        private static string _extractedDllPath = null;

        /// <summary>
        /// Flag: true si la DLL fue extraída Y cargada en el proceso exitosamente.
        /// </summary>
        private static bool _initialized = false;

        /// <summary>
        /// Lock para garantizar que Initialize() se ejecute una sola vez incluso
        /// si múltiples hilos lo llaman concurrentemente al arrancar.
        /// </summary>
        private static readonly object _initLock = new object();

        // ====================================================================
        // P/INVOKE: LoadLibrary
        // ====================================================================

        /// <summary>
        /// Win32 LoadLibrary: fuerza la carga de la DLL nativa en el proceso.
        /// Esto "ancla" la DLL en memoria para que load_extension() no la busque
        /// en rutas del sistema (puede fallar en entornos COM restringidos).
        /// </summary>
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Flag Interlocked para asegurar que solo un hilo intente la inicialización
        /// concurrente y prevenir stampedes.
        /// </summary>
        private static int _initGate = 0;

        /// <summary>
        /// Extrae crsqlite.dll del recurso embebido a %TEMP% y la pre-carga en el proceso.
        ///
        /// CUÁNDO LLAMAR: Una sola vez en ExcelAddIn.AutoOpen() antes de cualquier
        /// operación de sincronización.
        ///
        /// IDEMPOTENTE: Puede llamarse múltiples veces sin efecto secundario.
        /// La segunda llamada retorna inmediatamente.
        /// </summary>
        /// <summary>
        /// Extrae crsqlite.dll del recurso embebido a %TEMP% y la pre-carga en el proceso.
        ///
        /// CUÁNDO LLAMAR: Una sola vez en ExcelAddIn.AutoOpen() antes de cualquier
        /// operación de sincronización. También se llama defensivamente desde
        /// SqliteHelper.EnsureSchemaUpToDate() para entornos de test donde AutoOpen no corre.
        ///
        /// IDEMPOTENTE: Puede llamarse múltiples veces sin efecto secundario.
        /// La segunda llamada retorna inmediatamente vía fast-path _initialized.
        ///
        /// THREAD-SAFE: Patrón Interlocked gate + lock double-check.
        ///   - _initGate garantiza que solo UN hilo entre al bloque de trabajo.
        ///   - lock (_initLock) + double-check protege el estado _initialized.
        ///   - Si la extracción falla, _initGate se resetea a 0 para permitir reintento.
        /// </summary>
        internal static void Initialize()
        {
            // Fast path: DLL ya cargada. Sin lock, sin costo.
            if (_initialized) return;

            // Atomic gate: de N hilos concurrentes, solo UNO continúa (el que pasa 0→1).
            // Los demás retornan silenciosamente — el hilo ganador se encargará.
            if (Interlocked.CompareExchange(ref _initGate, 1, 0) != 0) return;

            try
            {
                lock (_initLock)
                {
                    // Double-check bajo lock: el hilo ganador del gate verifica de nuevo
                    // por si hubo una carrera muy estrecha antes del CompareExchange.
                    if (_initialized) return;

                    LoggerHelper.Info("[CrSqliteLoader] Iniciando extracción de crsqlite.dll...");

                    try
                    {
                        // Paso 1: Extraer DLL del recurso embebido a %TEMP%
                        _extractedDllPath = ExtractEmbeddedDll();

                        // Paso 2: Pre-cargar en el proceso con LoadLibrary.
                        // Previene problemas de PATH en entornos COM/Excel restringidos.
                        PreloadWithLoadLibrary(_extractedDllPath);

                        _initialized = true;
                        LoggerHelper.Info($"[CrSqliteLoader] Inicializado exitosamente. DLL en: {_extractedDllPath}");
                    }
                    catch (Exception ex)
                    {
                        // NO relanzamos: un fallo aquí deshabilita el sync silenciosamente.
                        // El add-in sigue funcionando sin sync (degradación graceful).
                        LoggerHelper.Error("[CrSqliteLoader] Error al inicializar. El sync estará deshabilitado.", ex);
                        _initialized = false;
                    }
                }
            }
            finally
            {
                // Si la inicialización falló (ej. I/O temporal), resetear el gate
                // para que una llamada futura (tras resolver el problema) pueda reintentar.
                if (!_initialized)
                {
                    Interlocked.Exchange(ref _initGate, 0);
                }
            }
        }

        /// <summary>
        /// Habilita la extensión CR-SQLite en una conexión SQLite específica.
        ///
        /// DEBE LLAMARSE:
        ///   - Antes de cualquier uso de crsql_changes, crsql_as_crr, crsql_apply_changes.
        ///   - Por cada conexión (la extensión NO se propaga entre conexiones SQLite).
        ///
        /// IMPORTANTE:
        ///   La conexión debe estar ABIERTA antes de llamar este método.
        ///   Requiere que Initialize() haya sido llamado previamente.
        ///
        /// RETORNA: true si la extensión fue habilitada, false si hubo un error.
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta en la cual habilitar CR-SQLite.</param>
        internal static bool EnableOnConnection(SQLiteConnection connection)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            if (!_initialized || string.IsNullOrEmpty(_extractedDllPath))
            {
                // Fase 2: Reducido a Debug. Es esperado que durante la creación inicial (Seed) 
                // o tests, CrSqliteLoader aún no esté inicializado.
                LoggerHelper.Debug("[CrSqliteLoader] EnableOnConnection llamado sin inicialización previa (esperado en creación DB/Tests).");
                return false;
            }

            try
            {
                if (connection.State != System.Data.ConnectionState.Open)
                {
                    connection.Open();
                }

                // Paso 1: Habilitar la capacidad de cargar extensiones en esta conexión.
                // Por seguridad, SQLite deshabilita load_extension por defecto.
                // Usamos el parámetro nativo de System.Data.SQLite.
                connection.EnableExtensions(true);

                // Paso 2: Cargar la extensión CR-SQLite.
                // Usamos la ruta absoluta extraída en %TEMP% para evitar problemas de PATH.
                // El segundo parámetro es el punto de entrada de la extensión.
                connection.LoadExtension(_extractedDllPath, ExtensionEntryPoint);

                LoggerHelper.Debug("[CrSqliteLoader] Extensión CR-SQLite habilitada en conexión.");
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[CrSqliteLoader] Error al habilitar extensión CR-SQLite en la conexión.", ex);
                return false;
            }
        }

        /// <summary>
        /// Indica si CR-SQLite fue inicializado correctamente en este proceso.
        /// Usar para verificar antes de intentar operaciones de sync.
        /// </summary>
        internal static bool IsAvailable
        {
            get { return _initialized; }
        }

        // ====================================================================
        // MÉTODOS PRIVADOS
        // ====================================================================

        /// <summary>
        /// Extrae la DLL embebida en el ensamblado al directorio %TEMP% del usuario.
        ///
        /// ESTRATEGIA DE NOMBRE FIJO:
        ///   Se usa un nombre fijo ("crsqlite_esap.dll") en vez de un GUID para evitar
        ///   que múltiples arranques del add-in dejen archivos huérfanos en %TEMP%.
        ///   Si el archivo ya existe y tiene el mismo tamaño, se reutiliza.
        ///
        /// RETORNA: Ruta absoluta al archivo extraído.
        /// LANZA: InvalidOperationException si el recurso no se encuentra en el ensamblado.
        /// </summary>
        private static string ExtractEmbeddedDll()
        {
            string tempPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), TempDllFileName));

            // Leer el recurso embebido del ensamblado
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream resourceStream = assembly.GetManifestResourceStream(ResourceName))
            {
                if (resourceStream == null)
                {
                    throw new InvalidOperationException(
                        $"[CrSqliteLoader] No se encontró el recurso embebido '{ResourceName}'. " +
                        $"Verificar que crsqlite.dll está registrado como EmbeddedResource en el .csproj.");
                }

                long resourceLength = resourceStream.Length;

                // Optimización: si el archivo ya existe con el mismo tamaño, no re-extraer.
                // Asume que si el tamaño coincide, el contenido es el mismo (versión de DLL igual).
                if (File.Exists(tempPath))
                {
                    var existingFile = new FileInfo(tempPath);
                    if (existingFile.Length == resourceLength)
                    {
                        LoggerHelper.Debug($"[CrSqliteLoader] DLL ya extraída en {tempPath}. Reutilizando.");
                        return tempPath;
                    }

                    // Tamaño diferente: nueva versión del add-in. Reemplazar.
                    LoggerHelper.Info($"[CrSqliteLoader] Tamaño diferente ({existingFile.Length} vs {resourceLength}). Re-extrayendo DLL.");
                }

                // Escribir bytes del recurso al archivo temporal
                byte[] dllBytes = new byte[resourceLength];
                resourceStream.Read(dllBytes, 0, dllBytes.Length);
                File.WriteAllBytes(tempPath, dllBytes);

                LoggerHelper.Info($"[CrSqliteLoader] DLL extraída a: {tempPath} ({resourceLength:N0} bytes)");
            }

            return tempPath;
        }

        /// <summary>
        /// Usa Win32 LoadLibrary para pre-cargar la DLL en el proceso.
        /// Esto "ancla" la DLL en el espacio de memoria del proceso y previene
        /// errores de "DLL not found" en entornos COM con restricciones de PATH.
        /// </summary>
        /// <param name="dllPath">Ruta absoluta al archivo DLL.</param>
        private static void PreloadWithLoadLibrary(string dllPath)
        {
            IntPtr handle = LoadLibrary(dllPath);
            if (handle == IntPtr.Zero)
            {
                // Win32 error code para diagnóstico
                int errorCode = Marshal.GetLastWin32Error();
                LoggerHelper.Warning(
                    $"[CrSqliteLoader] LoadLibrary falló (Win32 error: {errorCode}). " +
                    $"load_extension() intentará cargar la DLL igualmente.");
                // No lanzamos excepción: load_extension puede tener éxito aún
            }
            else
            {
                LoggerHelper.Debug($"[CrSqliteLoader] LoadLibrary exitoso. Handle: 0x{handle.ToInt64():X}");
            }
        }
    }
}
