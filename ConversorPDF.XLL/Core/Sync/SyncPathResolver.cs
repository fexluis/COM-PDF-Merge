using System;
using System.IO;
using Microsoft.Win32;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Construye y resuelve todas las rutas de carpetas utilizadas por el sistema
    ///   de sincronización CR-SQLite sobre OneDrive.
    ///
    /// ESTRUCTURA DE CARPETAS EN ONEDRIVE:
    ///   OneDrive - Escuela Superior de Administración Publica\
    ///   └── csync\
    ///       ├── VALLE\
    ///       │   ├── from_{DeviceId}_VALLE_{timestamp}.crsync   ← cambios salientes
    ///       │   └── .processed\
    ///       │       └── from_*.crsync                           ← archivos ya aplicados
    ///       ├── ANTIOQUIA\
    ///       │   └── ...
    ///       └── _Master\
    ///           └── (uso futuro - MasterConsolidator)
    ///
    /// REUTILIZACIÓN:
    ///   Internamente usa la lógica de OneDriveHelper (registro + filesystem scan)
    ///   para localizar la raíz de OneDrive. No duplica código.
    ///
    /// CONFIGURACIÓN:
    ///   AppFolderName es constante. Si se necesita cambiar en producción,
    ///   modificar solo aquí. En el futuro se puede leer de la tabla config.
    /// </summary>
    internal sealed class SyncPathResolver
    {
        // ====================================================================
        // CONFIGURACIÓN
        // ====================================================================

        /// <summary>
        /// Nombre de la carpeta raíz del sistema de sync dentro de OneDrive.
        /// Convenido con el equipo: "csync".
        /// </summary>
        private const string AppFolderName = "csync";

        /// <summary>
        /// Nombre de la subcarpeta donde se mueven los archivos ya procesados.
        /// Oculta por convención (nombre con punto) para no confundir usuarios.
        /// </summary>
        private const string ProcessedFolderName = ".processed";

        /// <summary>
        /// Nombre de la carpeta del MasterConsolidator (uso futuro).
        /// Solo el PC central (Master) tiene acceso de escritura aquí.
        /// </summary>
        private const string MasterFolderName = "_Master";

        // ====================================================================
        // ESTADO INTERNO
        // ====================================================================

        /// <summary>
        /// Raíz de OneDrive cacheada. Se resuelve una sola vez por instancia.
        /// null = aún no resuelta.
        /// </summary>
        private string _oneDriveRoot = null;

        /// <summary>
        /// Lock para resolución thread-safe de la raíz de OneDrive.
        /// </summary>
        private readonly object _lock = new object();

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Retorna la raíz de OneDrive de la ESAP para el usuario actual.
        ///
        /// ESTRATEGIA DE BÚSQUEDA (en orden de prioridad):
        ///   1. Registro de Windows (HKCU\Software\Microsoft\OneDrive\Accounts\Business*)
        ///      — más confiable en PCs corporativos con Azure AD.
        ///   2. Carpetas en el perfil del usuario (C:\Users\{user}\OneDrive - *)
        ///      — fallback cuando el registro no está disponible.
        ///
        /// CACHEO: El resultado se cachea en memoria por instancia (no por proceso)
        /// para evitar llamadas repetidas al registro/filesystem.
        ///
        /// LANZA: DirectoryNotFoundException si OneDrive no está instalado/configurado.
        /// </summary>
        internal string GetOneDriveRoot()
        {
            // Fast path (sin lock si ya está cacheado)
            if (_oneDriveRoot != null)
            {
                return _oneDriveRoot;
            }

            lock (_lock)
            {
                // Double-check
                if (_oneDriveRoot != null)
                {
                    return _oneDriveRoot;
                }

                string resolved = TryFromRegistry() ?? TryFromFileSystem();

                if (string.IsNullOrEmpty(resolved) || !Directory.Exists(resolved))
                {
                    throw new DirectoryNotFoundException(
                        "[SyncPathResolver] No se encontró la carpeta de OneDrive de la ESAP. " +
                        "Verifique que OneDrive esté instalado, autenticado y sincronizado. " +
                        "La sincronización estará deshabilitada hasta que OneDrive esté disponible.");
                }

                _oneDriveRoot = resolved;
                LoggerHelper.Info($"[SyncPathResolver] Raíz OneDrive resuelta: {_oneDriveRoot}");
                return _oneDriveRoot;
            }
        }

        /// <summary>
        /// Retorna la carpeta donde esta subunidad deposita sus archivos .crsync salientes.
        /// También es la carpeta donde otras subunidades depositan archivos para nosotros
        /// (en el modelo actual todos comparten la carpeta pública de su subunidad).
        ///
        /// Ej: C:\Users\luis\OneDrive - ESAP\csync\VALLE\
        /// </summary>
        /// <param name="subunidad">Código de subunidad en mayúsculas (ej: "VALLE").</param>
        internal string GetSubunidadInboxFolder(string subunidad)
        {
            if (string.IsNullOrWhiteSpace(subunidad))
            {
                throw new ArgumentNullException(nameof(subunidad));
            }

            return Path.Combine(GetOneDriveRoot(), AppFolderName, subunidad.ToUpperInvariant());
        }

        /// <summary>
        /// Retorna la carpeta .processed dentro de la carpeta de sync de una subunidad.
        /// Los archivos .crsync ya aplicados se mueven aquí para evitar re-procesamiento.
        ///
        /// Ej: C:\Users\luis\OneDrive - ESAP\csync\VALLE\.processed\
        /// </summary>
        /// <param name="subunidad">Código de subunidad en mayúsculas.</param>
        internal string GetProcessedFolder(string subunidad)
        {
            return Path.Combine(GetSubunidadInboxFolder(subunidad), ProcessedFolderName);
        }

        /// <summary>
        /// Retorna la carpeta raíz del sistema csync (sin subunidad).
        /// Ej: C:\Users\luis\OneDrive - ESAP\csync\
        /// </summary>
        internal string GetSyncRootFolder()
        {
            return Path.Combine(GetOneDriveRoot(), AppFolderName);
        }

        /// <summary>
        /// Retorna la carpeta _GT (Global Tables).
        /// Ej: C:\Users\luis\OneDrive - ESAP\csync\_GT\
        /// </summary>
        internal string GetGlobalFolder()
        {
            return Path.Combine(GetOneDriveRoot(), AppFolderName, "_GT");
        }

        internal string GetGlobalInboxFolder()
        {
            return GetGlobalFolder();
        }

        internal string GetGlobalProcessedFolder()
        {
            return Path.Combine(GetGlobalFolder(), ProcessedFolderName);
        }

        /// <summary>
        /// Genera el nombre de un archivo .crsync saliente.
        ///
        /// FORMATO: from_{DeviceId}_{Subunidad}_{YYYYMMDDHHmmss}.crsync
        ///
        /// EJEMPLO: from_PC-DIRECCION-VALLE_VALLE_20260613221500.crsync
        ///
        /// PROPIEDADES DEL NOMBRE:
        ///   - Ordenable lexicográficamente por timestamp.
        ///   - Identifica origen sin abrir el archivo.
        ///   - Sin espacios ni caracteres especiales (compatible con filesystems FAT32/NTFS).
        /// </summary>
        /// <param name="deviceId">Nombre del PC (Environment.MachineName, saneado).</param>
        /// <param name="subunidad">Código de la subunidad.</param>
        /// <param name="timestamp">Momento de generación (UTC).</param>
        internal static string BuildFileName(string deviceId, string subunidad, DateTime timestamp)
        {
            // Sanear deviceId: quitar caracteres que podrían romper el nombre de archivo
            string safeDeviceId = SanitizeForFileName(deviceId);
            string safeSubunidad = SanitizeForFileName(subunidad.ToUpperInvariant());
            string ts = timestamp.ToUniversalTime().ToString("yyyyMMddHHmmss");

            return $"from_{safeDeviceId}_{safeSubunidad}_{ts}.crsync";
        }

        // ====================================================================
        // MÉTODOS PRIVADOS
        // ====================================================================

        /// <summary>
        /// Intenta resolver la raíz de OneDrive desde el registro de Windows.
        /// Prioriza cuentas Business (corporativas) sobre Personal.
        ///
        /// CLAVE REGISTRY:
        ///   HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\{Name}\UserFolder
        ///   Donde {Name} es "Business1", "Business2", etc. para cuentas corporativas.
        /// </summary>
        private static string TryFromRegistry()
        {
            try
            {
                using (RegistryKey accountsKey = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\OneDrive\Accounts"))
                {
                    if (accountsKey == null)
                    {
                        return null;
                    }

                    string[] subKeys = accountsKey.GetSubKeyNames();

                    // Priorizar cuentas Business (corporativas) sobre Personal
                    Array.Sort(subKeys, (a, b) =>
                    {
                        bool aBiz = a.StartsWith("Business", StringComparison.OrdinalIgnoreCase);
                        bool bBiz = b.StartsWith("Business", StringComparison.OrdinalIgnoreCase);
                        if (aBiz != bBiz) { return aBiz ? -1 : 1; }
                        if (string.Equals(a, "Personal", StringComparison.OrdinalIgnoreCase)) { return 1; }
                        if (string.Equals(b, "Personal", StringComparison.OrdinalIgnoreCase)) { return -1; }
                        return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
                    });

                    foreach (string name in subKeys)
                    {
                        using (RegistryKey acct = accountsKey.OpenSubKey(name))
                        {
                            if (acct == null) { continue; }

                            string userFolder = acct.GetValue("UserFolder", null) as string;
                            if (!string.IsNullOrWhiteSpace(userFolder) && Directory.Exists(userFolder))
                            {
                                LoggerHelper.Debug($"[SyncPathResolver] OneDrive resuelto desde registro ({name}): {userFolder}");
                                return userFolder;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncPathResolver] Error leyendo registro OneDrive: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Fallback: busca carpetas "OneDrive - *" en el perfil del usuario.
        /// Prioriza carpetas con " - " (cuentas corporativas) sobre OneDrive personal.
        /// </summary>
        private static string TryFromFileSystem()
        {
            try
            {
                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (string.IsNullOrWhiteSpace(userProfile) || !Directory.Exists(userProfile))
                {
                    return null;
                }

                string[] candidates = Directory.GetDirectories(userProfile, "OneDrive*");

                // Primera pasada: carpeta corporativa (contiene " - ")
                foreach (string dir in candidates)
                {
                    string name = Path.GetFileName(dir);
                    if (name != null && name.Contains(" - "))
                    {
                        LoggerHelper.Debug($"[SyncPathResolver] OneDrive resuelto desde filesystem (corporativo): {dir}");
                        return dir;
                    }
                }

                // Segunda pasada: cualquier carpeta OneDrive (personal)
                if (candidates.Length > 0)
                {
                    LoggerHelper.Debug($"[SyncPathResolver] OneDrive resuelto desde filesystem (personal): {candidates[0]}");
                    return candidates[0];
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncPathResolver] Error buscando carpetas OneDrive en filesystem: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Reemplaza caracteres no válidos para nombres de archivo con guiones bajos.
        /// Asegura que el nombre de archivo generado sea compatible con NTFS, FAT32 y SMB.
        /// </summary>
        private static string SanitizeForFileName(string input)
        {
            if (string.IsNullOrEmpty(input)) { return "UNKNOWN"; }

            char[] invalid = Path.GetInvalidFileNameChars();
            char[] chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(invalid, chars[i]) >= 0 || chars[i] == ' ')
                {
                    chars[i] = '_';
                }
            }

            return new string(chars);
        }
    }
}
