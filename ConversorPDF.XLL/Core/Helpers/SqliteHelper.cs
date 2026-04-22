using System;
using System.IO;
using System.Reflection;
using System.Security;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Core.Data;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Ayudante para la gestión de la base de datos SQLite con encriptación Vault (App-Level).
    /// </summary>
    public static class SqliteHelper
    {
        internal static readonly string DbFolder = GetSafeDbFolder();
        internal static readonly string DbPathEnc = Path.Combine(DbFolder, "AcademicData.db.enc");
        private static bool _initialized = false;

        /// <summary>
        /// Retorna la ruta física del archivo Vault persistente.
        /// </summary>
        public static string GetVaultPath() => DbPathEnc;

        private static string GetSafeDbFolder()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrEmpty(appData))
            {
                appData = Environment.GetEnvironmentVariable("LOCALAPPDATA")
                       ?? Environment.GetEnvironmentVariable("USERPROFILE")
                       ?? Path.GetTempPath();
            }
            return Path.Combine(appData, "ConversorPDF", "Data");
        }

        private static SecureString GetSecurePassword()
        {
            var secureString = new SecureString();
            foreach (char c in ConstantsHelper.DB_PASSWORD)
            {
                secureString.AppendChar(c);
            }
            secureString.MakeReadOnly();
            return secureString;
        }

        private static string GetEmbeddedSchema()
        {
            var assembly = Assembly.GetExecutingAssembly();
            string resourceName = "ConversorPDF.XLL.Core.Data.AcademicSchema.sql";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException($"No se encontró el recurso embebido '{resourceName}'.");

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Obtiene un contexto encriptado seguro (Unit of Work).
        /// Cada llamada retorna una nueva instancia para aislamiento transaccional.
        /// </summary>
        public static EncryptedSQLiteContext GetContext()
        {
            // Nivel 1: Asegurar inicialización global
            SQLiteEngineBootstrap.EnsureInitialized();

            if (!_initialized || !File.Exists(DbPathEnc))
            {
                InitializeDatabase();
                _initialized = true;
            }

            return new EncryptedSQLiteContext(DbPathEnc, GetSecurePassword());
        }

        private static void InitializeDatabase()
        {
            try
            {
                if (!Directory.Exists(DbFolder))
                {
                    LoggerHelper.Info($"[SqliteHelper] Creando directorio de base de datos en: {DbFolder}");
                    Directory.CreateDirectory(DbFolder);
                }

                if (!File.Exists(DbPathEnc))
                {
                    LoggerHelper.Info("[SqliteHelper] No se encontró base de datos encriptada. Iniciando creación y semilla (seed)...");
                    using (var context = new EncryptedSQLiteContext(DbPathEnc, GetSecurePassword()))
                    {
                        var connection = context.Open();
                        string schema = GetEmbeddedSchema();
                        
                        LoggerHelper.Info($"[SqliteHelper] Ejecutando schema embebido ({schema.Length} bytes)...");
                        connection.Execute(schema);
                        
                        // Persistencia explícita del esquema inicial para que no se pierda al cerrar
                        context.CloseConnectionResiliently();
                        byte[] initialBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
                        if (initialBytes != null)
                        {
                            UpdateVaultDirectly(initialBytes);
                            LoggerHelper.Info("[SqliteHelper] Base de datos inicializada y persistida exitosamente.");
                        }
                    }
                }
                else
                {
                    LoggerHelper.Info("[SqliteHelper] Base de datos académica detectada y lista.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SqliteHelper] Error CRÍTICO durante InitializeDatabase", ex);
                throw;
            }
        }

        /// <summary>
        /// Método utilitario para actualizar el Vault directamente desde bytes estáticos,
        /// sin necesidad de inicializar un DbContext ligero dummy para la tarea.
        /// </summary>
        public static void UpdateVaultDirectly(byte[] plainData, IProgress<string> progress = null)
        {
            EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainData, DbPathEnc, GetSecurePassword(), progress);
        }
    }
}
