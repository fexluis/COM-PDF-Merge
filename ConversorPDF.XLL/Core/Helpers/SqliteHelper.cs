using System;
using System.IO;
using System.Reflection;
using System.Security;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Data;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Ayudante para la gestión de la base de datos SQLite con encriptación Vault (App-Level).
    /// </summary>
    public static class SqliteHelper
    {
        private static readonly object _initLock = new object();
        internal static readonly string DbFolder = GetSafeDbFolder();
        private static string _customVaultPath = null;
        public static string DbPathEnc => _customVaultPath ?? Path.Combine(DbFolder, "AcademicData.db.enc");
        public static bool IsTestVaultPathActive => _customVaultPath != null;
        private static bool _initialized = false;
        private static readonly HashSet<string> _migratedVaultsThisSession = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            try
            {
                return Path.GetFullPath(path).ToLowerInvariant();
            }
            catch
            {
                return path.ToLowerInvariant();
            }
        }

        internal static bool WasSchemaMigratedForPath(string path)
        {
            lock (_initLock)
            {
                string p = NormalizePath(path ?? DbPathEnc);
                return _migratedVaultsThisSession.Contains(p);
            }
        }

        internal static void MarkVaultAsMigrated(string path)
        {
            lock (_initLock)
            {
                string p = NormalizePath(path ?? DbPathEnc);
                _migratedVaultsThisSession.Add(p);
            }
        }

        internal static void ResetMigratedVaultsTrackingForTests()
        {
            lock (_initLock)
            {
                _migratedVaultsThisSession.Clear();
            }
        }

        internal static bool WasAnySchemaMigrated()
        {
            lock (_initLock)
            {
                return _migratedVaultsThisSession.Count > 0;
            }
        }

        public static void SetTestVaultPath(string path)
        {
            lock (_initLock)
            {
                _customVaultPath = path;
                _sharedContext?.Dispose();
                _sharedContext = null;
                _initialized = false; // Forzar re-inicialización con el nuevo path
            }
        }

        public static void ResetState()
        {
            lock (_initLock)
            {
                _sharedContext?.Dispose();
                _sharedContext = null;
                _initialized = false;
                ConversorPDF.XLL.Core.Sync.SyncInitializer.ResetStateForTests(); // Fase 54
            }
        }

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

        internal static string GetEmbeddedSqlResource(string resourceFileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string resourceName = $"ConversorPDF.XLL.Academic.Infrastructure.Persistence.Sql.{resourceFileName}";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException($"No se encontró el recurso embebido '{resourceName}'.");
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static EncryptedSQLiteContext _sharedContext;

        private static int _initInProgress = 0;
        private static volatile int _regeneratingCrdts = 0; // Fase 43

        /// <summary>
        /// Obtiene un contexto encriptado seguro (Unit of Work).
        /// Emplea un patrón Singleton para evitar la sobrecarga de extracción continua,
        /// manteniendo tolerancia a fallos si el contexto es disposed por otra clase.
        /// </summary>
        public static EncryptedSQLiteContext GetContext()
        {
            // Si hay inicialización en progreso en OTRO llamado, esperar o retornar error controlado
            if (System.Threading.Interlocked.CompareExchange(ref _initInProgress, 1, 0) != 0)
            {
                // Ya hay un GetContext() en progreso. Esperar activamente (máx 5s).
                int waited = 0;
                while (System.Threading.Volatile.Read(ref _initInProgress) == 1 && waited < 5000)
                {
                    System.Threading.Thread.Sleep(100);
                    waited += 100;
                }
                if (System.Threading.Volatile.Read(ref _initInProgress) == 1)
                    throw new InvalidOperationException("GetContext() reentrante detectado. Inicialización previa no completada.");
            }

            EncryptedSQLiteContext result = null;
            bool migrationOccurred = false;
            try
            {
                lock (_initLock)
                {
                    // Nivel 1: Asegurar inicialización global
                    SQLiteEngineBootstrap.EnsureInitialized();

                    if (!_initialized || !File.Exists(DbPathEnc))
                    {
                        InitializeDatabase(DbPathEnc);
                        _initialized = true;
                    }

                    if (_sharedContext != null)
                    {
                        // Fase 41 y 50: Esperar si hay regeneración CRDT en progreso.
                        // Sin timeout porque la regeneración de 15 tablas puede tardar varios segundos, 
                        // y lanzar excepción de timeout rompería el flujo en lugar de esperar pacientemente.
                        while (_regeneratingCrdts == 1)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        try
                        {
                            // Si no lanza excepción, el contexto sigue vivo
                            _sharedContext.Open();
                            return _sharedContext;
                        }
                        catch (ObjectDisposedException)
                        {
                            LoggerHelper.Warning("[SqliteHelper] El SharedContext fue disposed previamente. Recreando instancia...");
                            // No llamar Dispose() aquí; el contexto en vuelo podría fallar.
                            _sharedContext = null;
                        }
                    }

                    var context = new EncryptedSQLiteContext(DbPathEnc, GetSecurePassword());

                    // Pilar A: Asegurar que el esquema esté actualizado (Migraciones)
                    // Usamos el mismo contexto que acabamos de crear para no desencriptar dos veces
                    if (File.Exists(DbPathEnc))
                    {
                        EnsureSchemaUpToDate(context);
                        // FASE 23 Corrección: Registrar explícitamente el vault base como migrado 
                        // para evitar I/O innecesario en posteriores invocaciones a GetContextForPath(null)
                        MarkVaultAsMigrated(DbPathEnc);
                    }

                    // Si las migraciones o la inicialización persistieron datos, cerraron el contexto.
                    // Debemos recrearlo antes de retornarlo para que el llamador pueda usarlo.
                    if (context.IsDisposed)
                    {
                        context.Dispose();
                        context = new EncryptedSQLiteContext(DbPathEnc, GetSecurePassword());
                    }

                    _sharedContext = context;
                    migrationOccurred = WasAnySchemaMigrated();

                    // Fase 46: Activar bandera ANTES de liberar lock para cerrar ventana
                    // donde un hilo entra por el early-return justo después de soltar el lock.
                    if (migrationOccurred)
                    {
                        System.Threading.Interlocked.Exchange(ref _regeneratingCrdts, 1);
                    }
                }

                result = _sharedContext;
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _initInProgress, 0);
            }

            // Fase 39: EnableAllCrdts DESPUÉS de liberar _initInProgress (y fuera del lock)
            // para no bloquear a otros llamadores de GetContext.
            if (migrationOccurred)
            {
                try
                {
                    ConversorPDF.XLL.Core.Sync.SyncInitializer.EnableAllCrdts();
                }
                finally
                {
                    // Fase 45 y 46: Recrear _sharedContext dentro del lock para atomicidad
                    // y desactivar la bandera SOLO DESPUÉS de haber regenerado el contexto.
                    lock (_initLock)
                    {
                        // Fase 40: Recrear _sharedContext para reflejar los triggers CRDT
                        // regenerados y persistidos en el vault por SyncInitializer.
                        _sharedContext?.Dispose();
                        _sharedContext = new EncryptedSQLiteContext(DbPathEnc, GetSecurePassword());
                        result = _sharedContext;

                        System.Threading.Interlocked.Exchange(ref _regeneratingCrdts, 0);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Obtiene el contexto compartido existente sin desencadenar el pipeline pesado de inicialización ni bloquear el semáforo reentrante.
        /// Retorna null si no está inicializado o si el contexto fue descartado.
        /// </summary>
        public static EncryptedSQLiteContext TryGetSharedContextWithoutInit()
        {
            var ctx = _sharedContext;
            if (ctx == null || !_initialized) return null;
            try
            {
                ctx.Open();
                return ctx;
            }
            catch (ObjectDisposedException)
            {
                return null;
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"[SqliteHelper] TryGetSharedContextWithoutInit no pudo abrir contexto: {ex.Message}");
                return null;
            }
        }

        private static readonly Lazy<ISchemaMigrator> _schemaMigrator = new Lazy<ISchemaMigrator>(BuildMigrator);

        private static ISchemaMigrator BuildMigrator()
        {
            var migrations = new List<IMigration>
            {
                new MigrationV2_VistasAcademicas(),
                new MigrationV3_Periodos(),
                new MigrationV4_VistaSesiones(),
                new MigrationV5_SemestreSesiones(),
                new MigrationV6_CedulaTercero(),
                new MigrationV7_AsignaturaAbreviatura(),
                new MigrationV8_SemestreGrupo(),
                new MigrationV9_NewSchemaColumns(),
                new MigrationV10_FixVinculacionData(),
                new MigrationV11_AddProgramaToViews(),
                new MigrationV12_ExpandSesionesView(),
                new MigrationV13_AddFestivosTable(),
                new MigrationV14_UpdateSesionesSchema(),
                new MigrationV15_AddModalidadVinculacion(),
                new MigrationV16_AddSalvedadesConflictos(),
                new MigrationV17_AddTransporteSedeView(),
                new MigrationV18_AddBaseGastosTerrestre(),
                new MigrationV19_AddNivelToViews(),
                new MigrationV20_AddTipoFestivo(),
                new MigrationV21_AddDivipolaTable(),
                new MigrationV22_AddActivoToSedes(),
                new MigrationV23_AddResidenciaDivipola(),
                new MigrationV24_AddSubunidadSync(),
                new MigrationV25_AddSubsedeToSesionesView(),
                new MigrationV26_AddTerritoriales(),
                new MigrationV27_AddSiifToSedes(),
                new MigrationV28_AddCdpDependencias(),
                new MigrationV29_FixUsuariosPerfil(),
                new MigrationV30_AddTerminalAereo(),
                new MigrationV31_AddIdPeriodoToConfig(),
                new MigrationV32_RegenerateCrdtTriggers(),
                new MigrationV33_UpdateSesionesDetalleForTerminales(),
                new MigrationV34_AddSubunidadToCdpDependencias(), // FASE 29: Registro de migración
                new MigrationV35_SeedPermisosDefault(), // FASE 13: Tabla de Permisos dinámicos
                new MigrationV37_RecreatePermisos(),
                new MigrationV38_AddGestionUsuariosPermiso(),
                new MigrationV39_AddFirmaToUsuarios(),
                new MigrationV40_RegenerateUsuariosTriggers(),
                new MigrationV41_DeduplicateUsuarios(),
                new MigrationV42_DivipolaRepair(),
                new MigrationV43_FixBaseGastosTerrestreSubunidad(),
                new MigrationV44_AddDependenciaToCdpDependencias(),
                new MigrationV45_AddIdTerritorialToConfig(), // FASE 35: Registro de migración para soportar el cambio id_territorial
                new MigrationV46_AddDependenciaFullToCdpDependencias(), // FASE 21: Registro de migración para dependencia_full en cdpDependencias
                new MigrationV47_SeedTerminalesCanonicos(), // FASE 1: Siembra canónica de tabla terminales
                new MigrationV48_AddTercerosAsignaturas(), // FASE 3: Pool docente-asignatura y backfill histórico
                new MigrationV49_AddAgHorariosTables(), // FASE 19: Tablas y vistas del motor de Algoritmo Genético
                new MigrationV50_AddPeriodoVentana(), // FASE 153: Ventana temporal en periodos y backfill
                new MigrationV51_PeriodoVentanaRepair(), // FASE 182: Reparación idempotente de ventana temporal en periodos
                new MigrationV52_IsolateDemoAulas(), // FASE 233: Aislamiento idempotente de aulas demo en ag_aula
                new MigrationV53_AddAgArbolTables(), // FASE 236: Perfiles de selección de árbol ag_arbol y nodos
                new MigrationV54_AddVacanteColumns(), // FASE 317: Columnas de vacante en staging y producción
                new MigrationV55_AddAgGenPatronSemanal(), // FASE 382/395/402: Patrón multi-día y franja Lunes
                new MigrationV56_AjustarInicioPeriodo2026_02(), // FASE 386/406: Ajuste inicio período 2026-02
                new MigrationV57_RepairSalvedadesConflictos(), // FASE 459: Reparación de columnas idsesion_a/b en salvedades_conflictos
                new MigrationV58_EnsureTerceroVacante(), // FASE 465: Asegurar tercero centinela VACANTE (cédula 1111111111)
                new MigrationV59_RecreateAgViews(), // FASE 468: Recrear vistas de staging y asegurar ag_salvedades_conflictos
                new MigrationV60_AddAgStagingIndices(), // FASE 498: Índices de optimización en ag_vinculaciones y ag_sesiones
                new MigrationV61_AddFechaCreacionSalvedades(), // FASE 514: Columna fecha_creacion en ag_salvedades_conflictos
                new MigrationV62_AddAgAulaVentanasYBloquesSemilla(), // FASE 649: Columnas de decisión en ag_aula y ag_arbol_nodo, seed 42 bloques
                new MigrationV63_UnificarHorariosMilitar(), // FASE 650: Unificación de horarios a formato militar
                new MigrationV64_AddPernoctaUltimoDia(), // FASE 665: Adición de isPernoctaUltimoDia y actualización de vw_SesionesDetalle
                new MigrationV65_PeriodoVentanaCrdtRepair(), // FASE 745: Reparación CRDT de ventana temporal en periodos (resuelve hueco V50/V51)
                new MigrationV66_RenombrarModalidadTutoria(), // FASE 784: Renombrado canónico de modalidades (Plataforma y Mixta)
                new MigrationV67_RecrearVistasAgCanonicas(), // FASE 854: Vistas AG canónicas en BD existentes
                new MigrationV68_NormalizarFechasAgSesiones(), // FASE 875: Normalización de fechas legacy en ag_sesiones
                new MigrationV69_RevalidarVistasAg(), // FASE 893: Revalidación forzada de vistas canónicas AG
                new MigrationV70_AddAgCorridaVacantes(), // FASE 956: Columna vacantes en ag_corrida
                new MigrationV71_AuditarHorariosMeridiano(), // FASE 1000: Auditoría diagnóstica de horarios afectados por hh
                new MigrationV72_LimpiarBloquesDomingoYDuplicados(), // FASE 1009: Depuración de bloques de domingo y duplicados 8h
                new MigrationV73_LimpiarDiasPatronMultidia(),        // FASE 1005: Limpieza de días de patrón multidía (Fase 1212)
                new MigrationV74_SembrarBloques4hFaltantes(),         // FASE 1061: asegurar bloques 4h para pregrado
                new MigrationV75_SembrarBloquesDomingo(),             // FASE 1188: sembrar bloques canónicos de domingo
                new MigrationV76_SanearDimensionesAgAula(),           // FASE 1191: sanear dimensiones de decisión en ag_aula (I1)
                new MigrationV77_AddPernoctaAgVinculaciones(),        // FASE 1279: columna isPernoctaUltimoDia en vinculaciones y ag_vinculaciones
                new MigrationV78_AddAgFiltroSeleccion(),              // FASE 1315 / 1390: ag_filtro_seleccion
                new MigrationV79_NormalizeUsuariosIds(),              // Normalización determinista de usuarios y deduplicación
                new MigrationV80_AddOrigenPatronAgAulaAsignatura(),    // FASE 1735: columna origen_patron en ag_aula_asignatura
                new MigrationV81_AddAgGenFechasGen(),                 // FASE 1815: columna fechas_gen en ag_gen
                new MigrationV82_AddAgCandidatoPlanAprobado(),       // T4: candidato y ratificación aditivos
                new MigrationV83_AddAgPeriodoVentanaAuditoria(),     // FASE 1854: auditoría de extensión de ventana en periodos
                new MigrationV84_AddAgCorridaVacantesConsentidas()   // FASE 1885: columna vacantes_consentidas en ag_corrida
            };
            return new SchemaMigrator(migrations, new VaultPersister(), new LoggerAdapter());
        }

        private static void EnsureSchemaUpToDate(EncryptedSQLiteContext context)
        {
            // FIX FASE 13: Garantizar que la extensión CR-SQLite esté disponible ANTES
            // de ejecutar las migraciones (especialmente V24 que requiere crsql_as_crr).
            // Esto es vital para entornos de test donde ExcelAddIn.AutoOpen() no se ejecuta.
            ConversorPDF.XLL.Core.Sync.CrSqliteLoader.Initialize();

            _schemaMigrator.Value.EnsureSchemaUpToDate(context);
        }

        private static readonly System.Text.RegularExpressions.Regex IdentificadorSql =
            new System.Text.RegularExpressions.Regex(@"^[A-Za-z_][A-Za-z0-9_]{0,63}$", System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>
        /// Comprueba de forma segura si una tabla existe en la base de datos SQLite (Fase 1890/1891/1892/1894).
        /// Retorna tri-estado: Existe, NoExiste o Indeterminada.
        /// Nota de arquitectura: Migraciones previas (V54-V81) conservan consultas inline a sqlite_master;
        /// se unifica este helper para migraciones y repositorios modernos sin reescritura retroactiva.
        /// </summary>
        public static ExistenciaSql TablaExiste(IDbConnection conn, string tabla, IDbTransaction transaction = null)
        {
            if (conn == null || conn.State != ConnectionState.Open || string.IsNullOrWhiteSpace(tabla) || !IdentificadorSql.IsMatch(tabla))
            {
                return ExistenciaSql.Indeterminada;
            }

            // Fase 1894: prefijo reservado 'sqlite_' no corresponde a tablas de negocio de la aplicación
            if (tabla.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase))
            {
                return ExistenciaSql.Indeterminada;
            }

            try
            {
                int count = conn.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@tabla;",
                    new { tabla }, transaction);

                return count > 0 ? ExistenciaSql.Existe : ExistenciaSql.NoExiste;
            }
            catch (Exception ex)
            {
                LoggerHelper.WarningLimitado($"TablaExiste|{tabla}",
                    $"[SQL] No se pudo verificar la tabla '{tabla}': {ex.Message}");
                return ExistenciaSql.Indeterminada;
            }
        }

        /// <summary>
        /// Comprueba de forma segura y tolerante si una columna existe en una tabla SQLite (Fase 1885/1888/1889/1891/1894).
        /// Retorna tri-estado: Existe, NoExiste o Indeterminada.
        /// </summary>
        public static ExistenciaSql ColumnaExiste(IDbConnection conn, string tabla, string columna, IDbTransaction transaction = null)
        {
            if (conn == null || conn.State != ConnectionState.Open || string.IsNullOrWhiteSpace(tabla) || string.IsNullOrWhiteSpace(columna))
            {
                return ExistenciaSql.Indeterminada;
            }

            if (!IdentificadorSql.IsMatch(tabla) || tabla.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase))
            {
                return ExistenciaSql.Indeterminada;
            }

            try
            {
                var info = conn.Query($"PRAGMA table_info({tabla});", transaction: transaction);
                foreach (var row in info)
                {
                    var dict = (IDictionary<string, object>)row;
                    if (dict.TryGetValue("name", out var val) && string.Equals(val?.ToString(), columna, StringComparison.OrdinalIgnoreCase))
                    {
                        return ExistenciaSql.Existe;
                    }
                }
                return ExistenciaSql.NoExiste;
            }
            catch (Exception ex)
            {
                LoggerHelper.WarningLimitado($"ColumnaExiste|{tabla}.{columna}",
                    $"[SQL] No se pudo verificar '{tabla}.{columna}': {ex.Message}");
                return ExistenciaSql.Indeterminada;
            }
        }

        internal static string ExtraerSubunidadDeRuta(string vaultPath)
        {
            if (string.IsNullOrWhiteSpace(vaultPath)) return null;
            string nombre = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(vaultPath));
            const string prefijo = "AcademicData_";
            if (!nombre.StartsWith(prefijo, StringComparison.OrdinalIgnoreCase)) return null;
            string sub = nombre.Substring(prefijo.Length).Trim();
            return sub.Length == 0 ? null : sub.ToUpperInvariant();
        }

        internal static void ApplyBaseSchemaAndSeeds(IDbConnection connection, IDbTransaction transaction, string subunidadDestino = null)
        {
            LoggerHelper.Info("[SqliteHelper] Ejecutando siembra de esquema base y datos semilla (AcademicSchema, vinculaciones, usuarios)...");

            // FASE 76 FIX: Garantizar que si las tablas legadas (v2) existían, contengan las columnas necesarias antes de ejecutar scripts de esquema y datos
            int tercerosExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='terceros';", transaction: transaction);
            if (tercerosExists > 0)
            {
                var terCols = connection.Query<string>("SELECT name FROM pragma_table_info('terceros');", transaction: transaction);
                bool hasRol = System.Linq.Enumerable.Any(terCols, x => string.Equals(x, "rol", StringComparison.OrdinalIgnoreCase));
                bool hasCat = System.Linq.Enumerable.Any(terCols, x => string.Equals(x, "categoria", StringComparison.OrdinalIgnoreCase));
                if (hasRol && !hasCat)
                    try { connection.Execute("ALTER TABLE terceros RENAME COLUMN rol TO categoria;", transaction: transaction); } catch { }

                var tercerosCols = new[]
                {
                    new { Name = "cedula", Type = "TEXT" },
                    new { Name = "nombre", Type = "TEXT" },
                    new { Name = "departamento", Type = "TEXT" },
                    new { Name = "tipo", Type = "TEXT" },
                    new { Name = "salario", Type = "REAL" },
                    new { Name = "categoria", Type = "TEXT" },
                    new { Name = "entidad_bancaria", Type = "TEXT" },
                    new { Name = "cuenta_bancaria", Type = "TEXT" },
                    new { Name = "tipo_cuenta_bancaria", Type = "TEXT" },
                    new { Name = "telefono", Type = "TEXT" },
                    new { Name = "correo", Type = "TEXT" },
                    new { Name = "correo_personal", Type = "TEXT" },
                    new { Name = "fecha_nacimiento", Type = "TEXT" },
                    new { Name = "contrato", Type = "TEXT" },
                    new { Name = "residencia", Type = "TEXT" },
                    new { Name = "id_residencia_divipola", Type = "INTEGER NULL" },
                    new { Name = "pensionado", Type = "INTEGER DEFAULT 0" },
                    new { Name = "activo", Type = "INTEGER DEFAULT 1" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" },
                    new { Name = "puntaje_ocasionales", Type = "REAL DEFAULT 0" }
                };
                foreach (var c in tercerosCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('terceros') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE terceros ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            int programasExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='programas';", transaction: transaction);
            if (programasExists > 0)
            {
                var progCols = connection.Query<string>("SELECT name FROM pragma_table_info('programas');", transaction: transaction);
                bool hasTipo = System.Linq.Enumerable.Any(progCols, x => string.Equals(x, "tipo", StringComparison.OrdinalIgnoreCase));
                bool hasMod = System.Linq.Enumerable.Any(progCols, x => string.Equals(x, "modalidad", StringComparison.OrdinalIgnoreCase));
                if (hasTipo && !hasMod)
                    try { connection.Execute("ALTER TABLE programas RENAME COLUMN tipo TO modalidad;", transaction: transaction); } catch { }

                var programasCols = new[]
                {
                    new { Name = "pensum", Type = "INTEGER" },
                    new { Name = "actual", Type = "INTEGER DEFAULT 0" },
                    new { Name = "activo", Type = "INTEGER DEFAULT 0" },
                    new { Name = "modalidad", Type = "TEXT" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" }
                };
                foreach (var c in programasCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('programas') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE programas ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            int asignaturasExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='asignaturas';", transaction: transaction);
            if (asignaturasExists > 0)
            {
                var asigCols = connection.Query<string>("SELECT name FROM pragma_table_info('asignaturas');", transaction: transaction);
                bool hasCod = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "codigo", StringComparison.OrdinalIgnoreCase));
                bool hasCodMat = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "cod_materia", StringComparison.OrdinalIgnoreCase));
                bool hasNomAsig = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "nombreasignatura", StringComparison.OrdinalIgnoreCase));
                bool hasNom = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "nombre", StringComparison.OrdinalIgnoreCase));
                bool hasTotHoras = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "total_horas", StringComparison.OrdinalIgnoreCase));
                bool hasIntHoraria = System.Linq.Enumerable.Any(asigCols, x => string.Equals(x, "int_horaria", StringComparison.OrdinalIgnoreCase));

                if (hasCod && !hasCodMat)
                    try { connection.Execute("ALTER TABLE asignaturas RENAME COLUMN codigo TO cod_materia;", transaction: transaction); } catch { }
                if (hasNomAsig && !hasNom)
                    try { connection.Execute("ALTER TABLE asignaturas RENAME COLUMN nombreasignatura TO nombre;", transaction: transaction); } catch { }
                if (hasTotHoras && !hasIntHoraria)
                    try { connection.Execute("ALTER TABLE asignaturas RENAME COLUMN total_horas TO int_horaria;", transaction: transaction); } catch { }

                var asignaturasCols = new[]
                {
                    new { Name = "semestre", Type = "TEXT" },
                    new { Name = "num_nivel", Type = "INTEGER" },
                    new { Name = "cod_unidad", Type = "TEXT" },
                    new { Name = "cod_materia", Type = "TEXT" },
                    new { Name = "uni_teorica", Type = "INTEGER" },
                    new { Name = "metodologia", Type = "TEXT" },
                    new { Name = "abreviatura", Type = "TEXT" },
                    new { Name = "nucleo", Type = "TEXT" },
                    new { Name = "nombre", Type = "TEXT" },
                    new { Name = "creditos", Type = "INTEGER" },
                    new { Name = "int_horaria", Type = "INTEGER" },
                    new { Name = "tipoasignatura", Type = "TEXT" },
                    new { Name = "tipocredito", Type = "TEXT" },
                    new { Name = "prerequisito", Type = "TEXT" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" }
                };
                foreach (var c in asignaturasCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('asignaturas') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE asignaturas ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            int sedesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='sedes';", transaction: transaction);
            if (sedesExists > 0)
            {
                var sedesCols = new[]
                {
                    new { Name = "iddivipola", Type = "INTEGER" },
                    new { Name = "departamento", Type = "TEXT" },
                    new { Name = "ciudad", Type = "TEXT" },
                    new { Name = "departamentoSiif", Type = "TEXT" },
                    new { Name = "ciudadSiif", Type = "TEXT" },
                    new { Name = "subsede", Type = "TEXT" },
                    new { Name = "categoria", Type = "INTEGER DEFAULT 2" },
                    new { Name = "transporte", Type = "REAL" },
                    new { Name = "activo", Type = "INTEGER DEFAULT 1" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" }
                };
                foreach (var c in sedesCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('sedes') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE sedes ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            int sesionesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='sesiones';", transaction: transaction);
            if (sesionesExists > 0)
            {
                var sesionesCols = new[]
                {
                    new { Name = "fechasesion", Type = "TEXT" },
                    new { Name = "horariosesion", Type = "TEXT" },
                    new { Name = "horas", Type = "REAL" },
                    new { Name = "viatica", Type = "INTEGER" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" }
                };
                foreach (var c in sesionesCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('sesiones') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE sesiones ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            int vinculacionesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction);
            if (vinculacionesExists > 0)
            {
                var vinculacionesCols = new[]
                {
                    new { Name = "idperiodo", Type = "INTEGER" },
                    new { Name = "grupo", Type = "TEXT" },
                    new { Name = "semestre", Type = "TEXT" },
                    new { Name = "total_horas", Type = "INTEGER" },
                    new { Name = "dia_horas_semanales", Type = "TEXT" },
                    new { Name = "fecha_inicio", Type = "TEXT" },
                    new { Name = "fecha_fin", Type = "TEXT" },
                    new { Name = "modalidadVinculacion", Type = "TEXT" },
                    new { Name = "subunidad", Type = "TEXT DEFAULT ''" },
                    new { Name = "isTerminalAereo", Type = "INTEGER DEFAULT 0" },
                    new { Name = "idTerminales", Type = "INTEGER DEFAULT 0" },
                    new { Name = "gastoTerminalAerea", Type = "REAL DEFAULT 0" }
                };
                foreach (var c in vinculacionesCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('vinculaciones') WHERE name=@Name;", new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE vinculaciones ADD COLUMN {c.Name} {c.Type};", transaction: transaction); } catch { }
                    }
                }
            }

            // FASE 201 FIX: Garantizar que la tabla periodos contenga periodo_inicio y periodo_fin antes de aplicar AcademicSchema o habilitar CRR
            int periodosExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';", transaction: transaction);
            if (periodosExists > 0)
            {
                var periodosCols = new[]
                {
                    new { Name = "periodo_inicio", Type = "TEXT" },
                    new { Name = "periodo_fin",    Type = "TEXT" },
                    new { Name = "subunidad",      Type = "TEXT DEFAULT ''" }
                };
                foreach (var c in periodosCols)
                {
                    int hasCol = connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(*) FROM pragma_table_info('periodos') WHERE name=@Name;",
                        new { Name = c.Name }, transaction: transaction);
                    if (hasCol == 0)
                    {
                        try { connection.Execute($"ALTER TABLE periodos ADD COLUMN {c.Name} {c.Type};", transaction: transaction); }
                        catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0) { }
                    }
                }

                // Backfill de fechas (idempotente)
                var periodosDatos = new[]
                {
                    new { Id = 1, Ini = "2026-01-12", Fin = "2026-06-19" },
                    new { Id = 2, Ini = "2026-08-15", Fin = "2026-12-12" },
                    new { Id = 3, Ini = "2027-01-01", Fin = "2027-06-30" }
                };
                foreach (var d in periodosDatos)
                {
                    connection.Execute(
                        @"UPDATE periodos SET periodo_inicio = @Ini, periodo_fin = @Fin
                           WHERE idperiodo = @Id AND (periodo_inicio IS NULL OR periodo_inicio = '');",
                        d, transaction: transaction);
                }
            }

            // Fase 221: Aislar aulas demo del seed de V49 marcándolas como DEMO para no contaminar catálogos reales
            int agAulaExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_aula';", transaction: transaction);
            if (agAulaExists > 0)
            {
                connection.Execute(
                    @"UPDATE ag_aula 
                      SET subunidad = 'DEMO' 
                      WHERE codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01') 
                        AND (subunidad IS NULL OR subunidad = '' OR subunidad = 'VALLE');",
                    transaction: transaction);
            }

            connection.Execute(GetEmbeddedSqlResource("AcademicSchema.sql"), transaction: transaction);

            // Fase 191: Omitir dump histórico si vinculaciones ya contiene datos
            int nVinc = connection.QueryFirstOrDefault<int>("SELECT COUNT(1) FROM vinculaciones;", transaction: transaction);
            if (nVinc == 0)
            {
                connection.Execute(GetEmbeddedSqlResource("Academic_migracion_vinculaciones_sesiones.sql"), transaction: transaction);
            }
            else
            {
                LoggerHelper.Info("[SqliteHelper] Dump omitido: `vinculaciones` ya contiene datos.");
            }

            connection.Execute(GetEmbeddedSqlResource("usuarios.sql"), transaction: transaction);
            connection.Execute(@"
                UPDATE vinculaciones SET fecha_inicio = NULL WHERE fecha_inicio = '--';
                UPDATE vinculaciones SET fecha_fin    = NULL WHERE fecha_fin    = '--';", transaction: transaction);

            // FASE 75 FIX: Actualizar inmediatamente config.subunidad si se conoce la subunidad de destino del vault
            if (!string.IsNullOrWhiteSpace(subunidadDestino))
            {
                connection.Execute("UPDATE config SET subunidad = @Sub WHERE id = 1;",
                                   new { Sub = subunidadDestino.ToUpperInvariant() }, transaction: transaction);
            }

            // FASE 73 y 75 FIX: Usar la subunidad configurada en el propio vault (config.subunidad) y limitar
            // el fallback 'VALLE' únicamente a los usuarios semilla canónicos de usuarios.sql.
            connection.Execute(@"
                UPDATE usuarios SET subunidad = (SELECT subunidad FROM config LIMIT 1)
                 WHERE (subunidad IS NULL OR trim(subunidad) = '')
                   AND EXISTS (SELECT 1 FROM config WHERE subunidad IS NOT NULL AND trim(subunidad) <> '');

                UPDATE usuarios SET subunidad = 'VALLE'
                 WHERE (subunidad IS NULL OR trim(subunidad) = '')
                   AND nombreusuario IN ('Usuario','luis.bechara','alexander.montano','carmroja',
                                         'eduin.palacios','isadora.chavez','elizabeth.cordobao');
            ", transaction: transaction);

            int sinSubunidad = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM usuarios WHERE subunidad IS NULL OR trim(subunidad) = '';", transaction: transaction);
            if (sinSubunidad > 0)
            {
                LoggerHelper.Warning($"[SqliteHelper] {sinSubunidad} usuario(s) sin subunidad tras la siembra; se resolverá/n al registrar en SyncContext.");
            }

            // FASE 70 FIX: El DDL canónico de divipola (8 columnas) vive en Academic_divipola.sql; no duplicarlo con esquemas estrechos.
            int divipolaTableExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='divipola';", transaction: transaction);

            int divipolaCanonicas = (divipolaTableExists == 0) ? 0 : connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM pragma_table_info('divipola') " +
                "WHERE name IN ('departamento_codigo','municipio','municipio_codigo','municipio_normalizado','municipio_categoria');",
                transaction: transaction);

            int divipolaRows = (divipolaTableExists == 0) ? 0 : connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM divipola;", transaction: transaction);

            if (divipolaTableExists == 0 || divipolaCanonicas < 5 || divipolaRows == 0)
            {
                if (divipolaTableExists > 0 && divipolaCanonicas < 5)
                {
                    LoggerHelper.Warning("[SqliteHelper] divipola con esquema no canónico. Se recrea desde el recurso.");
                    connection.Execute("DROP TABLE IF EXISTS divipola;", transaction: transaction);
                }
                LoggerHelper.Info("[SqliteHelper] Cargando Academic_divipola.sql (DDL canónico + catálogo)...");
                connection.Execute(GetEmbeddedSqlResource("Academic_divipola.sql"), transaction: transaction);
            }

            // FASE 71: Aseguramiento de vistas del motor AG
            EnsureAgViews(connection, transaction);
        }

        private static void ValidarVista(IDbConnection connection, IDbTransaction transaction, string vista)
        {
            try
            {
                connection.Execute($"SELECT * FROM {vista} LIMIT 0;", transaction: transaction);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[SqliteHelper] Vista inválida {vista}: {ex.Message}");
                throw;
            }
        }

        public static void EnsureAgViews(IDbConnection connection, IDbTransaction transaction = null)
        {
            int agSesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';", transaction: transaction);
            if (agSesExists == 0) return;

            string sessionsViewSql = @"
DROP VIEW IF EXISTS vw_AgSesionesDetalle;
CREATE VIEW vw_AgSesionesDetalle AS
SELECT 
    s.id AS IdSesion,
    s.idvinculacion AS IdVinculacion,
    v.id_corrida AS IdCorrida,
    v.idperiodo AS IdPeriodo,
    v.idsede AS IdSede,
    v.idprograma AS IdPrograma,
    s.idtercero AS IdTercero,
    p.modalidad AS ProgramaTipo,
    p.abreviatura AS ProgramaAbreviatura,
    CASE 
        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
        ELSE v.semestre || '-' || v.grupo 
    END AS Semestre,
    t.nombre AS DocenteNombre,
    t.cedula AS CedulaTercero,
    a.nombre AS AsignaturaNombre,
    a.abreviatura AS AsignaturaAbreviatura,
    s2.ciudad AS SedeCiudad,
    s2.subsede AS Subsede,
    CASE strftime('%w', NULLIF(s.fechasesion, ''))
        WHEN '0' THEN 'DOMINGO'
        WHEN '1' THEN 'LUNES'
        WHEN '2' THEN 'MARTES'
        WHEN '3' THEN 'MIÉRCOLES'
        WHEN '4' THEN 'JUEVES'
        WHEN '5' THEN 'VIERNES'
        WHEN '6' THEN 'SÁBADO'
    END AS DiaSemana,
    s.horas AS Horas,
    s.id_aula AS IdAula,
    s.idgen AS IdGen,
    s.es_vacante AS EsVacante,
    s.motivo_vacante AS MotivoVacante,
    s.candidatos_pool AS CandidatosPool,
    v.dia_horas_semanales AS DiaHorasSemanales,
    s.horariosesion AS HorarioSesion,
    NULLIF(s.fechasesion, '') AS FechaSesion,
    s.viatica AS Viatica,
    s2.departamento AS SedeDepartamento,
    s2.transporte AS TransporteSede,
    p.nombreprograma AS ProgramaNombre,
    a.creditos AS AsignaturaCreditos,
    a.metodologia AS AsignaturaMetodologia,
    a.cod_materia AS AsignaturaCodMateria,
    a.num_nivel AS Nivel,
    v.grupo AS Grupo,
    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin,
    v.modalidadVinculacion AS Modalidad,
    v.isTerminalAereo AS IsTerminalAereo,
    v.idTerminales AS IdTerminales,
    v.gastoTerminalAerea AS GastoTerminalAerea,
    term.terminal AS TerminalAereaNombre,
    v.subunidad AS Subunidad
FROM ag_sesiones s
JOIN ag_vinculaciones v ON s.idvinculacion = v.id AND IFNULL(v.estado, 0) <> 3
LEFT JOIN programas p ON v.idprograma = p.id
LEFT JOIN asignaturas a ON v.idasignatura = a.id
LEFT JOIN sedes s2 ON v.idsede = s2.id
LEFT JOIN terceros t ON s.idtercero = t.id
LEFT JOIN terminales term ON (CASE WHEN v.idTerminales IS NULL OR v.idTerminales <= 0 THEN 1 ELSE v.idTerminales END) = term.id;";

            connection.Execute(sessionsViewSql, transaction: transaction);
            ValidarVista(connection, transaction, "vw_AgSesionesDetalle");

            bool hayPeriodos = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';",
                transaction: transaction) > 0;

            if (!hayPeriodos)
            {
                LoggerHelper.Warning("[SqliteHelper] 'periodos' ausente: se omite vw_AgVinculacionesDetalle.");
                return;
            }

            string vinculacionesViewSql = @"
DROP VIEW IF EXISTS vw_AgVinculacionesDetalle;
CREATE VIEW vw_AgVinculacionesDetalle AS
SELECT v.id AS Id,
       v.id AS IdVinculacion,
       v.id_corrida AS IdCorrida,
       v.id_individuo AS IdIndividuo,
       v.idsede AS IdSede,
       v.idprograma AS IdPrograma,
       v.idasignatura AS IdAsignatura,
       v.idtercero AS IdTercero,
       v.grupo AS Grupo,
       IFNULL(v.estado, 0) AS Estado,
       v.dia_horas_semanales AS DiaHorasSemanales,
       v.total_horas AS TotalHoras,
       v.es_vacante AS EsVacante,
       v.motivo_vacante AS MotivoVacante,
       v.candidatos_pool AS CandidatosPool,
       pe.idperiodo AS IdPeriodo,
       pe.periodo AS PeriodoNombre,
       p.modalidad AS ProgramaTipo,
       p.abreviatura AS ProgramaAbreviatura,
       t.nombre AS TerceroNombre,
       CASE WHEN v.grupo IS NULL OR v.grupo='' OR v.grupo='--' THEN v.semestre
            ELSE v.semestre || '-' || v.grupo END AS Semestre,
       a.nombre AS AsignaturaNombre,
       s.ciudad AS SedeCiudad,
       NULLIF(NULLIF(v.fecha_inicio,''),'--') AS SesionFechaInicial,
       NULLIF(NULLIF(v.fecha_fin,''),'--')    AS SesionFechaFinal,
       NULLIF(NULLIF(v.fecha_inicio,''),'--') AS FechaInicio,
       NULLIF(NULLIF(v.fecha_fin,''),'--')    AS FechaFin,
       v.total_horas AS SesionHoras,
       (SELECT CASE WHEN EXISTS (SELECT 1 FROM ag_sesiones ses
                                 WHERE ses.idvinculacion=v.id AND ses.viatica=1)
                    THEN 1 ELSE 0 END) AS SesionViatica,
       v.modalidadVinculacion AS Modalidad,
       v.modalidadVinculacion AS ModalidadVinculacion,
       v.subunidad AS Subunidad
FROM ag_vinculaciones v
LEFT JOIN programas  p  ON v.idprograma  = p.id
LEFT JOIN asignaturas a ON v.idasignatura = a.id
LEFT JOIN sedes      s  ON v.idsede      = s.id
LEFT JOIN terceros   t  ON v.idtercero   = t.id
LEFT JOIN periodos   pe ON v.idperiodo   = pe.idperiodo;";

            connection.Execute(vinculacionesViewSql, transaction: transaction);
            ValidarVista(connection, transaction, "vw_AgVinculacionesDetalle");
        }

        private static void InitializeDatabase(string customPath = null)
        {
            string path = string.IsNullOrWhiteSpace(customPath) ? DbPathEnc : customPath;
            // Configuración global de Dapper: permite mapear snake_case (DB) → PascalCase (C#)
            // Ejemplo: cod_materia → CodMateria, int_horaria → IntHoraria
            // Sin esto, Dapper deja esas propiedades en null/0 sin lanzar excepción.
            Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true;

            try
            {
                string folder = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
                {
                    LoggerHelper.Info($"[SqliteHelper] Creando directorio de base de datos en: {folder}");
                    Directory.CreateDirectory(folder);
                }

                if (!File.Exists(path))
                {
                    LoggerHelper.Info($"[SqliteHelper] No se encontró base de datos encriptada en '{path}'. Iniciando creación y semilla (seed)...");
                    using (var context = new EncryptedSQLiteContext(path, GetSecurePassword()))
                    {
                        var connection = context.Open();
                        using (var tx = connection.BeginTransaction())
                        {
                            try
                            {
                                ApplyBaseSchemaAndSeeds(connection, tx, ExtraerSubunidadDeRuta(path));
                                tx.Commit();
                            }
                            catch
                            {
                                try { tx.Rollback(); } catch { }
                                throw;
                            }
                        }

                        // Persistencia explícita del esquema inicial para que no se pierda al cerrar
                        // Fase 17: Snapshot temporal de TempDbPath antes de la liberación del contexto interno
                        string tempDbPathForSeed = context.TempDbPath;
                        string encryptedFilePathForSeed = context.EncryptedFilePath; // Fase 20
                        context.CloseConnectionResiliently();
                        byte[] initialBytes = null;
                        if (!string.IsNullOrEmpty(tempDbPathForSeed) && File.Exists(tempDbPathForSeed))
                        {
                            initialBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(tempDbPathForSeed);
                        }

                        // Fase 18: Aseguramiento estricto de la inicialización
                        if (initialBytes != null && initialBytes.Length > 0)
                        {
                            UpdateVaultDirectly(initialBytes, encryptedFilePathForSeed); // Fase 20
                            context.MarkChangesPersisted();
                            LoggerHelper.Info("[SqliteHelper] Base de datos inicializada y persistida exitosamente.");
                        }
                        else
                        {
                            LoggerHelper.Error("[SqliteHelper] CRÍTICO: No se pudo leer el archivo temporal para persistir la DB inicial.");
                            throw new InvalidOperationException("Fallo en la persistencia inicial de la base de datos.");
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
        public static void UpdateVaultDirectly(byte[] plainData, string customPath = null, IProgress<string> progress = null)
        {
            EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainData, customPath ?? DbPathEnc, GetSecurePassword(), progress);
        }

        /// <summary>Publicación T4 condicionada a identidad capturada; no altera a los escritores externos.</summary>
        internal static bool UpdateVaultIfUnchanged(byte[] bytes, string path,
            EncryptedSQLiteContext.VaultIdentity expected, System.Threading.CancellationToken ct)
        {
            using (var password = GetSecurePassword())
                return EncryptedSQLiteContext.ReEncryptAndPersistIfUnchanged(bytes, path, password, expected, ct);
        }

        /// <summary>
        /// Retorna la ruta del vault para una subunidad específica.
        /// Ej: AcademicData_VALLE.db.enc, AcademicData_BOLIVAR.db.enc
        /// </summary>
        public static string GetDbPathForSubunidad(string subunidad)
        {
            string safeName = subunidad?.ToUpperInvariant().Trim() ?? "DEFAULT";
            return Path.Combine(DbFolder, $"AcademicData_{safeName}.db.enc");
        }

        /// <summary>
        /// Crea un <see cref="EncryptedSQLiteContext"/> para un vault en una ruta específica.
        ///
        /// USO EN SYNC:
        ///   El SubunidadSyncService se ejecuta en hilos background y necesita su propio
        ///   contexto independiente del que usa la UI (para no bloquear Excel).
        ///   Este método permite abrir el vault con la misma contraseña pero en una instancia
        ///   separada que tiene su propio temp .db file.
        ///
        /// IMPORTANTE:
        ///   El llamador es responsable de hacer Dispose() del contexto retornado.
        ///   Usar siempre en un bloque using { }.
        /// </summary>
        /// <param name="vaultPath">Ruta al archivo .db.enc. Si null, usa DbPathEnc.</param>
        public static EncryptedSQLiteContext GetContextForPath(string vaultPath = null)
        {
            SQLiteEngineBootstrap.EnsureInitialized();
            string path = string.IsNullOrWhiteSpace(vaultPath) ? DbPathEnc : vaultPath;

            lock (_initLock)
            {
                if (!File.Exists(path))
                {
                    InitializeDatabase(path);
                }
            }

            // Siempre retornamos un contexto nuevo para que los bloques using() en las
            // rutinas de sincronización no dispongan el _sharedContext.
            var context = new EncryptedSQLiteContext(path, GetSecurePassword());

            // Fase 36: Asegurar migraciones en vaults secundarios solo si no se migraron en esta sesión
            if (File.Exists(path) && !WasSchemaMigratedForPath(path))
            {
                EnsureSchemaUpToDate(context);

                // Fase 35: Recrear contexto si EnsureSchemaUpToDate lo dejó disposed
                if (context.IsDisposed)
                {
                    context.Dispose();
                    context = new EncryptedSQLiteContext(path, GetSecurePassword());
                }

                // Fase 42: Si se aplicaron migraciones a este vault secundario, regenerar triggers CRDT.
                if (WasSchemaMigratedForPath(path))
                {
                    try
                    {
                        ConversorPDF.XLL.Core.Sync.SyncInitializer.EnableAllCrdts(path);
                    }
                    finally
                    {
                        // Fase 44 y 51: Recrear contexto siempre - EnableAllCrdts usó contexto propio
                        // y persistió al vault. El temp DB de este contexto exterior está obsoleto.
                        // El finally garantiza que se recrea incluso si la regeneración CRDT falla,
                        // evitando dejar un contexto inconsistente.
                        context.Dispose();
                        context = new EncryptedSQLiteContext(path, GetSecurePassword());
                    }
                }
            }

            // NO llamar a MarkAsReadOnly() aquí — los contextos de sync SÍ modifican la DB (EnableAllCrdts).
            // Cada llamador es responsable de llamar a MarkAsReadOnly() o MarkChangesPersisted() según su uso.
            return context;
        }

    }
}
