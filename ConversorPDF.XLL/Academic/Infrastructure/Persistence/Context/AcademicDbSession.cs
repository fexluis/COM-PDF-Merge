using System;
using System.Data;
using System.Data.SQLite;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using System.Data.Common;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context
{
    /// <summary>
    /// Patrón Shared Context para el módulo académico.
    /// Maneja el ciclo de vida de un EncryptedSQLiteContext y su conexión abierta.
    /// </summary>
    public sealed class AcademicDbSession : IDisposable
    {
        private readonly EncryptedSQLiteContext _ctx;
        private DbConnection _conn;
        private readonly SemaphoreSlim _syncLock = new SemaphoreSlim(1, 1);
        private bool _disposed;

        public AcademicDbSession()
        {
            LoggerHelper.Info("[AcademicDbSession] Iniciando sesión de base de datos compartida...");
            // Garantiza que CrSqlite + DB + SyncContext estén listos antes de abrir la conexión
            DatabaseBootstrapper.EnsureInitialized();
            _ctx = SqliteHelper.GetContextForPath();
            _conn = _ctx.Open();
        }

        /// <summary>
        /// Constructor interno para pruebas unitarias. Permite inyectar un contexto controlado.
        /// </summary>
        internal AcademicDbSession(EncryptedSQLiteContext context)
        {
            _ctx = context;
            _conn = context.Open();
        }

        private NonDisposableDbConnection _wrapper;

        public DbConnection Connection
        {
            get
            {
                if (_disposed) throw new ObjectDisposedException(nameof(AcademicDbSession));

                try
                {
                    if (_conn != null && _conn.State == ConnectionState.Open)
                    {
                        if (_wrapper == null) _wrapper = new NonDisposableDbConnection(_conn);
                        return _wrapper;
                    }
                }
                catch (ObjectDisposedException)
                {
                    LoggerHelper.Warning("[AcademicDbSession] Conexión interna dispuesta. Intentando auto-recuperación...");
                }

                // Si estaba cerrada o dispuesta, pedimos al contexto que la reabra
                _conn = _ctx.Open();
                _wrapper = new NonDisposableDbConnection(_conn);
                return _wrapper;
            }
        }

        public void MarkAsReadOnly() => _ctx?.MarkAsReadOnly();
        public void MarkChangesPersisted() => _ctx?.MarkChangesPersisted();

        /// <summary>
        /// Persiste y re-encripta el vault activo en disco.
        /// </summary>
        public Result Persist()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Este método debe llamarse DESPUÉS de que la transacción SQL haya hecho Commit.
        /// </summary>
        public void NotifyChangesCommitted()
        {
            ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
        }

        /// <summary>
        /// Permite ejecutar operaciones asíncronas de manera segura (thread-safe),
        /// ya que SQLiteConnection no soporta concurrencia por defecto.
        /// </summary>
        public async Task<T> ExecuteAsync<T>(Func<IDbConnection, Task<T>> operation)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(AcademicDbSession));

            await _syncLock.WaitAsync();
            try
            {
                return await operation(Connection);
            }
            finally
            {
                _syncLock.Release();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            _syncLock.Wait();
            try
            {
                LoggerHelper.Info("[AcademicDbSession] Cerrando sesión y liberando recursos...");
                try
                {
                    var scheduler = ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler;
                    if (scheduler != null)
                    {
                        // Fase 16 FIX: pasar el tiempo restante a WaitForPendingSync para que
                        // el timeout total sea exactamente 10s y no se extienda por iteraciones.
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        var totalTimeout = TimeSpan.FromSeconds(10);
                        do
                        {
                            scheduler.Flush();
                            var remaining = totalTimeout - sw.Elapsed;
                            if (remaining <= TimeSpan.Zero) break;
                            scheduler.WaitForPendingSync(remaining);
                        }
                        while (sw.Elapsed < totalTimeout && scheduler.IsDirty);
                    }

                    if (_conn != null && _conn.State != ConnectionState.Closed)
                    {
                        _conn.Close();
                    }
                }
                catch { }

                // FIX: _ctx viene de GetContextForPath() (instancia nueva, NO singleton).
                // Disponerlo es seguro: el contador _activeContextsPerTempDb protege el temp file compartido.
                _ctx?.Dispose();
                
                _disposed = true;
            }
            finally
            {
                _syncLock.Release();
                _syncLock?.Dispose();
            }
        }
    }
}
