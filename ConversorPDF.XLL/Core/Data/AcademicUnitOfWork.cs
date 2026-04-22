using System;
using System.Data.SQLite;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Repositories;

namespace ConversorPDF.XLL.Core.Data
{
    /// <summary>
    /// Patrón Unit of Work para consolidar operaciones de base de datos en una sola conexión 
    /// y transacción de SQLite, evitando la creación de múltiples archivos temporales por guardado.
    /// </summary>
    public class AcademicUnitOfWork : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private SQLiteTransaction _transaction;
        private bool _disposed;

        public IAcademicRepository Repository { get; }
        
        public AcademicUnitOfWork(SQLiteConnection connection)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            _connection = connection;
            // Iniciamos la transacción de inmediato para asegurar atomicidad desde el inicio
            _transaction = _connection.BeginTransaction();
            
            // Inyectamos la conexión y la transacción en el repositorio
            Repository = new AcademicRepository(_connection, _transaction);
            
            sw.Stop();
            LoggerHelper.Info($"[PERF] [AcademicUnitOfWork] Conexión y Transacción iniciadas en {sw.ElapsedMilliseconds}ms.");
        }

        public void Commit()
        {
            try
            {
                if (_transaction != null)
                {
                    _transaction.Commit();
                    LoggerHelper.Info("[AcademicUnitOfWork] Transacción SQL (Commit) exitosa.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicUnitOfWork] Error en Commit", ex);
                Rollback();
                throw;
            }
            finally
            {
                _transaction?.Dispose();
                _transaction = null;
            }
        }

        public void Rollback()
        {
            try
            {
                if (_transaction != null)
                {
                    _transaction.Rollback();
                    LoggerHelper.Warning("[AcademicUnitOfWork] Transacción revirtida (Rollback).");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicUnitOfWork] Error crítico al intentar Rollback", ex);
            }
            finally
            {
                _transaction?.Dispose();
                _transaction = null;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _transaction?.Dispose();
                _disposed = true;
                LoggerHelper.Info("[AcademicUnitOfWork] Unit of Work cerrado.");
            }
        }
    }
}
