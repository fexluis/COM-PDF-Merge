using System;
using ConversorPDF.XLL.Core.Data;
using ConversorPDF.XLL.Core.Helpers;
using ExcelDna.Integration;

namespace ConversorPDF.XLL
{
    /// <summary>
    /// Punto de entrada del Add-in. Se ejecuta al cargar/descargar desde Excel.
    /// </summary>
    public class ExcelAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // 1. Inicialización global del motor SQLite (Nivel 1).
            //    Asegura que interop y dependencias estén listas una sola vez por AppDomain.
            SQLiteEngineBootstrap.EnsureInitialized();

            // 2. Registrar manejador global de excepciones no controladas
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex =>
            {
                var exception = ex as Exception;
                LoggerHelper.Error("Error no controlado capturado por ExcelIntegration", exception);
                
                // Limpieza de emergencia (Nivel 1 se encarga de crasheos anteriores, pero aquí forzamos)
                SQLiteEngineBootstrap.EnsureInitialized(); 
                
                return $"[ConversorPDF] Error no controlado: {exception?.Message ?? ex?.ToString()}";
            });

            // 3. Registrar limpieza segura al finalizar el proceso
            AppDomain.CurrentDomain.ProcessExit += (s, e) =>
            {
                EncryptedSQLiteContext.CleanupOrphanedTempFiles();
            };
        }

        public void AutoClose()
        {
            EncryptedSQLiteContext.CleanupOrphanedTempFiles();
        }
    }
}