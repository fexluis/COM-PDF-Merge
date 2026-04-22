using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Proporciona funcionalidades de registro de eventos (logging) para la aplicación.
    /// Utiliza una cola asíncrona para escribir logs en archivos de texto locales.
    /// </summary>
    public static class LoggerHelper
    {
        private static readonly ConcurrentQueue<string> _queue = new ConcurrentQueue<string>();
        private static readonly string _logDirectory;

        static LoggerHelper()
        {
            try
            {
                string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                _logDirectory = Path.Combine(documents, "Plantillas personalizadas de Office", "ConversorPDF_Logs");
                if (!Directory.Exists(_logDirectory)) Directory.CreateDirectory(_logDirectory);

                // Iniciar tarea de escritura en segundo plano
                Task.Run(() => ProcessQueue());

                // Limpiar logs antiguos (más de 7 días)
                CleanupOldLogs(7);
            }
            catch
            {
                _logDirectory = Path.GetTempPath();
            }
        }

        /// <summary> Registra un mensaje de depuración. </summary>
        public static void Debug(string msg, [CallerMemberName] string caller = "") => Enqueue("DEBUG", caller, msg);
        /// <summary> Registra un mensaje informativo. </summary>
        public static void Info(string msg, [CallerMemberName] string caller = "") => Enqueue("INFO", caller, msg);
        /// <summary> Registra un mensaje de éxito. </summary>
        public static void Success(string msg, [CallerMemberName] string caller = "") => Enqueue("SUCCESS", caller, msg);
        /// <summary> Registra una advertencia. </summary>
        public static void Warning(string msg, [CallerMemberName] string caller = "") => Enqueue("WARN", caller, msg);
        /// <summary>
        /// Registra un error detallado, incluyendo la excepción si se proporciona.
        /// </summary>
        /// <param name="msg">Resumen del error.</param>
        /// <param name="ex">Excepción capturada.</param>
        /// <param name="caller">Nombre del método que invoca el log (automático).</param>
        public static void Error(string msg, Exception ex = null, [CallerMemberName] string caller = "")
        {
            StringBuilder sb = new StringBuilder(msg);
            if (ex != null)
            {
                sb.AppendLine($"\nExcepción: {ex.Message}");
                sb.AppendLine($"StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null) sb.AppendLine($"Inner: {ex.InnerException.Message}");
            }
            Enqueue("ERROR", caller, sb.ToString());
        }

        private static void Enqueue(string level, string caller, string msg)
        {
            string logLine = $"[{DateTime.Now:HH:mm:ss}] [{level.PadRight(7)}] [{caller}] {msg}";
            _queue.Enqueue(logLine);
            System.Diagnostics.Debug.WriteLine(logLine);
        }

        private static void ProcessQueue()
        {
            while (true)
            {
                if (_queue.TryDequeue(out string line))
                {
                    try
                    {
                        string fileName = $"log_{DateTime.Now:yyyyMMdd}.txt";
                        string path = Path.Combine(_logDirectory, fileName);
                        File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
                    }
                    catch { /* Fallback silencioso */ }
                }
                else
                {
                    System.Threading.Thread.Sleep(500); // Esperar si la cola está vacía
                }
            }
        }

        private static void CleanupOldLogs(int days)
        {
            Task.Run(() =>
            {
                try
                {
                    var files = Directory.GetFiles(_logDirectory, "log_*.txt");
                    var cutoff = DateTime.Now.AddDays(-days);
                    foreach (var f in files)
                    {
                        if (File.GetCreationTime(f) < cutoff) File.Delete(f);
                    }
                }
                catch { }
            });
        }

        public static string GetLogPath() => _logDirectory;
    }
}