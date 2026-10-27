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
        private static string _logDirectory;
        private static readonly ConcurrentDictionary<string, int> _emitidos = new ConcurrentDictionary<string, int>();

        /// <summary> Nivel mínimo emitido a archivo (Fase 594). </summary>
        public static NivelLog Minimo { get; set; } = NivelLog.Info;
        /// <summary> Habilita detalle técnico (SQL/parámetros/datos por paso) (Fase 594). </summary>
        public static bool ModoVerboso { get; set; }

        /// <summary>
        /// Directorio destino de los archivos log_yyyyMMdd.txt.
        /// Permite redirigir la salida en entornos de pruebas unitarias/integración (Fase 1066)
        /// evitando la contaminación de logs de producción. Producción conserva la ruta por defecto.
        /// </summary>
        internal static string DirectorioLog
        {
            get => _logDirectory;
            set
            {
                if (string.IsNullOrWhiteSpace(value)) return;
                try
                {
                    Directory.CreateDirectory(value);
                    _logDirectory = value;
                }
                catch { /* se conserva el directorio vigente si no es posible crear la ruta */ }
            }
        }

        static LoggerHelper()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrEmpty(appData))
                {
                    appData = Environment.GetEnvironmentVariable("LOCALAPPDATA")
                           ?? Environment.GetEnvironmentVariable("USERPROFILE")
                           ?? Path.GetTempPath();
                }

                _logDirectory = Path.Combine(appData, "ConversorPDF", "log");
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

        /// <summary> Registra un mensaje de depuración (Fase 594: solo en modo verboso). </summary>
        public static void Debug(string msg, [CallerMemberName] string caller = "")
        {
            if (!ModoVerboso) return;
            Enqueue("DEBUG", caller, msg);
        }

        /// <summary> Registra un mensaje informativo. </summary>
        public static void Info(string msg, [CallerMemberName] string caller = "") => Enqueue("INFO", caller, msg);
        /// <summary> Registra un mensaje de éxito. </summary>
        public static void Success(string msg, [CallerMemberName] string caller = "") => Enqueue("SUCCESS", caller, msg);
        /// <summary> Registra una advertencia. </summary>
        public static void Warning(string msg, [CallerMemberName] string caller = "") => Enqueue("WARN", caller, msg);
        /// <summary> Registra una advertencia con excepción. </summary>
        public static void Warning(string msg, Exception ex, [CallerMemberName] string caller = "") =>
            Enqueue("WARN", caller, ex == null ? msg : $"{msg} | Excepción: {ex.Message}");

        /// <summary> Advertencia limitada por clave: máximo maxVeces por proceso (Fase 594). </summary>
        public static void WarningLimitado(string clave, string msg, int maxVeces = 1, [CallerMemberName] string caller = "")
        {
            int n = _emitidos.AddOrUpdate(clave, 1, (_, v) => v + 1);
            if (n <= maxVeces) Enqueue("WARN", caller, msg);
            else if (n == maxVeces + 1) Enqueue("WARN", caller, $"[…] '{clave}' superó {maxVeces} emisión(es); resto silenciado.");
        }

        /// <summary> Limpia los contadores de dedupe al iniciar un proceso (Fase 594). </summary>
        public static void ReiniciarDedupe() => _emitidos.Clear();

        /// <summary> Cierra una etapa con una única línea de resumen legible (Fase 594). </summary>
        public static void Resumen(string titulo, params (string Clave, object Valor)[] datos) =>
            Enqueue("INFO", titulo, string.Join(" | ", System.Linq.Enumerable.Select(datos, d => $"{d.Clave}={d.Valor}")));

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
            NivelLog nivel;
            switch (level)
            {
                case "DEBUG": nivel = NivelLog.Debug; break;
                case "WARN": nivel = NivelLog.Warn; break;
                case "ERROR": nivel = NivelLog.Error; break;
                case "SUCCESS": nivel = NivelLog.Success; break;
                default: nivel = NivelLog.Info; break;
            }
            if (nivel < Minimo) return;

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