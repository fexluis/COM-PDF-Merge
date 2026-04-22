using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Utilidad para forzar la carga de librerías nativas (DLLs) antes de que el motor managed las requiera.
    /// Resuelve problemas de EntryPointNotFoundException en entornos complejos como Excel.
    /// </summary>
    public static class NativeLibraryPreloader
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        /// <summary>
        /// Asegura que SQLite.Interop.dll esté cargado en el proceso actual desde la ruta correcta.
        /// </summary>
        public static void PreloadSQLite()
        {
            try
            {
                string baseDir = null;

                // ESTRATEGIA 1 (Robusta Excel-DNA): Ruta real del .XLL
                // Esto evade el shadow-copy del CLR que rompe Assembly.Location.
                try 
                { 
                    string xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                    if (!string.IsNullOrEmpty(xllPath))
                    {
                        baseDir = Path.GetDirectoryName(xllPath);
                        LoggerHelper.Info($"[NativePreloader] Resolucin va XLL: {baseDir}");
                    }
                }
                catch { /* No estamos en contexto Excel o ExcelDna no disponible */ }

                // ESTRATEGIA 2 (Fallback seguro): BaseDirectory del AppDomain
                // Funciona en entornos corporativos con restricciones de shadow-copy.
                if (string.IsNullOrEmpty(baseDir))
                {
                    baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    LoggerHelper.Info($"[NativePreloader] Resolucin va AppDomain.BaseDirectory: {baseDir}");
                }

                // ESTRATEGIA 3 (ltimo recurso): Assembly.Location (slo si no es shadow-copy vaco)
                if (string.IsNullOrEmpty(baseDir))
                {
                    string location = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    if (!string.IsNullOrWhiteSpace(location))
                    {
                        baseDir = Path.GetDirectoryName(location);
                    }
                }

                if (string.IsNullOrEmpty(baseDir))
                {
                    LoggerHelper.Error("[NativePreloader] No se pudo determinar directorio base para preloading.");
                    return;
                }

                string arch = Environment.Is64BitProcess ? "x64" : "x86";
                string interopPath = Path.Combine(baseDir, arch, "SQLite.Interop.dll");

                if (!File.Exists(interopPath))
                {
                    interopPath = Path.Combine(baseDir, "SQLite.Interop.dll");
                }

                if (File.Exists(interopPath))
                {
                    LoggerHelper.Info($"[NativePreloader] Forzando carga de SQLite nativo desde: {interopPath}");
                    IntPtr handle = LoadLibrary(interopPath);
                    if (handle == IntPtr.Zero)
                    {
                        int error = Marshal.GetLastWin32Error();
                        LoggerHelper.Warning($"[NativePreloader] Fall LoadLibrary. Cdigo Win32: {error}");
                    }
                    else
                    {
                        LoggerHelper.Info("[NativePreloader] SQLite nativo cargado exitosamente.");
                    }
                }
                else
                {
                    LoggerHelper.Warning($"[NativePreloader] No se encontr SQLite.Interop.dll en {baseDir} (buscado en raz y {arch})");
                }
            }
            catch (Exception ex)
            {
                // Guard defensivo: nunca dejar crashear la inicializacin por un fallo de preloading.
                LoggerHelper.Error("[NativePreloader] Error durante el preloading nativo.", ex);
            }
        }
    }
}
