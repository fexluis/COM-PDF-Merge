using System;
using System.IO;
using System.Runtime.CompilerServices;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests
{
    /// <summary>
    /// Bootstrap del ensamblado de pruebas que redirige la salida de LoggerHelper a %TEMP% (Fase 1068).
    /// Se ejecuta automáticamente mediante el inicializador de módulo al cargar el ensamblado de pruebas,
    /// previniendo la contaminación de los logs de producción en %LOCALAPPDATA%.
    /// </summary>
    internal static class TestLogBootstrap
    {
        [ModuleInitializer]
        internal static void Inicializar()
        {
            try
            {
                LoggerHelper.DirectorioLog =
                    Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", "log");
            }
            catch { /* Fallback silencioso si no es posible redirigir */ }
        }
    }
}
