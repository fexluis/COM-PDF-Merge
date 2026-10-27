using System;
using System.IO;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Helpers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Unit.Core
{
    /// <summary>
    /// Pruebas de no regresión para la redirección dinámica de DirectorioLog en LoggerHelper (Fase 1069).
    /// </summary>
    public class LoggerHelperRedirectionTests
    {
        [Fact]
        public void DirectorioLog_WhenRedirigido_ShouldEscribirEnCarpetaDestino()
        {
            // Arrange
            string directorioOriginal = LoggerHelper.DirectorioLog;
            NivelLog minimoOriginal = LoggerHelper.Minimo;
            string destino = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests.Log_" + Guid.NewGuid().ToString("N"));
            string marcador = "redireccion-" + Guid.NewGuid().ToString("N");

            try
            {
                // Act
                LoggerHelper.Minimo = NivelLog.Info;
                LoggerHelper.DirectorioLog = destino;
                LoggerHelper.Info(marcador);

                // Assert
                string archivo = Path.Combine(destino, $"log_{DateTime.Now:yyyyMMdd}.txt");
                string contenido = EsperarContenido(archivo, marcador, TimeSpan.FromSeconds(5));
                contenido.Should().NotBeNull();
                contenido.Should().Contain(marcador);
            }
            finally
            {
                LoggerHelper.Minimo = minimoOriginal;
                LoggerHelper.DirectorioLog = directorioOriginal;
                try
                {
                    Directory.Delete(destino, true);
                }
                catch { /* Limpieza tolerante */ }
            }
        }

        private static string EsperarContenido(string archivo, string marcador, TimeSpan timeout)
        {
            DateTime limite = DateTime.UtcNow.Add(timeout);
            while (DateTime.UtcNow < limite)
            {
                try
                {
                    if (File.Exists(archivo))
                    {
                        string texto = File.ReadAllText(archivo);
                        if (texto.Contains(marcador))
                        {
                            return texto;
                        }
                    }
                }
                catch (IOException) { }
                System.Threading.Thread.Sleep(100);
            }
            return null;
        }
    }
}
