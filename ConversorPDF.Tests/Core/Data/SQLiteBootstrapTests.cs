using System;
using System.IO;
using Xunit;
using FluentAssertions;
using ConversorPDF.XLL.Core.Data;

namespace ConversorPDF.Tests.Core.Data
{
    /// <summary>
    /// Tests de contrato defensivos para el bootstrap de SQLite bajo el entorno Excel-DNA.
    /// Estos tests garantizan que los workarounds para el bug de Shadow Copy de System.Data.SQLite
    /// no sean eliminados accidentalmente por refactors futuros.
    ///
    /// TRASFONDO: Excel-DNA coloca los assemblies en rutas shadow-copy temporales tipo:
    ///   %TEMP%\{guid}\assembly\dl3\{hash}\System.Data.SQLite.dll
    /// y SQLiteConnection..ctor intenta buscar su .config en esa ruta.
    /// Si no existe, lanza ArgumentNullException("path1").
    /// Ver: .agents/rules/sqlite-excel-dna-gotchas.md
    /// </summary>
    public class SQLiteBootstrapTests
    {
        // =====================================================================
        // CONTRATO 1: Idempotencia de EnsureInitialized
        // =====================================================================

        [Fact]
        public void EnsureInitialized_ShouldNotThrow()
        {
            // Arrange / Act
            Action act = () => SQLiteEngineBootstrap.EnsureInitialized();

            // Assert
            act.Should().NotThrow("el bootstrap debe ser resiliente a cualquier condición de entorno");
        }

        [Fact]
        public void EnsureInitialized_CalledTenTimes_ShouldNotThrow()
        {
            // Act
            Action act = () =>
            {
                for (int i = 0; i < 10; i++)
                {
                    SQLiteEngineBootstrap.EnsureInitialized();
                }
            };

            // Assert — idempotencia en el mismo AppDomain
            act.Should().NotThrow("debe ser seguro llamarlo múltiples veces sin estado residual");
        }
    }
}
