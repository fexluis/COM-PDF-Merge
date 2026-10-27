using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgDiagnosticoTrazaTests
    {
        [Fact]
        public void Paso_WhenCalledMultipleTimes_ShouldIncrementStepNumberSequentially()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaPaso", msg => lineas.Add(msg), "TEST-01") { Habilitado = true };

            // Act
            sut.Paso("PrimerPaso");
            sut.Paso("SegundoPaso");
            sut.Paso("TercerPaso");

            // Assert
            lineas.Should().HaveCount(3);
            lineas[0].Should().Contain("[P01] PrimerPaso");
            lineas[1].Should().Contain("[P02] SegundoPaso");
            lineas[2].Should().Contain("[P03] TercerPaso");
        }

        [Fact]
        public void Paso_WhenHabilitadoIsFalse_ShouldNotEmitToUi()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaDeshabilitado", msg => lineas.Add(msg), "TEST-02")
            {
                Habilitado = false
            };

            // Act
            sut.Paso("PasoSilencioso");
            sut.Dato("Clave", "Valor");
            sut.Sql("Consulta", "SELECT 1");

            // Assert
            lineas.Should().BeEmpty();
        }

        [Fact]
        public void Conteo_WhenCountBelowMinimum_ShouldEmitAlertMessage()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaConteo", msg => lineas.Add(msg), "TEST-03");

            // Act
            sut.Conteo("Asignaturas", 0, minimo: 1);

            // Assert
            lineas.Should().ContainSingle();
            lineas[0].Should().Contain("ALERTA (<1)");
            lineas[0].Should().Contain("Asignaturas: 0");
        }

        [Fact]
        public void Conteo_WhenCountAtOrAboveMinimum_ShouldEmitOkMessage()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaConteoOk", msg => lineas.Add(msg), "TEST-04") { Habilitado = true };

            // Act
            sut.Conteo("Aulas", 5, minimo: 1);

            // Assert
            lineas.Should().ContainSingle();
            lineas[0].Should().Contain("5 OK");
        }

        [Fact]
        public void Sql_WhenMultilineWithSpaces_ShouldNormalizeSqlLine()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaSql", msg => lineas.Add(msg), "TEST-05") { Habilitado = true };
            string sqlSucio = @"
                SELECT   id,    codigo 
                FROM     ag_aula 
                WHERE    estado = 1";

            // Act
            sut.Sql("BuscarAulas", sqlSucio, new { Sede = 2 });

            // Assert
            lineas.Should().ContainSingle();
            lineas[0].Should().Contain("SQL[BuscarAulas]: SELECT id, codigo FROM ag_aula WHERE estado = 1");
            lineas[0].Should().Contain("Sede=2");
        }

        [Fact]
        public void Causa_WhenSpecified_ShouldRecordLastCauseAndEmitDebug()
        {
            // Arrange
            var lineas = new List<string>();
            var sut = new AgDiagnosticoTraza("PruebaCausa", msg => lineas.Add(msg), "TEST-06");

            // Act
            sut.Causa(AgCausaFallo.SinAulaAsignaturas, "Aulas sin asignaturas vinculadas");

            // Assert
            sut.UltimaCausa.Should().Be(AgCausaFallo.SinAulaAsignaturas);
            lineas.Should().ContainSingle();
            lineas[0].Should().Contain("CAUSA: SinAulaAsignaturas");
        }

        [Fact]
        public void Inicio_WhenHabilitadoFalse_ShouldStillNotifyUi()
        {
            // Arrange (Fase 604): con Habilitado = false, las etapas INFO/WARN deben llegar a la UI, pero el detalle DEBUG no
            var recibidos = new List<string>();
            var traza = new AgDiagnosticoTraza("Prueba", recibidos.Add) { Habilitado = false };

            // Act
            traza.Inicio();
            traza.Advertencia("X", "y");
            traza.Dato("k", 1);

            // Assert: Inicio() (INFO) + Advertencia() (WARN) se notifican; Dato() (DEBUG) se omite
            recibidos.Should().HaveCount(2);
        }
    }
}

