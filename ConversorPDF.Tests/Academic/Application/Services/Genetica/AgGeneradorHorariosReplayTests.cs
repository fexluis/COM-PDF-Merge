using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgGeneradorHorariosReplayTests
    {
        [Fact]
        public void Fixture105_WhenCreated_ShouldMatchStructuralVolumes()
        {
            // Act
            var contexto = AgCorrida105Fixture.CrearContexto105();
            var individuo = AgCorrida105Fixture.CrearIndividuo105(contexto);

            // Assert
            contexto.Aulas.Should().HaveCount(AgCorrida105Fixture.TotalAulasEsperadas);
            contexto.Asignaturas.Should().HaveCount(AgCorrida105Fixture.TotalAsignaturasEsperadas);
            contexto.Bloques.Should().HaveCount(AgCorrida105Fixture.TotalBloquesEsperados);
            contexto.Docentes.Should().HaveCount(AgCorrida105Fixture.TotalDocentesEsperados);
            individuo.Genes.Should().HaveCount(AgCorrida105Fixture.TotalGenesEsperados);
        }

        [Fact]
        public async Task Replay_WhenSameSeed_ShouldProduceDeterministicTrajectoryAndHash()
        {
            // Arrange
            const int seed = 105;
            var contexto1 = AgCorrida105Fixture.CrearContexto105(seed);
            var contexto2 = AgCorrida105Fixture.CrearContexto105(seed);

            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            // Evaluador determinista basado en genes
            var evaluadorMock1 = CrearEvaluadorDeterminista();
            var evaluadorMock2 = CrearEvaluadorDeterminista();

            var sut1 = new AgGeneradorHorarios(validadorMock.Object, evaluadorMock1.Object, seed);
            var sut2 = new AgGeneradorHorarios(validadorMock.Object, evaluadorMock2.Object, seed);

            var sw = Stopwatch.StartNew();

            // Act - Ejecutar 5 generaciones con población 8
            var res1 = await sut1.EjecutarOptimizacionAsync(
                contexto1,
                poblacionTamano: 8,
                generacionesMax: 5,
                tasaCruce: 0.9,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 20);

            var res2 = await sut2.EjecutarOptimizacionAsync(
                contexto2,
                poblacionTamano: 8,
                generacionesMax: 5,
                tasaCruce: 0.9,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 20);

            sw.Stop();

            // Assert
            res1.IsSuccess.Should().BeTrue();
            res2.IsSuccess.Should().BeTrue();

            var corrida1 = res1.Value;
            var corrida2 = res2.Value;

            corrida1.GeneracionesEjecutadas.Should().Be(corrida2.GeneracionesEjecutadas);
            corrida1.HistorialConvergencia.Select(b => b.MejorFitness)
                .Should().Equal(corrida2.HistorialConvergencia.Select(b => b.MejorFitness));

            corrida1.HistorialConvergencia.Select(b => b.PromedioFitness)
                .Should().Equal(corrida2.HistorialConvergencia.Select(b => b.PromedioFitness));

            string hash1 = AgCorrida105Fixture.CalcularHashIndividuo(corrida1.MejorIndividuo);
            string hash2 = AgCorrida105Fixture.CalcularHashIndividuo(corrida2.MejorIndividuo);

            hash1.Should().NotBeNullOrWhiteSpace();
            hash1.Should().Be(hash2);

            // Verificación de presupuesto de prueba de CI
            sw.ElapsedMilliseconds.Should().BeLessThan(10000, "el arnés de replay debe ser eficiente para ejecución continua");
        }

        [Fact]
        public async Task Replay_WhenDifferentSeed_ShouldDiverge()
        {
            // Arrange
            var contexto1 = AgCorrida105Fixture.CrearContexto105(105);
            var contexto2 = AgCorrida105Fixture.CrearContexto105(999);

            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            var sut1 = new AgGeneradorHorarios(validadorMock.Object, CrearEvaluadorDeterminista().Object, seed: 105);
            var sut2 = new AgGeneradorHorarios(validadorMock.Object, CrearEvaluadorDeterminista().Object, seed: 999);

            // Act
            var res1 = await sut1.EjecutarOptimizacionAsync(
                contexto1,
                poblacionTamano: 8,
                generacionesMax: 4,
                tasaCruce: 0.9,
                tasaMutacion: 0.5,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None);

            var res2 = await sut2.EjecutarOptimizacionAsync(
                contexto2,
                poblacionTamano: 8,
                generacionesMax: 4,
                tasaCruce: 0.9,
                tasaMutacion: 0.5,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None);

            // Assert
            string hash1 = AgCorrida105Fixture.CalcularHashIndividuo(res1.Value.MejorIndividuo);
            string hash2 = AgCorrida105Fixture.CalcularHashIndividuo(res2.Value.MejorIndividuo);
            hash1.Should().NotBe(hash2);
        }

        private static Mock<IAgEvaluadorAptitud> CrearEvaluadorDeterminista()
        {
            var mock = new Mock<IAgEvaluadorAptitud>();
            mock.Setup(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                {
                    if (ind == null || ind.Genes == null) return;
                    int vacantes = ind.Genes.Count(g => g.EsVacante == 1 || g.IdTercero <= 0);
                    int fijos = ind.Genes.Count(g => g.EsFijo == 1);
                    int diversidad = ind.Genes.Select(g => g.IdTercero * 37 + g.IdBloqueHorario).Distinct().Count();
                    ind.HardViolations = vacantes > 20 ? 1 : 0;
                    ind.Fitness = 1000.0 - (vacantes * 5.0) + (fijos * 2.0) + (diversidad * 3.5);
                });
            return mock;
        }
    }
}
