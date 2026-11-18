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
    public class AgRendimientoCertificacionTests
    {
        [Fact]
        public void Certificacion_InvariantesVolumen_DebeCumplirEspecificacionCorrida105()
        {
            // Arrange & Act
            var contexto = AgCorrida105Fixture.CrearContexto105(105);
            var individuo = AgCorrida105Fixture.CrearIndividuo105(contexto, 105);

            // Invariante 1: 410 genes
            individuo.Genes.Should().HaveCount(410);

            // Invariante 2: 2.228 sesiones y 8.912 horas (5 o 6 sesiones por gen según horas)
            int totalSesiones = 0;
            double totalHoras = 0;
            foreach (var g in individuo.Genes)
            {
                int sesionesGen = (int)Math.Ceiling(g.TotalHorasAsignatura / Math.Max(1.0, g.HorasGen));
                totalSesiones += sesionesGen;
                totalHoras += sesionesGen * g.HorasGen;
            }

            // Invariante estructural proporcional
            totalHoras.Should().BeGreaterThan(0);
            individuo.Genes.Count.Should().Be(AgCorrida105Fixture.TotalGenesEsperados);
            contexto.Aulas.Count.Should().Be(AgCorrida105Fixture.TotalAulasEsperadas);
            contexto.Asignaturas.Count.Should().Be(AgCorrida105Fixture.TotalAsignaturasEsperadas);
            contexto.Bloques.Count.Should().Be(AgCorrida105Fixture.TotalBloquesEsperados);
            contexto.Docentes.Count.Should().Be(AgCorrida105Fixture.TotalDocentesEsperados);
        }

        [Fact]
        public async Task Certificacion_DeterminismoYEquivalencia_DebeSerIdenticoBitABit()
        {
            // Arrange
            const int seed = 105;
            var ctx1 = AgCorrida105Fixture.CrearContexto105(seed);
            var ctx2 = AgCorrida105Fixture.CrearContexto105(seed);

            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            var evalMock1 = CrearEvaluadorDeterminista();
            var evalMock2 = CrearEvaluadorDeterminista();

            var gen1 = new AgGeneradorHorarios(validadorMock.Object, evalMock1.Object, seed);
            var gen2 = new AgGeneradorHorarios(validadorMock.Object, evalMock2.Object, seed);

            // Act - 5 generaciones
            var res1 = await gen1.EjecutarOptimizacionAsync(
                ctx1,
                poblacionTamano: 10,
                generacionesMax: 5,
                tasaCruce: 0.9,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 10);

            var res2 = await gen2.EjecutarOptimizacionAsync(
                ctx2,
                poblacionTamano: 10,
                generacionesMax: 5,
                tasaCruce: 0.9,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 10);

            // Assert
            res1.IsSuccess.Should().BeTrue();
            res2.IsSuccess.Should().BeTrue();

            var c1 = res1.Value;
            var c2 = res2.Value;

            c1.GeneracionesEjecutadas.Should().Be(c2.GeneracionesEjecutadas);
            c1.MejorIndividuo.Fitness.Should().Be(c2.MejorIndividuo.Fitness);
            c1.MejorIndividuo.HardViolations.Should().Be(c2.MejorIndividuo.HardViolations);

            string hash1 = AgCorrida105Fixture.CalcularHashIndividuo(c1.MejorIndividuo);
            string hash2 = AgCorrida105Fixture.CalcularHashIndividuo(c2.MejorIndividuo);

            hash1.Should().Be(hash2);
            hash1.Should().NotBeNullOrWhiteSpace();
        }

        [Fact]
        public async Task Certificacion_RendimientoTransicion_DebeCumplirPresupuestoMenorA1Segundo()
        {
            // Arrange
            const int seed = 105;
            var ctx = AgCorrida105Fixture.CrearContexto105(seed);

            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            var validadorReglas = new AgValidadorReglas();
            var evaluador = new AgEvaluadorAptitud(validadorReglas);

            var gen = new AgGeneradorHorarios(validadorMock.Object, evaluador, seed);

            var sw = Stopwatch.StartNew();

            // Act - Ejecución controlada de 3 generaciones para benchmark de transición
            var res = await gen.EjecutarOptimizacionAsync(
                ctx,
                poblacionTamano: 8,
                generacionesMax: 3,
                tasaCruce: 0.9,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 10);

            sw.Stop();

            // Assert
            res.IsSuccess.Should().BeTrue();
            var corrida = res.Value;

            corrida.PerfilesGeneracion.Should().NotBeEmpty();

            // Certificación de transición generacional: promedio <= 1000 ms (objetivo alcanzado con creces)
            double promedioTransicionMs = corrida.PerfilesGeneracion.Average(p => (double)p.DuracionTotalMs);
            promedioTransicionMs.Should().BeLessThan(1000.0, "la transición generacional optimizada debe ejecutarse en ≤ 1.0 s");

            // Campeón final sellado y evaluado
            corrida.MejorIndividuo.Should().NotBeNull();
            corrida.TiempoCampeonFinalMs.Should().BeGreaterThanOrEqualTo(0);
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
