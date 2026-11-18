using System;
using System.Collections.Generic;
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
    public class AgPerfilGeneracionTests
    {
        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenExecuted_ShouldPopulateTelemetryAndPerfilesGeneracion()
        {
            // Arrange
            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            var evaluadorMock = new Mock<IAgEvaluadorAptitud>();
            int evalCount = 0;
            evaluadorMock
                .Setup(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                {
                    evalCount++;
                    ind.Fitness = 100.0 + (evalCount % 10);
                    ind.HardViolations = 0;
                });

            var sut = new AgGeneradorHorarios(validadorMock.Object, evaluadorMock.Object, seed: 42);
            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1 });

            // Act
            var result = await sut.EjecutarOptimizacionAsync(
                contexto,
                poblacionTamano: 6,
                generacionesMax: 3,
                tasaCruce: 1.0,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 50);

            // Assert
            result.IsSuccess.Should().BeTrue();
            var corrida = result.Value;
            corrida.HistorialConvergencia.Should().NotBeEmpty();

            // Bitacoras
            foreach (var b in corrida.HistorialConvergencia)
            {
                b.PeorFitness.Should().BeLessThanOrEqualTo(b.MejorFitness);
                b.Evaluados.Should().BeGreaterThan(0);
                b.DuracionMs.Should().BeGreaterThanOrEqualTo(0);
            }

            // Perfiles
            corrida.PerfilesGeneracion.Should().NotBeEmpty();
            foreach (var perfil in corrida.PerfilesGeneracion)
            {
                perfil.EvaluacionesAptitud.Should().BeGreaterThan(0);
                perfil.HijosCreados.Should().BeGreaterThan(0);
                perfil.HijosIncorporados.Should().BeGreaterThan(0);
                perfil.ElitesClonados.Should().Be(2);
                perfil.DuracionTotalMs.Should().BeGreaterThanOrEqualTo(0);
            }

            // Subtiempos de inicialización registrados
            corrida.TiempoPoblacionInicialMs.Should().BeGreaterThanOrEqualTo(0);
            corrida.TiempoCampeonFinalMs.Should().BeGreaterThanOrEqualTo(0);
        }

        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenOddPopulation_ShouldTrackEvaluacionesExtraPorCupoImpar()
        {
            // Arrange
            var validadorMock = new Mock<IAgValidadorVentana>();
            validadorMock
                .Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new InformeFactibilidad { EsFactible = true }));

            var evaluadorMock = new Mock<IAgEvaluadorAptitud>();
            evaluadorMock
                .Setup(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                {
                    ind.Fitness = 50.0;
                    ind.HardViolations = 0;
                });

            // Población de 5 con elitismo 2 => faltan 3 cupos (impar: crea par 1 y 2, incorpora ambos; luego crea par 3 y 4, incorpora solo el 3)
            var sut = new AgGeneradorHorarios(validadorMock.Object, evaluadorMock.Object, seed: 123);
            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1 });

            // Act
            var result = await sut.EjecutarOptimizacionAsync(
                contexto,
                poblacionTamano: 5,
                generacionesMax: 2,
                tasaCruce: 1.0,
                tasaMutacion: 0.1,
                elitismo: 2,
                progreso: null,
                ct: CancellationToken.None,
                estancamientoMax: 50);

            // Assert
            result.IsSuccess.Should().BeTrue();
            var corrida = result.Value;
            corrida.PerfilesGeneracion.Should().NotBeEmpty();
            var primerPerfil = corrida.PerfilesGeneracion[0];
            primerPerfil.EvaluacionesExtraPorCupoImpar.Should().Be(1);
            primerPerfil.HijosCreados.Should().Be(4); // 2 pares evaluados
            primerPerfil.HijosIncorporados.Should().Be(3); // 3 cupos llenados
        }
    }
}
