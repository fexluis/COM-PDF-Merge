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
    public class AgGeneradorHorariosTests
    {
        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenContextIsNull_ShouldReturnFailure()
        {
            // Arrange
            var mockVentana = new Mock<IAgValidadorVentana>();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();
            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);

            // Act
            var result = await sut.EjecutarOptimizacionAsync(null, 10, 5, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("contexto");
        }

        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenPreFlightInfeasible_ShouldReturnInfeasibleResult()
        {
            // Arrange
            var mockVentana = new Mock<IAgValidadorVentana>();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();

            var informe = new InformeFactibilidad();
            informe.AgregarBloqueo("Ventana bloqueada sin semanas suficientes.");
            mockVentana.Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(Result.Success(informe));

            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);
            var contexto = new AgContexto();

            // Act
            var result = await sut.EjecutarOptimizacionAsync(contexto, 10, 5, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("inviabilidad");
        }

        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenChampionPublished_ShouldBeSealedAndPreserveGroupInvariants()
        {
            // Arrange (Fase 582)
            var mockVentana = new Mock<IAgValidadorVentana>();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();

            var informe = new InformeFactibilidad(); // EsFactible = true
            mockVentana.Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(Result.Success(informe));

            mockAptitud.Setup(a => a.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                       .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                       {
                           ind.Fitness = 1.0;
                           ind.HardViolations = 0;
                       });

            var contexto = new AgContexto();
            contexto.IdSede = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var aula = new AgAula { Id = 1, IdSede = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30, Semestre = "I" };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            var asig = new Asignatura { Id = 10, Nombre = "Matemáticas", Horas = 48, Semestre = "I" };
            contexto.Asignaturas.Add(asig);
            contexto.AsignaturasPorId[10] = asig;

            var bloque = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 };
            contexto.Bloques.Add(bloque);
            contexto.BloquesPorId[1] = bloque;

            var doc = new Tercero { Id = 100, Nombre = "Docente A", Activo = 1 };
            contexto.Docentes.Add(doc);
            contexto.DocentesPorId[100] = doc;

            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);

            // Act
            var result = await sut.EjecutarOptimizacionAsync(contexto, 6, 2, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.MejorIndividuo.Should().NotBeNull();
            result.Value.MejorIndividuo.Genes.Should().NotBeEmpty();
            result.Value.MejorIndividuo.Genes
                .GroupBy(g => g.ClaveGrupo)
                .All(grp => Math.Abs(grp.Sum(g => g.HorasGen) - grp.Max(g => g.TotalHorasAsignatura)) < 0.001)
                .Should().BeTrue();
        }
    }
}
