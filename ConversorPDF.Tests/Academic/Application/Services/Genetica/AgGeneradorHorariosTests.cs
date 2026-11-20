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
    [Collection("AgGeneradorHorarios")]
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
            result.Error.Should().Contain("bloqueos de factibilidad detectados");
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

        #region Pruebas de Integración de Reparación de Campeón (Fase 2133)

        [Fact]
        public async Task EjecutarOptimizacionAsync_CampeonFinal_EntregaFechasSecuencialesEInvariantesI1I2()
        {
            // Arrange
            var mockVentana = new Mock<IAgValidadorVentana>();
            mockVentana.Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(Result.Success(new InformeFactibilidad()));

            var validadorReglas = new AgValidadorReglas();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();
            mockAptitud.Setup(a => a.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                       .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                       {
                           validadorReglas.Evaluar(ind, ctx);
                           ind.Fitness = 1.0 / (1.0 + ind.HardViolations);
                       });

            var contexto = new AgContexto();
            contexto.IdSede = 1;
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;

            var aula = new AgAula { Id = 1, IdSede = 1, IdPrograma = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30, Semestre = "I", Grupo = "A", Jornada = AgFranjaJornada.MANANA, DiasVentana = "LUN-VIE" };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            var asig1 = new Asignatura { Id = 101, Nombre = "Materia 1", Horas = 16, Semestre = "I" };
            var asig2 = new Asignatura { Id = 102, Nombre = "Materia 2", Horas = 16, Semestre = "I" };
            contexto.Asignaturas.Add(asig1);
            contexto.Asignaturas.Add(asig2);
            contexto.AsignaturasPorId[101] = asig1;
            contexto.AsignaturasPorId[102] = asig2;

            // Fase 2144: Poblar AulaAsignaturas en la lista canónica para que sobrevivan a ConstruirIndices()
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 101, OrdenCurricular = 1, DiasPatron = "1,2" });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 102, OrdenCurricular = 2, DiasPatron = "1,2" });
            contexto.ConstruirIndices();

            var b1 = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" };
            var b2 = new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" };
            contexto.Bloques.Add(b1);
            contexto.Bloques.Add(b2);
            contexto.BloquesPorId[1] = b1;
            contexto.BloquesPorId[2] = b2;

            var doc = new Tercero { Id = 10, Nombre = "Docente A", Activo = 1 };
            contexto.Docentes.Add(doc);
            contexto.DocentesPorId[10] = doc;
            contexto.RegistrarDocenteHabilitado(10, 1, 101, "I");
            contexto.RegistrarDocenteHabilitado(10, 1, 102, "I");

            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object, seed: 42);

            // Act
            var result = await sut.EjecutarOptimizacionAsync(contexto, 4, 2, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.MejorIndividuo.Should().NotBeNull();
            result.Value.MejorIndividuo.HardViolations.Should().Be(0);

            var genes101 = result.Value.MejorIndividuo.Genes.Where(g => g.IdAsignatura == 101 && g.FechaFinEstimada.HasValue).ToList();
            var genes102 = result.Value.MejorIndividuo.Genes.Where(g => g.IdAsignatura == 102 && g.FechaInicioEstimada.HasValue).ToList();

            genes101.Should().NotBeEmpty();
            genes102.Should().NotBeEmpty();
            var fin101 = genes101.Max(g => g.FechaFinEstimada.Value);
            var ini102 = genes102.Min(g => g.FechaInicioEstimada.Value);
            ini102.Should().BeAfter(fin101);

            // Fase 2160: Verificación estricta de invariantes I1/I2 y unicidad de fechas
            foreach (var grp in result.Value.MejorIndividuo.Genes.GroupBy(g => new { g.ClaveGrupo, g.IdAsignatura }))
            {
                grp.Sum(g => g.HorasGen).Should().Be(grp.Max(g => g.TotalHorasAsignatura));
                foreach (var g in grp)
                {
                    g.FechasGen.Count.Should().Be((int)Math.Round(g.HorasGen / 4.0));
                    g.FechasGen.Should().OnlyHaveUniqueItems();
                }
            }
        }

        #endregion
    }
}
