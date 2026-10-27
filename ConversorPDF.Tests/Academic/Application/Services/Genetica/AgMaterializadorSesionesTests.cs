using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgMaterializadorSesionesTests
    {
        [Fact]
        public async Task MaterializarEnStagingAsync_WhenTwoGenesShareSubjectTeacherRoomWindow_ShouldGroupIntoSingleVinculacion()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();

            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));
            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            List<AgSesion> savedSesiones = null;

            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                        savedSesiones = s.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) => Result.Success(new List<AgSesion>
                        {
                            new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 4 }
                        }));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);

            var individuo = new AgIndividuo { Id = 10 };
            var fIni = new DateTime(2026, 2, 1);
            var fFin = new DateTime(2026, 3, 1);

            // 2 genes de la misma asignatura/docente/aula/ventana pero diferente bloque horario (e.g. 2 días/semana)
            individuo.Genes.Add(new AgGen { Id = 101, IdSede = 2, IdAsignatura = 1, IdTercero = 10, IdAula = 1, IdBloqueHorario = 1, TotalHorasAsignatura = 8.0, HorasGen = 4.0, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });
            individuo.Genes.Add(new AgGen { Id = 102, IdSede = 2, IdAsignatura = 1, IdTercero = 10, IdAula = 1, IdBloqueHorario = 2, TotalHorasAsignatura = 8.0, HorasGen = 4.0, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });

            var contexto = new AgContexto();
            contexto.IdSede = 2;
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, DuracionHoras = 4, HoraInicio = "08:00", HoraFin = "12:00" };
            contexto.BloquesPorId[2] = new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, DuracionHoras = 4, HoraInicio = "08:00", HoraFin = "12:00" };

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            savedVincs.Should().NotBeNull();
            savedVincs.Should().HaveCount(1, "ambos genes deben agruparse en una sola vinculación");
            savedVincs[0].IdSede.Should().Be(2);
            savedVincs[0].GenesOrigen.Should().Contain(new long[] { 101, 102 });
            savedSesiones.Should().HaveCount(2);
            savedSesiones.Should().OnlyContain(s => s.IdVinculacion == savedVincs[0].Id);
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenIdSedeCannotBeResolved_ShouldReturnFailure()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();
            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);

            var individuo = new AgIndividuo { Id = 10 };
            individuo.Genes.Add(new AgGen { Id = 101, IdSede = 0, IdAsignatura = 1, IdTercero = 10, IdAula = 1, IdBloqueHorario = 1 });
            var contexto = new AgContexto { IdSede = 0 };
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1 };

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("IdSede <= 0");
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenGenesIsEmpty_ShouldReturnFailure()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();
            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);

            var individuo = new AgIndividuo { Id = 10 };
            var contexto = new AgContexto();

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Cromosoma vacío");
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenSingleGroupViolatesI1_ShouldQuarantineAndFailStagingVacioWithDiagnostic()
        {
            // Arrange (Fase 581, Fase 584 & Fase 592)
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();
            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);

            var individuo = new AgIndividuo { Id = 10 };
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 51,
                IdAula = 285,
                Grupo = "A",
                ClaveGrupo = "51|285|A",
                TotalHorasAsignatura = 48.0,
                HorasGen = 96.0,
                EsFijo = 1,          // Fase 592: gen fijado => Sellar no puede reparar Σ
                Estado = ConversorPDF.XLL.Academic.Domain.Enums.AgEstadoGen.Fijado,
                IdSede = 1
            });

            var contexto = new AgContexto { IdSede = 1 };

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("cuarentena");
            result.Error.Should().Contain("clave='51|285|A'");
            result.Error.Should().Contain("genes=1");
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenOneGroupValidAndOneInvalid_ShouldQuarantineInvalidAndPersistValid()
        {
            // Arrange (Fase 584)
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();
            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) => Result.Success(new List<AgSesion>
                        {
                            new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 4 }
                        }));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);
            var contexto = new AgContexto { IdSede = 1 };
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1 };

            var individuo = new AgIndividuo { Id = 10 };
            // Grupo A válido (8h, 2 genes de 4h c/u)
            individuo.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", TotalHorasAsignatura = 8.0, HorasGen = 4.0, IdSede = 1, IdBloqueHorario = 1 });
            individuo.Genes.Add(new AgGen { Id = 2, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", TotalHorasAsignatura = 8.0, HorasGen = 4.0, IdSede = 1, IdBloqueHorario = 1 });

            // Grupo B inválido (Σ=96, tot=48)
            individuo.Genes.Add(new AgGen { Id = 3, IdAsignatura = 20, IdAula = 2, Grupo = "B", ClaveGrupo = "20|2|B", TotalHorasAsignatura = 48.0, HorasGen = 96.0, IdSede = 1, IdBloqueHorario = 1 });

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            savedVincs.Should().NotBeNull();
            savedVincs.Should().HaveCount(1);
            savedVincs[0].Grupo.Should().Be("A");
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenGroupHasDefensiveCoherenceNeeded_ShouldSealAndMaterializeSuccessfully()
        {
            // Arrange (Fase 583 & Fase 589)
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();
            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            List<AgSesion> savedSesiones = null;
            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                        savedSesiones = s.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) =>
                        {
                            int count = (int)Math.Round(g.HorasGen / 4.0);
                            var list = new List<AgSesion>();
                            for (int i = 0; i < count; i++) list.Add(new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 4 });
                            return Result.Success(list);
                        });

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);
            var contexto = new AgContexto { IdSede = 1 };
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1 };

            var individuo = new AgIndividuo { Id = 10 };
            // 2 genes mismo grupo con Σ=96 pero TotalHorasAsignatura=48 -> Sellar en guarda Fase 583 repara a 24+24=48 (6 sesiones c/u = 12 total)
            individuo.Genes.Add(new AgGen { Id = 1, IdAsignatura = 30, IdAula = 1, Grupo = "A", ClaveGrupo = "30|1|A", TotalHorasAsignatura = 48.0, HorasGen = 48.0, IdSede = 1, IdBloqueHorario = 1 });
            individuo.Genes.Add(new AgGen { Id = 2, IdAsignatura = 30, IdAula = 1, Grupo = "A", ClaveGrupo = "30|1|A", TotalHorasAsignatura = 48.0, HorasGen = 48.0, IdSede = 1, IdBloqueHorario = 1 });

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            savedVincs.Should().NotBeNull();
            savedVincs.Should().HaveCount(1);
            savedVincs[0].TotalHoras.Should().Be(48.0); // Fase 589: vinculación mapeada con horas reparadas
            savedSesiones.Should().NotBeNull();
            savedSesiones.Should().HaveCount(12);
        }

        [Fact]
        public async Task MaterializarEnStagingAsync_WhenGeneHas8HoursSession_ShouldProduce3PartDiaHorasAndPassI2()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();

            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            List<AgSesion> savedSesiones = null;

            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                        savedSesiones = s.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) => Result.Success(new List<AgSesion>
                        {
                            new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 8.0 }
                        }));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);
            var individuo = new AgIndividuo { Id = 10 };
            var fIni = new DateTime(2026, 2, 1);
            var fFin = new DateTime(2026, 3, 1);

            var gen = new AgGen
            {
                Id = 201,
                IdSede = 1,
                IdAsignatura = 10,
                IdTercero = 5,
                IdAula = 1,
                IdBloqueHorario = 10,
                DiaSemana = 2,
                HorasSemanales = 8.0,
                TotalHorasAsignatura = 8.0,
                HorasGen = 8.0,
                FechaInicioEstimada = fIni,
                FechaFinEstimada = fFin
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto { IdSede = 1 };
            contexto.BloquesPorId[10] = new AgBloqueHorario
            {
                Id = 10,
                DiaSemanaNumero = 2,
                DuracionHoras = 8.0,
                HoraInicio = "08:00",
                HoraFin = "17:00",
                Franja = "MANANA"
            };

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            savedVincs.Should().HaveCount(1);
            savedVincs[0].DiaHorasSemanales.Should().Be("MARTES|8|08:00 A 17:00");
            savedSesiones.Should().HaveCount(1);
            savedSesiones[0].Horas.Should().Be(8.0);
        }

        [Fact]
        public async Task Materializar_ConOleada_NoSobrescribeFechasConVentana()
        {
            // Arrange (Fase 1108 / Fase 1112): El gen tiene sub-ventana de oleada (2026-03-01..2026-04-15)
            // y el contexto posee una ventana semestral más amplia (2026-02-01..2026-06-30 con IdPeriodo > 0).
            // MaterializarEnStagingAsync debe conservar las fechas de la oleada y no sobrescribirlas con la ventana completa.
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();

            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) => Result.Success(new List<AgSesion>
                        {
                            new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 4.0 }
                        }));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);
            var individuo = new AgIndividuo { Id = 1 };

            var gen = new AgGen
            {
                Id = 1,
                IdSede = 1,
                IdAsignatura = 10,
                IdTercero = 5,
                IdAula = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasSemanales = 4.0,
                TotalHorasAsignatura = 4.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2026, 3, 1),
                FechaFinEstimada = new DateTime(2026, 4, 15)
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto { IdSede = 1 };
            contexto.Ventana = VentanaSemestre.Crear(200, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.BloquesPorId[1] = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 1,
                DuracionHoras = 4.0,
                HoraInicio = "08:00",
                HoraFin = "12:00",
                Franja = "MANANA"
            };

            // Act
            var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            savedVincs.Should().NotBeNull();
            savedVincs.Should().HaveCount(1);
            savedVincs[0].FechaInicio.Should().Be("2026-03-01");
            savedVincs[0].FechaFin.Should().Be("2026-04-15");
            savedVincs[0].IdPeriodo.Should().Be(200);
        }

        [Fact]
        public async Task Materializar_FechaNoParseable_RegistraAdvertenciaYAcota()
        {
            // Arrange (Fase 1115): Gen con fechas fuera de la ventana (2026-01-10 anterior a 2026-02-01 y 2026-07-15 posterior a 2026-06-30)
            // bajo cultura alternativa (th-TH); TryParseExact con InvariantCulture debe parsear correctamente y acotar
            // las fechas a la ventana semestral (2026-02-01..2026-06-30).
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockExpansor = new Mock<IAgExpansorSesiones>();

            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Returns(Task.FromResult(Result.Success()));

            List<AgVinculacion> savedVincs = null;
            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(), It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                    .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>((c, ind, v, s, ct) =>
                    {
                        savedVincs = v.ToList();
                    })
                    .Returns(Task.FromResult(Result.Success()));

            mockExpansor.Setup(e => e.Expandir(It.IsAny<AgGen>(), It.IsAny<int>(), It.IsAny<AgContexto>()))
                        .Returns((AgGen g, int vId, AgContexto ctx) => Result.Success(new List<AgSesion>
                        {
                            new AgSesion { IdGen = g.Id, IdVinculacion = vId, Horas = 4.0 }
                        }));

            var sut = new AgMaterializadorSesiones(mockRepo.Object, mockExpansor.Object);
            var individuo = new AgIndividuo { Id = 1 };

            var gen = new AgGen
            {
                Id = 1,
                IdSede = 1,
                IdAsignatura = 10,
                IdTercero = 5,
                IdAula = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasSemanales = 4.0,
                TotalHorasAsignatura = 4.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2026, 1, 10),
                FechaFinEstimada = new DateTime(2026, 7, 15)
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto { IdSede = 1 };
            contexto.Ventana = VentanaSemestre.Crear(200, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.BloquesPorId[1] = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 1,
                DuracionHoras = 4.0,
                HoraInicio = "08:00",
                HoraFin = "12:00",
                Franja = "MANANA"
            };

            // Act
            var prevCulture = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("th-TH");
                var result = await sut.MaterializarEnStagingAsync(1, individuo, contexto, CancellationToken.None);

                // Assert
                result.IsSuccess.Should().BeTrue();
                savedVincs.Should().NotBeNull();
                savedVincs.Should().HaveCount(1);
                savedVincs[0].FechaInicio.Should().Be("2026-02-01");
                savedVincs[0].FechaFin.Should().Be("2026-06-30");
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = prevCulture;
            }
        }
    }
}
