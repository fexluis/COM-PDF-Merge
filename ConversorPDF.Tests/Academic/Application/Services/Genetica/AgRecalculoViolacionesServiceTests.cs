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
    public class AgRecalculoViolacionesServiceTests
    {
        private static AgContexto ContextoVacio() => new AgContexto
        {
            Aulas = new List<AgAula> { new AgAula { Id = 30, Nombre = "AULA 30" } },
            Bloques = new List<AgBloqueHorario>(),
            Asignaturas = new List<Asignatura> { new Asignatura { Id = 10, Nombre = "ASIGNATURA 10" } },
            Docentes = new List<Tercero> { new Tercero { Id = 20, Nombre = "DOCENTE 20" } },
            Sedes = new List<Sede>(),
            Programas = new List<Programa>(),
            PoolDocentes = new List<TerceroAsignatura>(),
            AulaAsignaturas = new List<AgAulaAsignatura>(),
            Festivos = new List<(System.DateTime, System.DateTime)>()
        };

        private static AgIndividuo IndividuoConGen() => new AgIndividuo
        {
            Id = 7,
            IdCorrida = 44,
            Estado = AgEstadoIndividuo.Factible,
            Genes = new List<AgGen> { new AgGen { Id = 1, IdAsignatura = 10, IdTercero = 20, IdAula = 30, HorasGen = 4.0 } },
            Violaciones = new List<AgViolacion>()
        };

        private static void SetupStaging(Mock<IAgScheduleRepository> mockRepo, long idCorrida = 44, long idGen = 1, double horas = 4.0)
        {
            mockRepo.Setup(r => r.GetSesionesStagingAsync(idCorrida, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success((IReadOnlyCollection<AgSesion>)new List<AgSesion>
                    {
                        new AgSesion
                        {
                            IdGen = idGen,
                            Horas = horas,
                            FechaSesion = new System.DateTime(2026, 3, 1)
                        }
                    }));
        }

        [Fact]
        public async Task RecalcularAsync_SinContexto_DevuelveFallo()
        {
            var sut = new AgRecalculoViolacionesService(
                Mock.Of<IAgScheduleRepository>(), Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, null, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("contexto");
        }

        [Fact]
        public async Task RecalcularAsync_CorridaSinIndividuoOptimo_DevuelveFallo()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdIndividuoOptimo = null }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("individuo óptimo");
        }

        [Fact]
        public async Task RecalcularAsync_RepueblaDetalleYPersisteDesglose()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockValidador = new Mock<IAgValidadorReglas>();

            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            SetupStaging(mockRepo);

            IReadOnlyCollection<AgViolacion> persistidas = null;
            int hardPersistido = -1;
            double softPersistido = -1;
            long corridaPersistida = -1;

            mockRepo.Setup(r => r.ReemplazarViolacionesAsync(
                        It.IsAny<long>(), It.IsAny<IReadOnlyCollection<AgViolacion>>(),
                        It.IsAny<int>(), It.IsAny<double>(), It.IsAny<AgEstadoIndividuo>(),
                        It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .Callback<long, IReadOnlyCollection<AgViolacion>, int, double, AgEstadoIndividuo, long, CancellationToken>(
                        (id, viols, hard, soft, estado, corrida, ct) =>
                        {
                            persistidas = viols;
                            hardPersistido = hard;
                            softPersistido = soft;
                            corridaPersistida = corrida;
                        })
                    .ReturnsAsync(Result.Success());

            mockValidador.Setup(v => v.Evaluar(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                         .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                         {
                             ind.Violaciones = new List<AgViolacion>
                             {
                                 new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "choque 1", 1),
                                 new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "choque 2", 1),
                                 new AgViolacion(AgReglaViolacion.R2_SuperposicionAula, false, "aviso aula", 1)
                             };
                             ind.HardViolations = 2;
                             ind.SoftViolations = 1.5;
                         });

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, mockValidador.Object);

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Hard.Should().Be(2);
            res.Value.Soft.Should().Be(1.5);
            res.Value.GenesEvaluados.Should().Be(1);
            res.Value.PorRegla[AgReglaViolacion.R1_SuperposicionDocente].Should().Be(2);
            res.Value.PorRegla[AgReglaViolacion.R2_SuperposicionAula].Should().Be(1);
            res.Value.ResumenPorRegla.Should().Contain("R1_SuperposicionDocente=2");

            persistidas.Should().HaveCount(3);
            hardPersistido.Should().Be(2);
            softPersistido.Should().Be(1.5);
            corridaPersistida.Should().Be(44);
            mockValidador.Verify(v => v.Evaluar(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()), Times.Once);
        }

        [Fact]
        public async Task RecalcularAsync_IndividuoSinGenes_DevuelveFallo()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgIndividuo { Id = 7, Genes = new List<AgGen>() }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("genes");
        }

        [Fact]
        public async Task RecalcularAsync_ContextoNoCorrespondeACorrida_DevuelveFallo()
        {
            // Fase 1532: rechazar si el período o programa del contexto no coincide con la corrida
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 2, IdPrograma = 10, IdIndividuoOptimo = 7 }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio();
            ctx.IdPeriodo = 1; // Discrepante
            ctx.IdPrograma = 10;

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no corresponde a la corrida");
        }

        [Fact]
        public async Task RecalcularAsync_ContextoIncompleto_DevuelveFallo()
        {
            // Fase 1533 & Fase 1540: no lanzar NRE si alguna colección requerida (Aulas, Bloques, Festivos, etc.) es null
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            SetupStaging(mockRepo);

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio();
            ctx.Aulas = null; // Incompleto

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("incompleto");

            // Fase 1540: Festivos es null
            var ctxFestivosNull = ContextoVacio();
            ctxFestivosNull.Festivos = null;
            var resFestivos = await sut.RecalcularAsync(44, ctxFestivosNull, CancellationToken.None);
            resFestivos.IsFailure.Should().BeTrue();
            resFestivos.Error.Should().Contain("incompleto");
        }

        [Fact]
        public async Task RecalcularAsync_SubunidadUnilateral_NoBloqueaRecalculo()
        {
            // Fase 1538: si el contexto no tiene subunidad pero la corrida sí, se tolera
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, Subunidad = "VALLE", IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            SetupStaging(mockRepo);
            mockRepo.Setup(r => r.ReemplazarViolacionesAsync(
                        It.IsAny<long>(), It.IsAny<IReadOnlyCollection<AgViolacion>>(),
                        It.IsAny<int>(), It.IsAny<double>(), It.IsAny<AgEstadoIndividuo>(),
                        It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success());

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio();
            ctx.Subunidad = string.Empty; // Contexto sin subunidad

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public async Task RecalcularAsync_PreservaIdCorridaPrevioDelContexto()
        {
            // Fase 1534: restaurar IdCorrida del contexto en el bloque finally
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockValidador = new Mock<IAgValidadorReglas>();

            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            SetupStaging(mockRepo);
            mockRepo.Setup(r => r.ReemplazarViolacionesAsync(
                        It.IsAny<long>(), It.IsAny<IReadOnlyCollection<AgViolacion>>(),
                        It.IsAny<int>(), It.IsAny<double>(), It.IsAny<AgEstadoIndividuo>(),
                        It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success());

            long idCorridaDuranteEvaluacion = -1;
            mockValidador.Setup(v => v.Evaluar(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                         .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                         {
                             idCorridaDuranteEvaluacion = ctx.IdCorrida;
                         });

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, mockValidador.Object);
            var contexto = ContextoVacio();
            contexto.IdCorrida = 999;

            var res = await sut.RecalcularAsync(44, contexto, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            idCorridaDuranteEvaluacion.Should().Be(44);
            contexto.IdCorrida.Should().Be(999);
        }

        [Fact]
        public async Task RecalcularAsync_ContextoNoCubreAulasDeLaCorrida_DevuelveFallo()
        {
            // Fase 1549: si el contexto no contiene las aulas de los genes de la corrida, rechazar
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgIndividuo
                    {
                        Id = 7,
                        IdCorrida = 44,
                        Genes = new List<AgGen> { new AgGen { Id = 1, IdAula = 99 } }
                    }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio(); // Contiene aula 30, pero no aula 99

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no cubre");
        }

        [Fact]
        public async Task RecalcularAsync_ContextoSinAsignaturasDeLaCorrida_DevuelveFallo()
        {
            // Fase 1551: aula cubierta, asignatura ausente -> rechazar
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgIndividuo
                    {
                        Id = 7,
                        IdCorrida = 44,
                        Genes = new List<AgGen> { new AgGen { Id = 1, IdAula = 30, IdAsignatura = 77, IdTercero = 0 } }
                    }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio();   // aula 30 presente; Asignatura 77 ausente (contiene 10)

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("asignatura");
        }

        [Fact]
        public async Task RecalcularAsync_DocenteAusenteDelContexto_NoBloqueaRecalculo()
        {
            // Fase 1552: los docentes ausentes no condicionan las violaciones y no deben bloquear el recálculo
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgIndividuo
                    {
                        Id = 7,
                        IdCorrida = 44,
                        Genes = new List<AgGen> { new AgGen { Id = 1, IdAula = 30, IdAsignatura = 10, IdTercero = 999, HorasGen = 4.0 } }
                    }));
            SetupStaging(mockRepo);
            mockRepo.Setup(r => r.ReemplazarViolacionesAsync(
                        It.IsAny<long>(), It.IsAny<IReadOnlyCollection<AgViolacion>>(),
                        It.IsAny<int>(), It.IsAny<double>(), It.IsAny<AgEstadoIndividuo>(),
                        It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success());

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());
            var ctx = ContextoVacio();   // aula 30 presente, asignatura 10 presente, docente 999 ausente

            var res = await sut.RecalcularAsync(44, ctx, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public async Task RecalcularAsync_StagingConHorasDiscrepantes_DevuelveFallo()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            // Staging reporta 2.0h cuando gen espera 4.0h
            SetupStaging(mockRepo, 44, 1, 2.0);

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no coinciden");
        }

        [Fact]
        public async Task RecalcularAsync_StagingConSesionesHuerfanas_DevuelveFallo()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(IndividuoConGen()));
            mockRepo.Setup(r => r.GetSesionesStagingAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success((IReadOnlyCollection<AgSesion>)new List<AgSesion>
                    {
                        new AgSesion { IdGen = 1, Horas = 4.0, FechaSesion = new System.DateTime(2026, 3, 1) },
                        new AgSesion { IdGen = 999, Horas = 4.0, FechaSesion = new System.DateTime(2026, 3, 1) } // Gen desconocido
                    }));

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("desconocidos");
        }

        [Fact]
        public async Task RecalcularAsync_GenConCeroHorasYSinSesiones_RecalculaExitosamente()
        {
            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.GetCorridaByIdAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 44, IdPeriodo = 0, IdPrograma = 0, IdIndividuoOptimo = 7 }));
            mockRepo.Setup(r => r.GetIndividuoByIdAsync(7, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgIndividuo
                    {
                        Id = 7,
                        IdCorrida = 44,
                        Genes = new List<AgGen> { new AgGen { Id = 1, IdAula = 30, IdAsignatura = 10, IdTercero = 20, HorasGen = 0.0 } }
                    }));
            mockRepo.Setup(r => r.GetSesionesStagingAsync(44, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success((IReadOnlyCollection<AgSesion>)new List<AgSesion>()));
            mockRepo.Setup(r => r.ReemplazarViolacionesAsync(
                        It.IsAny<long>(), It.IsAny<IReadOnlyCollection<AgViolacion>>(),
                        It.IsAny<int>(), It.IsAny<double>(), It.IsAny<AgEstadoIndividuo>(),
                        It.IsAny<long>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success());

            var sut = new AgRecalculoViolacionesService(mockRepo.Object, Mock.Of<IAgValidadorReglas>());

            var res = await sut.RecalcularAsync(44, ContextoVacio(), CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
        }
    }
}