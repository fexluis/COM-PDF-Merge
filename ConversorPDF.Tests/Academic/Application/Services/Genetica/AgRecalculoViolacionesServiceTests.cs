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
            Aulas = new List<AgAula>(),
            Bloques = new List<AgBloqueHorario>(),
            Asignaturas = new List<Asignatura>(),
            Docentes = new List<Tercero>(),
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
            Genes = new List<AgGen> { new AgGen { Id = 1, IdAsignatura = 10, IdTercero = 20, IdAula = 30 } },
            Violaciones = new List<AgViolacion>()
        };

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
    }
}