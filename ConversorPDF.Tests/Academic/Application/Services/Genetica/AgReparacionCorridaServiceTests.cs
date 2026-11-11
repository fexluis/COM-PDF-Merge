using System.Collections.Generic;
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
    /// <summary>
    /// Pruebas del servicio de reparación de corridas huérfanas del AG (Fase 1566).
    /// Verifican que nunca se escribe con un contexto parcial y que la reconstrucción de cabecera e
    /// individuo se acompaña de la persistencia del detalle de violaciones.
    /// </summary>
    public class AgReparacionCorridaServiceTests
    {
        private readonly Mock<IAgScheduleRepository> _repo = new Mock<IAgScheduleRepository>();
        private readonly Mock<IAgEvaluadorAptitud> _evaluador = new Mock<IAgEvaluadorAptitud>();

        private static AgContexto CrearContexto(int idAula = 31, int idAsignatura = 148)
        {
            var ctx = new AgContexto
            {
                IdPeriodo = 3,
                IdPrograma = 2,
                Subunidad = "VALLE"
            };
            ctx.Aulas.Add(new AgAula { Id = idAula, IdPeriodo = 3, IdPrograma = 2, Subunidad = "VALLE" });
            ctx.Asignaturas.Add(new Asignatura { Id = idAsignatura, IdPrograma = 2 });
            ctx.ConstruirIndices();
            return ctx;
        }

        private static AgCorridaHuerfana Huerfana()
        {
            return new AgCorridaHuerfana
            {
                IdCorrida = 48,
                IdPeriodo = 3,
                IdPrograma = 2,
                Subunidad = "VALLE",
                IdIndividuo = 33,
                TotalVinculaciones = 15,
                TotalSesiones = 180,
                ExisteCabecera = false,
                ExisteIndividuo = false
            };
        }

        private static AgGen Gen(int id, int idAula = 31, int idAsignatura = 148)
        {
            return new AgGen { Id = id, IdAula = idAula, IdAsignatura = idAsignatura };
        }

        private static IReadOnlyCollection<AgGen> Genes(params AgGen[] genes) => genes;

        [Fact]
        public async Task RepararAsync_SinCoberturaDeAulas_NoEscribeNada()
        {
            _repo.Setup(r => r.DetectarCorridaHuerfanaAsync(48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Huerfana()));
            _repo.Setup(r => r.GetGenesByIndividuoAsync(33, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Genes(Gen(1, idAula: 999))));

            var sut = new AgReparacionCorridaService(_repo.Object, _evaluador.Object);
            var res = await sut.RepararAsync(48, CrearContexto(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no cubre la corrida");
            _evaluador.Verify(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()), Times.Never);
            _repo.Verify(r => r.ReinsertarCabeceraHuerfanaAsync(
                It.IsAny<AgCorrida>(), It.IsAny<long>(), It.IsAny<AgIndividuo>(), It.IsAny<CancellationToken>()),
                Times.Never);
        }

        [Fact]
        public async Task RepararAsync_Exito_ReconstruyeCabeceraYPersisteViolaciones()
        {
            _repo.Setup(r => r.DetectarCorridaHuerfanaAsync(48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Huerfana()));
            _repo.Setup(r => r.GetGenesByIndividuoAsync(33, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Genes(Gen(1))));
            _repo.Setup(r => r.ReinsertarCabeceraHuerfanaAsync(
                    It.IsAny<AgCorrida>(), 48, It.IsAny<AgIndividuo>(), It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success());
            _repo.Setup(r => r.ReemplazarViolacionesAsync(
                    33, It.IsAny<IReadOnlyCollection<AgViolacion>>(), 2, 1.5,
                    AgEstadoIndividuo.Infactible, 48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success());

            _evaluador.Setup(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                {
                    ind.Fitness = -100;
                    ind.HardViolations = 2;
                    ind.SoftViolations = 1.5;
                    ind.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R1_SuperposicionDocente, true, "Solape de docente"));
                });

            var sut = new AgReparacionCorridaService(_repo.Object, _evaluador.Object);
            var res = await sut.RepararAsync(48, CrearContexto(), CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.IdCorrida.Should().Be(48);
            res.Value.IdIndividuo.Should().Be(33);
            res.Value.IdPeriodo.Should().Be(3);
            res.Value.IdPrograma.Should().Be(2);
            res.Value.Hard.Should().Be(2);
            res.Value.GenesEvaluados.Should().Be(1);
            res.Value.Violaciones.Should().Be(1);

            _repo.Verify(r => r.ReinsertarCabeceraHuerfanaAsync(
                It.IsAny<AgCorrida>(), 48, It.IsAny<AgIndividuo>(), It.IsAny<CancellationToken>()),
                Times.Once);
            _repo.Verify(r => r.ReemplazarViolacionesAsync(
                33, It.IsAny<IReadOnlyCollection<AgViolacion>>(), 2, 1.5,
                AgEstadoIndividuo.Infactible, 48, It.IsAny<CancellationToken>()),
                Times.Once);
        }

        [Fact]
        public async Task RepararAsync_CuandoLaCorridaConservaCabecera_NoRepara()
        {
            var huerfana = Huerfana();
            huerfana.ExisteCabecera = true;

            _repo.Setup(r => r.DetectarCorridaHuerfanaAsync(48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(huerfana));

            var sut = new AgReparacionCorridaService(_repo.Object, _evaluador.Object);
            var res = await sut.RepararAsync(48, CrearContexto(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("conserva cabecera");
            _repo.Verify(r => r.GetGenesByIndividuoAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [Fact]
        public async Task RepararAsync_SinGenesPersistidos_Falla()
        {
            _repo.Setup(r => r.DetectarCorridaHuerfanaAsync(48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Huerfana()));
            _repo.Setup(r => r.GetGenesByIndividuoAsync(33, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Genes()));

            var sut = new AgReparacionCorridaService(_repo.Object, _evaluador.Object);
            var res = await sut.RepararAsync(48, CrearContexto(), CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no conserva genes");
        }

        [Fact]
        public async Task RepararAsync_ConContextoDeOtroPrograma_FallaSinEscribir()
        {
            // Fase 1577: APT-53 (id 1) y APT-52 (id 2) comparten nombre; el contexto no puede pertenecer a otro id.
            _repo.Setup(r => r.DetectarCorridaHuerfanaAsync(48, It.IsAny<CancellationToken>()))
                 .ReturnsAsync(Result.Success(Huerfana()));

            var contextoOtroPrograma = CrearContexto();
            contextoOtroPrograma.IdPrograma = 1;

            var sut = new AgReparacionCorridaService(_repo.Object, _evaluador.Object);
            var res = await sut.RepararAsync(48, contextoOtroPrograma, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no corresponde a la corrida");
            _repo.Verify(r => r.GetGenesByIndividuoAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()), Times.Never);
            _repo.Verify(r => r.ReinsertarCabeceraHuerfanaAsync(
                It.IsAny<AgCorrida>(), It.IsAny<long>(), It.IsAny<AgIndividuo>(), It.IsAny<CancellationToken>()), Times.Never);
        }

    }
}

