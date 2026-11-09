using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Unit.Academic
{
    public class VinculacionListServiceTests
    {
        private readonly Mock<IVinculacionQueryRepository> _queryRepoMock = new();
        private readonly Mock<ISedeLookupRepository> _sedeRepoMock = new();
        private readonly Mock<ITerceroLookupRepository> _terceroRepoMock = new();
        private readonly Mock<IPeriodoLookupRepository> _periodoRepoMock = new();
        private readonly Mock<IProgramaLookupRepository> _programaRepoMock = new();

        [Fact]
        public async Task GetListaSesionesAsync_FiltroPorFechas_PasaParametrosAlRepo()
        {
            var service = new VinculacionListService(
                _queryRepoMock.Object, _sedeRepoMock.Object, 
                _terceroRepoMock.Object, _periodoRepoMock.Object,
                _programaRepoMock.Object);
            
            var filtros = new VinculacionFilterDto 
            { 
                FechaDesde = new DateTime(2026, 4, 1),
                FechaHasta = new DateTime(2026, 4, 30)
            };

            await service.GetListaSesionesAsync(filtros, CancellationToken.None);

            _queryRepoMock.Verify(r => r.GetSesionesAsync(
                It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>(),
                new DateTime(2026, 4, 1), new DateTime(2026, 4, 30),
                It.IsAny<int>(), It.IsAny<int>(), It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task GetCatalogosFiltroAsync_RetornaTupleCorrecto()
        {
            _sedeRepoMock.Setup(r => r.GetSedesLookupAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<LookupItem>>(new List<LookupItem> { new LookupItem { Id = 1, DisplayName = "Cali" } }));
            _terceroRepoMock.Setup(r => r.GetTercerosLookupAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<LookupItem>>(new List<LookupItem>()));
            _periodoRepoMock.Setup(r => r.GetPeriodosLookupAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<LookupItem>>(new List<LookupItem>()));
            _programaRepoMock.Setup(r => r.GetProgramasLookupAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<LookupItem>>(new List<LookupItem>()));

            var service = new VinculacionListService(
                _queryRepoMock.Object, _sedeRepoMock.Object,
                _terceroRepoMock.Object, _periodoRepoMock.Object,
                _programaRepoMock.Object);

            var result = await service.GetCatalogosFiltroAsync(CancellationToken.None);

            result.IsSuccess.Should().BeTrue();
            result.Value.Sedes.Should().HaveCount(1);
        }

        [Fact]
        public async Task GetCorridasStagingAsync_DelegaEnRepositorioYAnteponeComodin()
        {
            _queryRepoMock
                .Setup(r => r.GetCorridasStagingAsync(It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<CorridaStagingItem>>(
                    new List<CorridaStagingItem>
                    {
                        new CorridaStagingItem { IdCorrida = 44, PeriodoNombre = "2027-01", TotalSesiones = 3 }
                    }));

            var service = new VinculacionListService(
                _queryRepoMock.Object, _sedeRepoMock.Object,
                _terceroRepoMock.Object, _periodoRepoMock.Object,
                _programaRepoMock.Object);

            var res = await service.GetCorridasStagingAsync(null, null, null, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().HaveCount(2, "el comodín '-- Todas las corridas --' se antepone en Application");
            res.Value.First().IdCorrida.Should().Be(0);
            res.Value.Last().Etiqueta.Should().StartWith("#44");

            _queryRepoMock.Verify(
                r => r.GetCorridasStagingAsync(null, null, null, It.IsAny<CancellationToken>()),
                Times.Once);
        }

        [Fact]
        public async Task GetResumenStagingAsync_TraduceFiltrosYDelegaEnRepositorio()
        {
            _queryRepoMock
                .Setup(r => r.GetResumenStagingAsync(
                    It.IsAny<long>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(),
                    It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>(),
                    It.IsAny<DateTime?>(), It.IsAny<DateTime?>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(new AgStagingResumen { IdCorrida = 44, SesionesCorrida = 244 }));

            var service = new VinculacionListService(
                _queryRepoMock.Object, _sedeRepoMock.Object,
                _terceroRepoMock.Object, _periodoRepoMock.Object,
                _programaRepoMock.Object);

            var filtros = new VinculacionFilterDto { IdSede = 1, IdPeriodo = 2, Subunidad = "VALLE" };
            var res = await service.GetResumenStagingAsync(44, filtros, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.SesionesCorrida.Should().Be(244);

            _queryRepoMock.Verify(
                r => r.GetResumenStagingAsync(44, 1, null, 2, null, null, "VALLE", null, false, null, null, It.IsAny<CancellationToken>()),
                Times.Once);
        }

        [Fact]
        public async Task EliminarCorridaStagingAsync_DelegaEnRepositorio()
        {
            _queryRepoMock
                .Setup(r => r.EliminarCorridaStagingAsync(44, It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<(int Sesiones, int Vinculaciones)>((3, 2)));

            var service = new VinculacionListService(
                _queryRepoMock.Object, _sedeRepoMock.Object,
                _terceroRepoMock.Object, _periodoRepoMock.Object,
                _programaRepoMock.Object);

            var res = await service.EliminarCorridaStagingAsync(44, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Sesiones.Should().Be(3);
            res.Value.Vinculaciones.Should().Be(2);

            _queryRepoMock.Verify(
                r => r.EliminarCorridaStagingAsync(44, It.IsAny<CancellationToken>()),
                Times.Once);
        }
    }
}
