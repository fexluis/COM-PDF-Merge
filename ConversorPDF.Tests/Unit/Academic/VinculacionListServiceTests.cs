using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
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
    }
}
