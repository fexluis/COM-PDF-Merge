using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Models;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class PeriodoCrudServiceTests
    {
        private readonly Mock<IAcademicRepository> _repoMock;
        private readonly PeriodoCrudService _service;

        public PeriodoCrudServiceTests()
        {
            _repoMock = new Mock<IAcademicRepository>();
            _service = new PeriodoCrudService(_repoMock.Object);
        }

        [Fact]
        public void Guardar_WhenCodigoPeriodoEsVacio_ShouldReturnFailure()
        {
            // Arrange
            var dto = new PeriodoGridDto
            {
                PeriodoNombre = "   ",
                PeriodoInicio = "2026-01-12",
                PeriodoFin = "2026-06-19"
            };

            // Act
            var result = _service.Guardar(dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("código o nombre del período es obligatorio");
        }

        [Fact]
        public void Guardar_WhenFechaFinAnteriorAFechaInicio_ShouldReturnFailure()
        {
            // Arrange
            var dto = new PeriodoGridDto
            {
                PeriodoNombre = "2026-01",
                PeriodoInicio = "2026-06-19",
                PeriodoFin = "2026-01-12"
            };

            // Act
            var result = _service.Guardar(dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("no puede ser anterior");
        }

        [Fact]
        public void Guardar_WhenNuevoPeriodoValido_ShouldCallInsertAndReturnId()
        {
            // Arrange
            var dto = new PeriodoGridDto
            {
                IdPeriodo = 0,
                PeriodoNombre = "2027-02",
                PeriodoInicio = "2027-08-15",
                PeriodoFin = "2027-12-12",
                Activo = 1
            };

            _repoMock.Setup(r => r.InsertPeriodo(It.IsAny<Periodo>()))
                     .Returns(SqliteResult<int>.Success(4));

            // Act
            var result = _service.Guardar(dto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(4);
            _repoMock.Verify(r => r.InsertPeriodo(It.Is<Periodo>(p => p.PeriodoNombre == "2027-02")), Times.Once);
        }

        [Fact]
        public void Guardar_WhenPeriodoExistenteValido_ShouldCallUpdate()
        {
            // Arrange
            var dto = new PeriodoGridDto
            {
                IdPeriodo = 2,
                PeriodoNombre = "2026-02",
                PeriodoInicio = "2026-08-15",
                PeriodoFin = "2026-12-12",
                Activo = 1
            };

            var existing = new Periodo { IdPeriodo = 2, PeriodoNombre = "2026-02" };
            _repoMock.Setup(r => r.GetPeriodoById(2)).Returns(SqliteResult<Periodo>.Success(existing));
            _repoMock.Setup(r => r.UpdatePeriodo(It.IsAny<Periodo>())).Returns(SqliteResult<int>.Success(1));

            // Act
            var result = _service.Guardar(dto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            _repoMock.Verify(r => r.UpdatePeriodo(It.Is<Periodo>(p => p.IdPeriodo == 2)), Times.Once);
        }

        [Fact]
        public void Eliminar_WhenIdInvalido_ShouldReturnFailure()
        {
            // Act
            var result = _service.Eliminar(0);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("inválido");
        }
    }
}
