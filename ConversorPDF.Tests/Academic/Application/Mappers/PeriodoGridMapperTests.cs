using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class PeriodoGridMapperTests
    {
        [Fact]
        public void MapToDto_WhenValidEntity_ShouldReturnMatchingDto()
        {
            // Arrange
            var entity = new Periodo
            {
                IdPeriodo = 2,
                PeriodoNombre = "2026-02",
                PeriodoInicio = "2026-08-15",
                PeriodoFin = "2026-12-12",
                Activo = true
            };

            // Act
            var dto = PeriodoGridMapper.MapToDto(entity);

            // Assert
            dto.Should().NotBeNull();
            dto.IdPeriodo.Should().Be(2);
            dto.PeriodoNombre.Should().Be("2026-02");
            dto.PeriodoInicio.Should().Be("2026-08-15");
            dto.PeriodoFin.Should().Be("2026-12-12");
            dto.Activo.Should().Be(1);
            dto.EstadoLegible.Should().Be("Activo");
        }

        [Fact]
        public void MapToDto_WhenNullEntity_ShouldReturnNull()
        {
            // Act
            var dto = PeriodoGridMapper.MapToDto(null);

            // Assert
            dto.Should().BeNull();
        }

        [Fact]
        public void MapToEntity_WhenValidDto_ShouldReturnMatchingEntity()
        {
            // Arrange
            var dto = new PeriodoGridDto
            {
                IdPeriodo = 1,
                PeriodoNombre = "2026-01",
                PeriodoInicio = "2026-01-12",
                PeriodoFin = "2026-06-19",
                Activo = 0
            };

            // Act
            var entity = PeriodoGridMapper.MapToEntity(dto);

            // Assert
            entity.Should().NotBeNull();
            entity.IdPeriodo.Should().Be(1);
            entity.PeriodoNombre.Should().Be("2026-01");
            entity.PeriodoInicio.Should().Be("2026-01-12");
            entity.PeriodoFin.Should().Be("2026-06-19");
            entity.Activo.Should().BeFalse();
        }

        [Fact]
        public void MapToDtoList_WhenEntitiesProvided_ShouldReturnEquivalentDtoList()
        {
            // Arrange
            var entities = new List<Periodo>
            {
                new Periodo { IdPeriodo = 1, PeriodoNombre = "2026-01", Activo = false },
                new Periodo { IdPeriodo = 2, PeriodoNombre = "2026-02", Activo = true }
            };

            // Act
            var dtos = PeriodoGridMapper.MapToDtoList(entities).ToList();

            // Assert
            dtos.Should().HaveCount(2);
            dtos[0].PeriodoNombre.Should().Be("2026-01");
            dtos[1].PeriodoNombre.Should().Be("2026-02");
        }
    }
}
