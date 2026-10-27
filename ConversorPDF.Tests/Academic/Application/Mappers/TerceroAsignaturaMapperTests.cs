using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class TerceroAsignaturaMapperTests
    {
        [Fact]
        public void ToDto_WhenValidEntity_ShouldMapAllFields()
        {
            // Arrange
            var entity = new TerceroAsignatura
            {
                Id = 123456789L,
                IdTercero = 10,
                IdPrograma = 2,
                IdAsignatura = 105,
                Semestre = "IV",
                Activo = 1,
                Prioridad = 5,
                VecesAsignada = 3,
                FechaUltimaAsignacion = "2026-05-21",
                Origen = "MANUAL",
                Observaciones = "Especialista en la materia",
                Subunidad = "SEDE_CENTRAL"
            };

            // Act
            var dto = TerceroAsignaturaMapper.ToDto(entity, "Juan Perez", "Administracion", "Economia");

            // Assert
            dto.Should().NotBeNull();
            dto.Id.Should().Be(123456789L);
            dto.IdTercero.Should().Be(10);
            dto.NombreTercero.Should().Be("Juan Perez");
            dto.IdPrograma.Should().Be(2);
            dto.NombrePrograma.Should().Be("Administracion");
            dto.IdAsignatura.Should().Be(105);
            dto.NombreAsignatura.Should().Be("Economia");
            dto.Semestre.Should().Be("IV");
            dto.Activo.Should().Be(1);
            dto.Prioridad.Should().Be(5);
            dto.VecesAsignada.Should().Be(3);
            dto.FechaUltimaAsignacion.Should().Be("2026-05-21");
            dto.Origen.Should().Be("MANUAL");
            dto.Observaciones.Should().Be("Especialista en la materia");
            dto.Subunidad.Should().Be("SEDE_CENTRAL");
        }

        [Fact]
        public void ToEntity_WhenDtoHasNoId_ShouldCalculateDeterministicId()
        {
            // Arrange
            var dto = new TerceroAsignaturaDto
            {
                Id = 0,
                IdTercero = 10,
                IdPrograma = 2,
                IdAsignatura = 105,
                Semestre = "IV",
                Activo = 1,
                Prioridad = 2,
                VecesAsignada = 1,
                FechaUltimaAsignacion = "2026-05-21",
                Origen = "MANUAL"
            };

            // Act
            var entity = TerceroAsignaturaMapper.ToEntity(dto);

            // Assert
            entity.Should().NotBeNull();
            long expectedId = TerceroAsignaturaKey.Calcular(10, 2, 105, "IV");
            entity.Id.Should().Be(expectedId);
            entity.IdTercero.Should().Be(10);
            entity.IdPrograma.Should().Be(2);
            entity.IdAsignatura.Should().Be(105);
            entity.Semestre.Should().Be("IV");
        }

        [Fact]
        public void FromVinculacion_WhenValidVinculacion_ShouldMapToEntityWithProvidedVecesAsignada()
        {
            // Arrange
            var vinc = new Vinculacion
            {
                Id = 1,
                IdTercero = 10,
                IdPrograma = 2,
                IdAsignatura = 105,
                Semestre = "IV",
                Subunidad = "SEDE_CENTRAL",
                FechaInicio = new System.DateTime(2026, 3, 1)
            };

            // Act
            var entity = TerceroAsignaturaMapper.FromVinculacion(vinc, vecesAsignada: 5);

            // Assert
            entity.Should().NotBeNull();
            entity.Id.Should().Be(TerceroAsignaturaKey.Calcular(10, 2, 105, "IV"));
            entity.IdTercero.Should().Be(10);
            entity.IdPrograma.Should().Be(2);
            entity.IdAsignatura.Should().Be(105);
            entity.Semestre.Should().Be("IV");
            entity.VecesAsignada.Should().Be(5);
            entity.FechaUltimaAsignacion.Should().Be("2026-03-01");
            entity.Origen.Should().Be("VINCULACION");
            entity.Subunidad.Should().Be("SEDE_CENTRAL");
        }
    }
}
