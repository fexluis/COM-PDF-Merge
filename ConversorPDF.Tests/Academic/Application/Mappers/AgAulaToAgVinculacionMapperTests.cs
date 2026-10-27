using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class AgAulaToAgVinculacionMapperTests
    {
        [Fact]
        public void AplicarAtributosAula_WhenValuesMissingInVinculacion_ShouldPopulateFromAula()
        {
            // Arrange
            var vinc = new AgVinculacion
            {
                IdPeriodo = 0,
                IdPrograma = 0,
                Semestre = "",
                Grupo = "",
                FechaInicio = "",
                FechaFin = ""
            };

            var aula = new AgAula
            {
                IdPeriodo = 20261,
                IdPrograma = 10,
                Semestre = "IV",
                Grupo = "C",
                FechaInicio = "2026-02-01",
                FechaFin = "2026-06-30",
                ModalidadVinculacion = "Virtual"
            };

            // Act
            AgAulaToAgVinculacionMapper.AplicarAtributosAula(vinc, aula);

            // Assert
            vinc.IdPeriodo.Should().Be(20261);
            vinc.IdPrograma.Should().Be(10);
            vinc.Semestre.Should().Be("IV");
            vinc.Grupo.Should().Be("C");
            vinc.FechaInicio.Should().Be("2026-02-01");
            vinc.ModalidadVinculacion.Should().Be("Plataforma");
        }

        [Fact]
        public void AplicarAtributosAula_WhenSemipresencialInAula_ShouldMapToMixtaInVinculacion()
        {
            // Arrange
            var vinc = new AgVinculacion();
            var aula = new AgAula
            {
                ModalidadVinculacion = "Semipresencial"
            };

            // Act
            AgAulaToAgVinculacionMapper.AplicarAtributosAula(vinc, aula);

            // Assert
            vinc.ModalidadVinculacion.Should().Be("Mixta");
        }
    }
}
