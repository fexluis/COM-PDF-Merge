using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class AgVinculacionToVinculacionMapperTests
    {
        [Fact]
        public void MapToProduccion_WhenAgVinculacionProvided_ShouldMapToProduccionVinculacion()
        {
            // Arrange
            var agVinc = new AgVinculacion
            {
                IdTercero = 200,
                IdPrograma = 3,
                Idsede = 4,
                IdAsignatura = 15,
                IdPeriodo = 20261,
                Semestre = "VI",
                Grupo = "A",
                TotalHoras = 48.0,
                DiaHorasSemanales = "LUNES|4, MARTES|4",
                FechaInicio = "2026-02-15",
                FechaFin = "2026-05-30",
                ModalidadVinculacion = "Presencial",
                IsTerminalAereo = 1,
                IdTerminal = 2,
                GastoTerminalAerea = 120000.0,
                Subunidad = "SUB-BOG"
            };

            // Act
            var result = AgVinculacionToVinculacionMapper.MapToProduccion(agVinc);

            // Assert
            result.Should().NotBeNull();
            result.IdTercero.Should().Be(200);
            result.IdPrograma.Should().Be(3);
            result.IdSede.Should().Be(4);
            result.IdAsignatura.Should().Be(15);
            result.IdPeriodo.Should().Be(20261);
            result.Semestre.Should().Be("VI");
            result.Grupo.Should().Be("A");
            result.TotalHoras.Should().Be(48);
            result.DiaHorasSemanales.Should().Be("LUNES|4, MARTES|4");
            result.FechaInicio.Should().Be(new System.DateTime(2026, 2, 15));
            result.FechaFin.Should().Be(new System.DateTime(2026, 5, 30));
            result.Modalidad.Should().Be("Presencial");
            result.IsTerminalAereo.Should().Be(1);
            result.IdTerminales.Should().Be(2);
            result.GastoTerminalAerea.Should().Be(120000.0);
        }

        [Fact]
        public void MapToProduccion_WhenNull_ShouldReturnNull()
        {
            // Act
            var result = AgVinculacionToVinculacionMapper.MapToProduccion(null);

            // Assert
            result.Should().BeNull();
        }
    }
}
