using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgFranjaJornadaTests
    {
        [Theory]
        [InlineData("MAÑANA", 4, "08:00", "12:00", "MANANA", 4.0)]
        [InlineData("MANANA", 4, "08:00", "12:00", "MANANA", 4.0)]
        [InlineData("MAÑANA", 8, "08:00", "17:00", "MANANA", 8.0)]
        [InlineData("MANANA", 8, "08:00", "17:00", "MANANA", 8.0)]
        [InlineData("TARDE", 4, "13:00", "17:00", "TARDE", 4.0)]
        [InlineData("TARDE", 8, "14:00", "22:00", "TARDE", 8.0)]
        [InlineData("NOCHE", 4, "18:00", "22:00", "NOCHE", 4.0)]
        [InlineData("NOCHE", 8, "14:00", "22:00", "NOCHE", 8.0)]
        public void Resolve_WhenJornadaAndHorasDiaAreValid_ShouldReturnExpectedFranja(
            string jornada, int horasDia, string expectedIni, string expectedFin, string expectedFranja, double expectedDur)
        {
            // Act
            var result = AgFranjaJornada.Resolve(jornada, horasDia);

            // Assert
            result.Should().NotBeNull();
            result.Value.HoraInicio.Should().Be(expectedIni);
            result.Value.HoraFin.Should().Be(expectedFin);
            result.Value.Franja.Should().Be(expectedFranja);
            result.Value.DuracionHoras.Should().Be(expectedDur);
        }

        [Fact]
        public void Resolve_WhenJornadaIsInvalid_ShouldReturnNull()
        {
            // Act
            var result = AgFranjaJornada.Resolve("OTRA", 4);

            // Assert
            result.Should().BeNull();
        }
    }
}
