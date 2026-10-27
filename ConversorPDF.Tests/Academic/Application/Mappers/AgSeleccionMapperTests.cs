using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class AgSeleccionMapperTests
    {
        [Fact]
        public void DesdeNodo_WhenDatosCompletosYValidos_ShouldReturnSuccessDtoWithExactDimensions()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "PRESENCIAL")
            {
                Jornada = "NOCHE",
                HorasDia = 4,
                DiasVentana = "MAR-VIE"
            };

            // Act
            var result = AgSeleccionMapper.DesdeNodo(key, 1, 100, 50, "II", "1");

            // Assert
            result.IsSuccess.Should().BeTrue();
            var dto = result.Value;
            dto.IdPeriodo.Should().Be(1);
            dto.IdSede.Should().Be(100);
            dto.IdPrograma.Should().Be(50);
            dto.Semestre.Should().Be("II");
            dto.Grupo.Should().Be("1");
            dto.Modalidad.Should().Be("Presencial");
            dto.Jornada.Should().Be("NOCHE");
            dto.HorasDia.Should().Be(4);
            dto.DiasVentana.Should().Be("MAR-VIE");
        }

        [Theory]
        [InlineData(2)]
        [InlineData(6)]
        [InlineData(10)]
        public void DesdeNodo_WhenHorasDiaInvalido_ShouldFailWithI6(int horasDiaInvalido)
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "PRESENCIAL")
            {
                Jornada = "NOCHE",
                HorasDia = horasDiaInvalido,
                DiasVentana = "MAR-VIE"
            };

            // Act
            var result = AgSeleccionMapper.DesdeNodo(key, 1, 100, 50, "II", "1");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I6_HORAS_DIA_INVALIDO");
        }

        [Fact]
        public void DesdeNodo_WhenVentanaVacia_ShouldFailWithI1()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "PRESENCIAL")
            {
                Jornada = "NOCHE",
                HorasDia = 4,
                DiasVentana = ""
            };

            // Act
            var result = AgSeleccionMapper.DesdeNodo(key, 1, 100, 50, "II", "1");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_VENTANA_VACIA");
        }

        [Fact]
        public void DesdeNodo_WhenJornadaVacia_ShouldFailWithI1()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "PRESENCIAL")
            {
                Jornada = "",
                HorasDia = 4,
                DiasVentana = "MAR-VIE"
            };

            // Act
            var result = AgSeleccionMapper.DesdeNodo(key, 1, 100, 50, "II", "1");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_JORNADA_VACIA");
        }

        [Fact]
        public void DesdeNodo_WhenKeyNula_ShouldFail()
        {
            // Act
            var result = AgSeleccionMapper.DesdeNodo(null, 1, 100, 50, "II", "1");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_KEY_NULA");
        }
    }
}
