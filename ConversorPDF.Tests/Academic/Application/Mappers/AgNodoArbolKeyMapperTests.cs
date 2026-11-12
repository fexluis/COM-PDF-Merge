using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class AgNodoArbolKeyMapperTests
    {
        [Fact]
        public void AplicarDimensiones_WhenDimensionesValidas_ShouldAsignarKeyYRetornarSuccess()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial")
            {
                Jornada = "TARDE",
                HorasDia = 8,
                DiasVentana = "LUN-JUE"
            };

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, dto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            key.Jornada.Should().Be("TARDE");
            key.HorasDia.Should().Be(8);
            key.DiasVentana.Should().Be("LUN-JUE");
        }

        [Fact]
        public void AplicarDimensiones_WhenJornadaVacia_ShouldAplicarFallbackMananaYRetornarFailure()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial")
            {
                Jornada = "   ",
                HorasDia = 8,
                DiasVentana = "MAR-VIE"
            };

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_RESTAURACION_INCOMPLETA");
            key.Jornada.Should().Be("MANANA");
            key.HorasDia.Should().Be(8);
            key.DiasVentana.Should().Be("MAR-VIE");
        }

        [Fact]
        public void AplicarDimensiones_WhenDiasVentanaInvalido_ShouldAplicarFallbackMarVieYRetornarFailure()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial")
            {
                Jornada = "TARDE",
                HorasDia = 4,
                DiasVentana = "INVALIDO"
            };

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_RESTAURACION_INCOMPLETA");
            key.Jornada.Should().Be("TARDE");
            key.HorasDia.Should().Be(4);
            key.DiasVentana.Should().Be("MAR-VIE");
        }

        [Theory]
        [InlineData(2)]
        [InlineData(6)]
        [InlineData(10)]
        public void AplicarDimensiones_WhenHorasDiaInvalido_ShouldAplicarFallback4hYRetornarFailure(int horasInvalidas)
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial")
            {
                Jornada = "NOCHE",
                HorasDia = horasInvalidas,
                DiasVentana = "MAR-VIE"
            };

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_RESTAURACION_INCOMPLETA");
            key.Jornada.Should().Be("NOCHE");
            key.HorasDia.Should().Be(4);
            key.DiasVentana.Should().Be("MAR-VIE");
        }

        [Fact]
        public void AplicarDimensiones_WhenKeyNula_ShouldRetornarFailureI1KeyNula()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial");

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(null, dto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_KEY_NULA");
        }

        [Fact]
        public void AplicarDimensiones_WhenDtoNulo_ShouldRetornarFailureI1DtoNulo()
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, null);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_DTO_NULO");
        }

        [Theory]
        [InlineData("LUN-VIE")]
        [InlineData("JUE-DOM")]
        [InlineData("MIE-DOM")]
        public void AplicarDimensiones_WhenVentanasNuevasValidas_ShouldAsignarKeyYRetornarSuccess(string ventana)
        {
            // Arrange
            var key = new AgNodoArbolKey(6, 10, "PRESENCIAL", "Presencial");
            var dto = new AgContextoSeleccionDto(1, 100, 50, "II", "A", "Presencial")
            {
                Jornada = "MANANA",
                HorasDia = 4,
                DiasVentana = ventana
            };

            // Act
            var result = AgNodoArbolKeyMapper.AplicarDimensiones(key, dto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            key.DiasVentana.Should().Be(ventana);
        }
    }
}
