using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgMensajeDimensionesAulaTests
    {
        [Fact]
        public void Faltantes_WhenAulaNullOrAllEmpty_ShouldReturnAllThreeDimensions()
        {
            // Arrange
            var aula = new AgAula
            {
                AulaRef = "3|1|1|III|A|PRESENCIAL",
                Jornada = "",
                HorasDia = 0,
                DiasVentana = ""
            };

            // Act
            var faltantes = AgMensajeDimensionesAula.Faltantes(aula);

            // Assert
            faltantes.Should().BeEquivalentTo(new[] { "días de ventana", "jornada", "horas/día" });
        }

        [Fact]
        public void Faltantes_WhenAulaHasOnlyJornada_ShouldReturnDiasVentanaAndHorasDia()
        {
            // Arrange
            var aula = new AgAula
            {
                AulaRef = "3|1|1|III|A|PRESENCIAL",
                Jornada = "NOCHE",
                HorasDia = 0,
                DiasVentana = ""
            };

            // Act
            var faltantes = AgMensajeDimensionesAula.Faltantes(aula);

            // Assert
            faltantes.Should().BeEquivalentTo(new[] { "días de ventana", "horas/día" });
        }

        [Fact]
        public void Faltantes_WhenAulaHasAllDimensions_ShouldReturnEmpty()
        {
            // Arrange
            var aula = new AgAula
            {
                AulaRef = "3|1|1|III|A|PRESENCIAL",
                Jornada = "NOCHE",
                HorasDia = 4,
                DiasVentana = "VIE-SAB"
            };

            // Act
            var faltantes = AgMensajeDimensionesAula.Faltantes(aula);

            // Assert
            faltantes.Should().BeEmpty();
        }

        [Fact]
        public void Construir_WhenDimensionsMissing_ShouldReturnAccionableTextWithGenerarAulasAndAulaRef()
        {
            // Arrange
            var aula = new AgAula
            {
                AulaRef = "3|1|1|III|A|PRESENCIAL",
                Jornada = "",
                HorasDia = 0,
                DiasVentana = ""
            };

            // Act
            string mensaje = AgMensajeDimensionesAula.Construir(aula, "Error genérico");

            // Assert
            mensaje.Should().Contain("3|1|1|III|A|PRESENCIAL");
            mensaje.Should().Contain("Generar Aulas");
            mensaje.Should().Contain("días de ventana, jornada, horas/día");
        }

        [Fact]
        public void Construir_WhenDimensionsComplete_ShouldReturnValidatorError()
        {
            // Arrange
            var aula = new AgAula
            {
                AulaRef = "3|1|1|III|A|PRESENCIAL",
                Jornada = "NOCHE",
                HorasDia = 4,
                DiasVentana = "VIE-SAB"
            };

            // Act
            string mensaje = AgMensajeDimensionesAula.Construir(aula, "Error específico de catálogo");

            // Assert
            mensaje.Should().Be("Aula '3|1|1|III|A|PRESENCIAL': Error específico de catálogo");
        }
    }
}
