using ConversorPDF.XLL.Academic.Application.DTOs;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.DTOs
{
    public class AgTerceroSeleccionDtoTests
    {
        [Fact]
        public void TieneAsignaturasHabilitadas_WhenSetTrue_ShouldMarcarHabilitacionLocal()
        {
            var dto = new AgTerceroSeleccionDto { TieneAsignaturasHabilitadas = true };

            dto.HabilitacionesLocales.Should().Be(1);
            dto.HabilitacionesOtras.Should().Be(0);
            dto.TotalHabilitaciones.Should().Be(1);
        }

        [Fact]
        public void TotalHabilitaciones_WhenSet_ShouldAsignarALocales()
        {
            var dto = new AgTerceroSeleccionDto { TotalHabilitaciones = 4 };

            dto.HabilitacionesLocales.Should().Be(4);
            dto.HabilitacionesOtras.Should().Be(0);
        }

        [Fact]
        public void TotalHabilitaciones_WhenHayOtras_ShouldSumarAmbas()
        {
            var dto = new AgTerceroSeleccionDto { HabilitacionesLocales = 2, HabilitacionesOtras = 3 };

            dto.TotalHabilitaciones.Should().Be(5);
            dto.TieneAsignaturasHabilitadas.Should().BeTrue();
        }
    }
}
