using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgResumenFiltrosTests
    {
        [Theory]
        [InlineData(0, 0, "Todos")]
        [InlineData(5, 5, "5")]
        [InlineData(5, 3, "3 (+2 otra modalidad)")]
        [InlineData(5, 0, "0 (+5 otra modalidad)")]
        [InlineData(2, 7, "2")]
        public void Programas_WhenEvaluated_ShouldFormatearResumen(int seleccionados, int visibles, string esperado)
        {
            AgResumenFiltros.Programas(seleccionados, visibles).Should().Be(esperado);
        }
    }
}
