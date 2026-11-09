using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    /// <summary>
    /// Pruebas de la política de filtros por defecto según el modo de datos (Fase 1495).
    /// </summary>
    public class AgFiltrosPolicyTests
    {
        [Fact]
        public void AplicarPeriodoConfig_ModoAg_EsFalso()
        {
            AgFiltrosPolicy.AplicarPeriodoConfig(AgModoDatos.Ag).Should().BeFalse(
                "en MODO AG la corrida es la fuente de verdad y el período de config escondía sesiones");
        }

        [Fact]
        public void AplicarSubunidadConfig_ModoAg_EsFalso()
        {
            AgFiltrosPolicy.AplicarSubunidadConfig(AgModoDatos.Ag).Should().BeFalse();
        }

        [Fact]
        public void AplicarPeriodoConfig_Produccion_EsVerdadero()
        {
            AgFiltrosPolicy.AplicarPeriodoConfig(AgModoDatos.Determinista).Should().BeTrue();
        }

        [Fact]
        public void AplicarSubunidadConfig_Produccion_EsVerdadero()
        {
            AgFiltrosPolicy.AplicarSubunidadConfig(AgModoDatos.Determinista).Should().BeTrue();
        }

        [Fact]
        public void AplicarPeriodoConfig_ModoNulo_AsumeProduccion()
        {
            AgFiltrosPolicy.AplicarPeriodoConfig(null).Should().BeTrue();
            AgFiltrosPolicy.AplicarSubunidadConfig(null).Should().BeTrue();
        }
    }
}
