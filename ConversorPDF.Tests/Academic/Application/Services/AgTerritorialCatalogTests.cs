using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgTerritorialCatalogTests
    {
        [Fact]
        public void Separar_WhenValorCompuesto_ShouldDevolverPartes()
        {
            AgTerritorialCatalog.Separar("BOLIVAR, VALLE").Should().Equal("BOLIVAR", "VALLE");
        }

        [Fact]
        public void Separar_WhenValorVacioONull_ShouldDevolverVacio()
        {
            AgTerritorialCatalog.Separar(null).Should().BeEmpty();
            AgTerritorialCatalog.Separar("   ").Should().BeEmpty();
            AgTerritorialCatalog.Separar(",,").Should().BeEmpty();
        }

        [Fact]
        public void Construir_WhenCatalogoYValores_ShouldUnirSinDuplicados()
        {
            var catalogo = new[] { "ANTIOQUIA", "BOLIVAR", "VALLE" };
            var terceros = new[] { "BOLIVAR,VALLE", "VALLE" };

            var resultado = AgTerritorialCatalog.Construir(catalogo, terceros);

            resultado.Should().Equal("ANTIOQUIA", "BOLIVAR", "VALLE");
        }

        [Fact]
        public void Construir_WhenValoresConMinusculasYEspacios_ShouldNormalizar()
        {
            var resultado = AgTerritorialCatalog.Construir(null, new[] { " valle ", "bolivar" });

            resultado.Should().Equal("BOLIVAR", "VALLE");
        }

        [Fact]
        public void Construir_WhenSinFuentes_ShouldDevolverVacio()
        {
            AgTerritorialCatalog.Construir(null, null).Should().BeEmpty();
        }
    }
}
