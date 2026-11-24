using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgCarrilesTests
    {
        [Fact]
        public void DelCarril_WhenVentanaParMarVie_ShouldSplitIntoTwoDisjointLanes()
        {
            // Arrange
            var ventana = new[] { 2, 3, 4, 5 };

            // Act
            var carril0 = AgCarriles.DelCarril(ventana, 0);
            var carril1 = AgCarriles.DelCarril(ventana, 1);
            var carril2 = AgCarriles.DelCarril(ventana, 2);
            var carril3 = AgCarriles.DelCarril(ventana, 3);

            // Assert
            carril0.Should().Equal(2, 3);
            carril1.Should().Equal(4, 5);
            carril2.Should().Equal(2, 3);
            carril3.Should().Equal(4, 5);
            AgCarriles.Cantidad(ventana).Should().Be(2);
            AgRanuras.PorVentana(new[] { 4, 5, 6 }, null, "NOCHE", 4.0).Should().HaveCount(4);
        }

        [Fact]
        public void DelCarril_WhenVentanaImparJueSab_ShouldDistributeTwoLanesWithoutDroppingDays()
        {
            // Arrange
            var ventana = new[] { 4, 5, 6 };

            // Act
            var carril0 = AgCarriles.DelCarril(ventana, 0);
            var carril1 = AgCarriles.DelCarril(ventana, 1);

            // Assert
            carril0.Should().Equal(4, 5);
            carril1.Should().Equal(6);
            AgCarriles.Cantidad(ventana).Should().Be(2);
        }

        [Fact]
        public void DelCarril_WhenVentanaImparVieDom_ShouldReturnLane56AndLane7()
        {
            // Arrange
            var ventana = new[] { 5, 6, 7 };

            // Act
            var carril0 = AgCarriles.DelCarril(ventana, 0);
            var carril1 = AgCarriles.DelCarril(ventana, 1);

            // Assert
            carril0.Should().Equal(5, 6);
            carril1.Should().Equal(7);
            AgCarriles.Cantidad(ventana).Should().Be(2);
        }

        [Fact]
        public void DelCarril_WhenRanurasJueSab_ShouldProduceTwoDisjointLanes()
        {
            // Arrange
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, "NOCHE", 4.0),
                new AgRanura(5, "NOCHE", 4.0),
                new AgRanura(6, "MANANA", 4.0),
                new AgRanura(6, "TARDE", 4.0)
            };

            // Act
            int cantidad = AgCarriles.Cantidad(ranuras);
            var carril0 = AgCarriles.DelCarril(ranuras, 0);
            var carril1 = AgCarriles.DelCarril(ranuras, 1);

            // Assert
            cantidad.Should().Be(2);
            carril0.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("4:NOCHE", "5:NOCHE");
            carril1.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("6:MANANA", "6:TARDE");
        }

        [Fact]
        public void Cantidad_WhenVentanaEmptyOrNull_ShouldReturnOne()
        {
            // Act & Assert
            AgCarriles.Cantidad((IReadOnlyList<int>)null).Should().Be(1);
            AgCarriles.Cantidad(new int[0]).Should().Be(1);
        }

        [Theory]
        [InlineData("MAR-VIE", new[] { 2, 3, 4, 5 })]
        [InlineData("2,3,4,5", new[] { 2, 3, 4, 5 })]
        [InlineData("VIE-DOM", new[] { 5, 6, 7 })]
        public void Ventana_WhenStringInput_ShouldParseCorrectly(string raw, int[] expected)
        {
            // Act
            var res = AgCarriles.Ventana(raw);

            // Assert
            res.Should().Equal(expected);
        }

        [Fact]
        public void ParticionSinSingleton_CincoDias_DevuelveDosBandas()
        {
            // Arrange (5 días: 1..5)
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, "NOCHE", 4.0),
                new AgRanura(2, "NOCHE", 4.0),
                new AgRanura(3, "NOCHE", 4.0),
                new AgRanura(4, "NOCHE", 4.0),
                new AgRanura(5, "NOCHE", 4.0)
            };

            // Act: MarVie
            var resMarVie = AgCarriles.ParticionSinSingleton(ranuras, esMarVie: true);
            // Act: LunJue / Default
            var resLunJue = AgCarriles.ParticionSinSingleton(ranuras, esMarVie: false);

            // Assert
            resMarVie.Should().NotBeNull();
            resMarVie.Should().HaveCount(2);
            resMarVie[0].Select(r => r.Dia).Should().Equal(2, 3);
            resMarVie[1].Select(r => r.Dia).Should().Equal(1, 4, 5);

            resLunJue.Should().NotBeNull();
            resLunJue.Should().HaveCount(2);
            resLunJue[0].Select(r => r.Dia).Should().Equal(1, 2, 5);
            resLunJue[1].Select(r => r.Dia).Should().Equal(3, 4);
        }

        [Fact]
        public void ParticionSinSingleton_SieteOTresDias_DevuelveNull()
        {
            // Arrange 3 días
            var ranuras3 = new List<AgRanura>
            {
                new AgRanura(4, "NOCHE", 4.0),
                new AgRanura(5, "NOCHE", 4.0),
                new AgRanura(6, "NOCHE", 4.0)
            };
            // Arrange 7 días
            var ranuras7 = new List<AgRanura>
            {
                new AgRanura(1, "NOCHE", 4.0),
                new AgRanura(2, "NOCHE", 4.0),
                new AgRanura(3, "NOCHE", 4.0),
                new AgRanura(4, "NOCHE", 4.0),
                new AgRanura(5, "NOCHE", 4.0),
                new AgRanura(6, "NOCHE", 4.0),
                new AgRanura(7, "NOCHE", 4.0)
            };

            // Act & Assert
            AgCarriles.ParticionSinSingleton(ranuras3, true).Should().BeNull();
            AgCarriles.ParticionSinSingleton(ranuras7, false).Should().BeNull();
        }

        [Fact]
        public void BandaSinSingletonParaDia_CincoDias_MapeaCorrectamente()
        {
            var ventana = new[] { 1, 2, 3, 4, 5 };

            // MarVie (default o explícito true)
            AgCarriles.BandaSinSingletonParaDia(ventana, 1, esMarVie: true).Should().Equal(1, 4, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 4, esMarVie: true).Should().Equal(1, 4, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 5, esMarVie: true).Should().Equal(1, 4, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 2, esMarVie: true).Should().Equal(2, 3);
            AgCarriles.BandaSinSingletonParaDia(ventana, 3, esMarVie: true).Should().Equal(2, 3);

            // LunJue (explícito false)
            AgCarriles.BandaSinSingletonParaDia(ventana, 1, esMarVie: false).Should().Equal(1, 2, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 2, esMarVie: false).Should().Equal(1, 2, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 5, esMarVie: false).Should().Equal(1, 2, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 3, esMarVie: false).Should().Equal(3, 4);
            AgCarriles.BandaSinSingletonParaDia(ventana, 4, esMarVie: false).Should().Equal(3, 4);

            // Sobrecarga default con heurística residual (Fase 2410, Fase 2415, Fase 2417: 5 -> LunJue {1,2,5}; 1 -> MarVie {1,4,5}; 2/3 -> MarVie {2,3}; 4 -> MarVie {1,4,5})
            AgCarriles.BandaSinSingletonParaDia(ventana, 5).Should().Equal(1, 2, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 1).Should().Equal(1, 4, 5);
            AgCarriles.BandaSinSingletonParaDia(ventana, 2).Should().Equal(2, 3);
            AgCarriles.BandaSinSingletonParaDia(ventana, 3).Should().Equal(2, 3);
            AgCarriles.BandaSinSingletonParaDia(ventana, 4).Should().Equal(1, 4, 5);

            // Ventana no de 5 días
            AgCarriles.BandaSinSingletonParaDia(new[] { 1, 2, 3, 4 }, 1, true).Should().BeNull();

            // Ventana 5 días no canónica {1,2,3,4,6}
            AgCarriles.BandaSinSingletonParaDia(new[] { 1, 2, 3, 4, 6 }, 1, true).Should().BeNull();
        }
    }
}
