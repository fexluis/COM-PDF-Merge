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
    }
}
