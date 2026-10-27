using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgRanurasTests
    {
        [Fact]
        public void PorVentana_WhenJueSab4h_ShouldProduceFourOrderedSlotsAndTwoLanes()
        {
            // Arrange
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 1, DiaSemanaNumero = 4, Franja = "NOCHE", DuracionHoras = 4.0, Activo = 1 },
                new AgBloqueHorario { Id = 2, DiaSemanaNumero = 5, Franja = "NOCHE", DuracionHoras = 4.0, Activo = 1 },
                new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, Franja = "MANANA", DuracionHoras = 4.0, Activo = 1 },
                new AgBloqueHorario { Id = 4, DiaSemanaNumero = 6, Franja = "TARDE", DuracionHoras = 4.0, Activo = 1 }
            };

            // Act
            var ranuras = AgRanuras.PorVentana(new[] { 4, 5, 6 }, bloques, "NOCHE", 4.0);
            int carriles = AgCarriles.Cantidad(ranuras);
            var carril0 = AgCarriles.DelCarril(ranuras, 0);
            var carril1 = AgCarriles.DelCarril(ranuras, 1);

            // Assert
            ranuras.Select(r => $"{r.Dia}:{r.Franja}")
                   .Should().Equal("4:NOCHE", "5:NOCHE", "6:MANANA", "6:TARDE");
            carriles.Should().Be(2);
            carril0.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("4:NOCHE", "5:NOCHE");
            carril1.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("6:MANANA", "6:TARDE");
        }

        [Fact]
        public void PorVentana_WhenVieDom4h_ShouldProduceLanesVieSabM_and_SabTDom()
        {
            // Arrange & Act
            var ranuras = AgRanuras.PorVentana(new[] { 5, 6, 7 }, null, "NOCHE", 4.0);
            int carriles = AgCarriles.Cantidad(ranuras);
            var carril0 = AgCarriles.DelCarril(ranuras, 0);
            var carril1 = AgCarriles.DelCarril(ranuras, 1);

            // Assert
            ranuras.Select(r => $"{r.Dia}:{r.Franja}")
                   .Should().Equal("5:NOCHE", "6:MANANA", "6:TARDE", "7:MANANA");
            carriles.Should().Be(2);
            carril0.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("5:NOCHE", "6:MANANA");
            carril1.Select(r => $"{r.Dia}:{r.Franja}").Should().Equal("6:TARDE", "7:MANANA");
        }

        [Fact]
        public void PorVentana_WhenSabado8h_ShouldProduceOneSlotInsteadOfTwo()
        {
            // Arrange & Act
            var ranuras = AgRanuras.PorVentana(new[] { 6 }, null, "MANANA", 8.0);

            // Assert
            ranuras.Should().HaveCount(1);
            ranuras[0].Dia.Should().Be(6);
            ranuras[0].Franja.Should().Be(AgFranjaJornada.MANANA);
            ranuras[0].Horas.Should().Be(8.0);
        }

        [Fact]
        public void FranjasDe_WhenSabado4h_ShouldReturnMananaAndTarde()
        {
            // Act
            var franjas = AgRanuras.FranjasDe(6, "NOCHE", 4.0);

            // Assert
            franjas.Should().Equal(AgFranjaJornada.MANANA, AgFranjaJornada.TARDE);
        }

        [Fact]
        public void FranjasDe_WhenSabado8h_ShouldReturnOnlyManana()
        {
            // Act
            var franjas = AgRanuras.FranjasDe(6, "NOCHE", 8.0);

            // Assert
            franjas.Should().Equal(AgRanuras.FranjasDe(6, "NOCHE", 8.0));
            franjas.Should().Equal(AgFranjaJornada.MANANA);
        }

        [Fact]
        public void PorVentana_WhenJornadaNullSinBloques_ShouldNotDropWeekdays()
        {
            var ranuras = AgRanuras.PorVentana(new[] { 4, 5, 6 }, null, null, 4.0);
            ranuras.Select(r => $"{r.Dia}:{r.Franja}")
                   .Should().Equal("4:NOCHE", "5:NOCHE", "6:MANANA", "6:TARDE");
            AgCarriles.Cantidad(ranuras).Should().Be(2);   // F1171
        }

        [Fact]
        public void PorVentana_WhenVieDomConCatalogoSinDomingo_ShouldKeepSundaySlot()
        {
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 1, DiaSemanaNumero = 5, Franja = "NOCHE", DuracionHoras = 4, Activo = 1 },
                new AgBloqueHorario { Id = 2, DiaSemanaNumero = 6, Franja = "MANANA", DuracionHoras = 4, Activo = 1 },
                new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, Franja = "TARDE", DuracionHoras = 4, Activo = 1 }
            };

            var ranuras = AgRanuras.PorVentana(new[] { 5, 6, 7 }, bloques, "NOCHE", 4.0);

            ranuras.Select(r => $"{r.Dia}:{r.Franja}")
                   .Should().Equal("5:NOCHE", "6:MANANA", "6:TARDE", "7:MANANA");
            AgCarriles.Cantidad(ranuras).Should().Be(2);
        }
    }
}
