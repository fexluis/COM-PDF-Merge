using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Unit.Academic
{
    public class SesionListItemTests
    {
        [Theory]
        [InlineData("PREGRADO", "APT", "APT (PREGRADO)")]
        [InlineData("POSGRADO", "MBA", "MBA (POSGRADO)")]
        public void ProgramaDetalle_FormatoCorrecto(string tipo, string abrev, string esperado)
        {
            var item = new SesionListItem { ProgramaTipo = tipo, ProgramaAbreviatura = abrev };
            item.ProgramaDetalle.Should().Be(esperado);
        }

        [Fact]
        public void RangoFechas_FechaValida_FormatoDDMMYYYY()
        {
            var item = new SesionListItem 
            { 
                FechaSesion = new DateTime(2026, 4, 23)
            };
            item.RangoFechas.Should().Be("23/04/2026");
        }

        [Fact]
        public void RangoFechas_FechasNulas_RetornaNoDefinida()
        {
            var item = new SesionListItem();
            item.RangoFechas.Should().Be("No definida");
        }
    }
}
