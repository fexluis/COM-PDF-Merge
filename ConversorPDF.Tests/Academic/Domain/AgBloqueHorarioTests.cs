using System;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgBloqueHorarioTests
    {
        [Theory]
        [InlineData(1, DayOfWeek.Monday)]
        [InlineData(2, DayOfWeek.Tuesday)]
        [InlineData(3, DayOfWeek.Wednesday)]
        [InlineData(4, DayOfWeek.Thursday)]
        [InlineData(5, DayOfWeek.Friday)]
        [InlineData(6, DayOfWeek.Saturday)]
        [InlineData(7, DayOfWeek.Sunday)]
        public void DiaSemanaNumero_WhenValid_ShouldMapCorrectDayOfWeek(int diaNumero, DayOfWeek expectedDay)
        {
            // Arrange & Act
            var bloque = new AgBloqueHorario
            {
                DiaSemanaNumero = diaNumero
            };

            // Assert
            bloque.DiaSemanaNumero.Should().Be(diaNumero);
            bloque.DiaSemana.Should().Be(expectedDay);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(8)]
        [InlineData(-1)]
        public void DiaSemanaNumero_WhenInvalid_ShouldDefaultToMonday(int invalidDiaNumero)
        {
            // Arrange & Act
            var bloque = new AgBloqueHorario
            {
                DiaSemanaNumero = invalidDiaNumero
            };

            // Assert
            bloque.DiaSemana.Should().Be(DayOfWeek.Monday);
        }
    }
}
