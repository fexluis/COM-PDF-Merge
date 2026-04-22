using Xunit;
using FluentAssertions;
using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Unit
{
    public class DayPatternParserTests
    {
        [Theory]
        [InlineData("LUNES", "Monday")]
        [InlineData("MARTES", "Tuesday")]
        [InlineData("MIERCOLES", "Wednesday")]
        [InlineData("MIÉRCOLES", "Wednesday")]
        [InlineData("JUEVES", "Thursday")]
        [InlineData("VIERNES", "Friday")]
        [InlineData("SABADO", "Saturday")]
        [InlineData("SÁBADO", "Saturday")]
        [InlineData("DOMINGO", "Sunday")]
        public void TranslateDay_TodosLosDias_MapeanCorrectamente(string espanol, string ingles)
        {
            var result = DayPatternParser.TranslateDay(espanol);
            result.Should().Be(ingles);
        }

        [Fact]
        public void Parse_PatronJueves8_RetornaDiccionarioCorrecto()
        {
            var input = new Dictionary<string, int> { { "JUEVES", 8 } };
            var result = DayPatternParser.Parse(input);
            result.Should().ContainKey(DayOfWeek.Thursday).WhoseValue.Should().Be(8);
        }

        [Fact]
        public void Parse_DiaInvalido_IgnoraEntrada()
        {
            var input = new Dictionary<string, int> { { "FERIADO", 8 }, { "LUNES", 4 } };
            var result = DayPatternParser.Parse(input);
            result.Should().HaveCount(1).And.ContainKey(DayOfWeek.Monday);
        }
    }
}
