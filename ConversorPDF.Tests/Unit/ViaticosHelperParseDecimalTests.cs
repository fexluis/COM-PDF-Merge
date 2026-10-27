using System;
using Xunit;
using FluentAssertions;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Unit
{
    public class ViaticosHelperParseDecimalTests
    {
        [Theory]
        [InlineData(150.5, 150.5)]
        [InlineData(50689, 50689)]
        [InlineData(0, 0)]
        [InlineData(100L, 100)]
        public void ParseDecimal_WhenNativeNumericTypesPassed_ShouldReturnExactDecimal(object input, decimal expected)
        {
            // Act
            decimal actual = ViaticosHelper.ParseDecimal(input);

            // Assert
            actual.Should().Be(expected);
        }

        [Theory]
        [InlineData("150.5", 150.5)]
        [InlineData("150,5", 150.5)]
        [InlineData("$ 50.689,00", 50689.00)]
        [InlineData("1.234,56", 1234.56)]
        [InlineData("1,234.56", 1234.56)]
        [InlineData("\u00A0$ 162.634,00\u00A0", 162634.00)]
        [InlineData("50689 COP", 50689)]
        [InlineData("-50.689,00", -50689.00)]
        [InlineData("(50.689,00)", -50689.00)]
        [InlineData("-50.689", -50.689)]
        [InlineData("(50.689)", -50.689)]
        [InlineData("0", 0)]
        [InlineData("0.0", 0)]
        [InlineData("0,00", 0)]
        [InlineData("", 0)]
        [InlineData("   ", 0)]
        [InlineData(null, 0)]
        public void ParseDecimal_WhenFormattedStringsPassed_ShouldParseCorrectlyWithoutAllowThousandsMultiplication(string input, decimal expected)
        {
            // Act
            decimal actual = ViaticosHelper.ParseDecimal(input);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public void TryParseDecimal_WhenInputIsBlankOrNull_ShouldReturnBlankFlagTrue()
        {
            // Act
            bool parsedNull = ViaticosHelper.TryParseDecimal(null, out decimal valNull, out bool blankNull);
            bool parsedEmpty = ViaticosHelper.TryParseDecimal("   ", out decimal valEmpty, out bool blankEmpty);
            bool parsedZero = ViaticosHelper.TryParseDecimal("0", out decimal valZero, out bool blankZero);
            bool parsedNumber = ViaticosHelper.TryParseDecimal("50689", out decimal valNum, out bool blankNum);

            // Assert
            parsedNull.Should().BeFalse();
            blankNull.Should().BeTrue();
            valNull.Should().Be(0m);

            parsedEmpty.Should().BeFalse();
            blankEmpty.Should().BeTrue();
            valEmpty.Should().Be(0m);

            parsedZero.Should().BeTrue();
            blankZero.Should().BeFalse();
            valZero.Should().Be(0m);

            parsedNumber.Should().BeTrue();
            blankNum.Should().BeFalse();
            valNum.Should().Be(50689m);
        }

        [Fact]
        public void TryParseDecimal_WhenDoubleIsNaNOrInfinity_ShouldReturnFalseWithoutException()
        {
            // Act
            bool parsedNaN = ViaticosHelper.TryParseDecimal(double.NaN, out decimal valNaN, out bool blankNaN);
            bool parsedInf = ViaticosHelper.TryParseDecimal(double.PositiveInfinity, out decimal valInf, out bool blankInf);

            // Assert
            parsedNaN.Should().BeFalse();
            blankNaN.Should().BeTrue();
            valNaN.Should().Be(0m);

            parsedInf.Should().BeFalse();
            blankInf.Should().BeTrue();
            valInf.Should().Be(0m);
        }
    }
}
