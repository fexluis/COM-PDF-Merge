using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgFormatoAsignaturaTests
    {
        [Theory]
        [InlineData("2", "II")]
        [InlineData("ii", "II")]
        [InlineData("VIII", "VIII")]
        [InlineData("", "SIN ASIGNAR")]
        [InlineData(null, "SIN ASIGNAR")]
        [InlineData("   ", "SIN ASIGNAR")]
        [InlineData("XI", "XI")]
        [InlineData("1", "I")]
        [InlineData("3", "III")]
        [InlineData("4", "IV")]
        [InlineData("5", "V")]
        [InlineData("6", "VI")]
        [InlineData("7", "VII")]
        [InlineData("8", "VIII")]
        [InlineData("9", "IX")]
        [InlineData("10", "X")]
        public void SemestreRomano_WhenInputProvided_ShouldReturnExpectedFormat(string input, string expected)
        {
            // Act
            var actual = AgFormatoAsignatura.SemestreRomano(input);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public void Linea_WhenAsignaturaValid_ShouldFormatWithRomanSemesterAndArrow()
        {
            // Arrange
            var asig = new Asignatura
            {
                Id = 132,
                Semestre = "II",
                Nombre = "MATEMATICAS II"
            };

            // Act
            var actual = AgFormatoAsignatura.Linea(asig);

            // Assert
            actual.Should().Be("132: II → MATEMATICAS II");
        }

        [Fact]
        public void Linea_WhenAsignaturaIsNull_ShouldReturnEmptyString()
        {
            // Act
            var actual = AgFormatoAsignatura.Linea(null);

            // Assert
            actual.Should().BeEmpty();
        }

        [Fact]
        public void BulletList_WhenMultipleAsignaturas_ShouldFormatWithBulletSeparators()
        {
            // Arrange
            var lista = new List<Asignatura>
            {
                new Asignatura { Id = 132, Semestre = "II", Nombre = "MATEMATICAS II" },
                new Asignatura { Id = 133, Semestre = "II", Nombre = "PROYECTO DE FUTURO I" }
            };

            // Act
            var actual = AgFormatoAsignatura.BulletList(lista);

            // Assert
            actual.Should().Be("132: II → MATEMATICAS II\n• 133: II → PROYECTO DE FUTURO I");
        }

        [Theory]
        [InlineData("IV", "4")]
        [InlineData("iv", "4")]
        [InlineData("VIII", "8")]
        [InlineData("", "")]
        [InlineData(null, "")]
        [InlineData("   ", "")]
        [InlineData("SIN ASIGNAR", "")]
        [InlineData("2", "2")]
        [InlineData("10", "10")]
        public void SemestreArabigo_WhenInputProvided_ShouldReturnExpectedArabicFormat(string input, string expected)
        {
            // Act
            var actual = AgFormatoAsignatura.SemestreArabigo(input);

            // Assert
            actual.Should().Be(expected);
        }

        [Theory]
        [InlineData("II-A", "II", "A")]
        [InlineData("VIII", "VIII", "")]
        [InlineData("8-b", "VIII", "B")]
        [InlineData("", "", "")]
        [InlineData(null, "", "")]
        [InlineData("4-c", "IV", "C")]
        public void ParsearSemestreGrupo_WhenInputProvided_ShouldReturnTuple(string input, string expectedSem, string expectedGrp)
        {
            // Act
            var actual = AgFormatoAsignatura.ParsearSemestreGrupo(input);

            // Assert
            actual.Semestre.Should().Be(expectedSem);
            actual.Grupo.Should().Be(expectedGrp);
        }

        [Fact]
        public void Flecha_WhenConfiguredToAscii_ShouldUseConfiguredArrow()
        {
            // Arrange
            string flechaOriginal = AgFormatoAsignatura.Flecha;
            try
            {
                AgFormatoAsignatura.Flecha = "->";
                var asig = new Asignatura { Id = 10, Semestre = "2", Nombre = "ALGEBRA" };

                // Act
                var actual = AgFormatoAsignatura.Linea(asig);

                // Assert
                actual.Should().Be("10: II -> ALGEBRA");
            }
            finally
            {
                AgFormatoAsignatura.Flecha = flechaOriginal;
            }
        }
    }
}
