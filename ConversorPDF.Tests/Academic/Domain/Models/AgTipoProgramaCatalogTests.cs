using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain.Models
{
    public class AgTipoProgramaCatalogTests
    {
        [Theory]
        [InlineData("TODOS", "PREGRADO", true)]
        [InlineData("TODOS", "POSGRADO", true)]
        [InlineData("TODOS", null, true)]
        [InlineData("PREGRADO", "PREGRADO", true)]
        [InlineData("PREGRADO", "Pregrado", true)]
        [InlineData("PREGRADO", "POSGRADO", false)]
        [InlineData("POSGRADO", "POSGRADO", true)]
        [InlineData("POSGRADO", "Posgrado", true)]
        [InlineData("POSGRADO", "PREGRADO", false)]
        public void Coincide_WhenEvaluated_ShouldMatchExpectedResult(string filtro, string programaModalidad, bool expected)
        {
            // Act
            bool result = AgTipoProgramaCatalog.Coincide(filtro, programaModalidad);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("POSGRADO", 8)]
        [InlineData("posgrado", 8)]
        [InlineData("PREGRADO", 4)]
        [InlineData("pregrado", 4)]
        [InlineData(null, 4)]
        [InlineData("", 4)]
        public void HorasDiaSugeridas_WhenEvaluated_ShouldReturnExpectedHours(string modalidad, int expectedHoras)
        {
            // Act
            int horas = AgTipoProgramaCatalog.HorasDiaSugeridas(modalidad);

            // Assert
            horas.Should().Be(expectedHoras);
        }

        [Fact]
        public void Filtrar_WhenFiltroEsPregrado_ShouldReturnOnlyPregradoProgramas()
        {
            // Arrange
            var programas = new List<Programa>
            {
                new Programa { Id = 1, NombrePrograma = "Prog 1", Modalidad = "PREGRADO" },
                new Programa { Id = 2, NombrePrograma = "Prog 2", Modalidad = "POSGRADO" },
                new Programa { Id = 3, NombrePrograma = "Prog 3", Modalidad = "PREGRADO" }
            };

            // Act
            var resultado = AgTipoProgramaCatalog.Filtrar(programas, AgTipoProgramaCatalog.Pregrado);

            // Assert
            resultado.Should().HaveCount(2)
                .And.OnlyContain(p => p.Modalidad == "PREGRADO");
        }

        [Fact]
        public void EtiquetaPrograma_WhenAbreviaturaPresente_ShouldReturnAbreviaturaTrimmed()
        {
            // Arrange
            var prog = new Programa { Id = 1, NombrePrograma = "ADMINISTRACION PUBLICA", Abreviatura = "  AP  " };

            // Act
            string etiqueta = AgTipoProgramaCatalog.EtiquetaPrograma(prog);

            // Assert
            etiqueta.Should().Be("AP");
        }

        [Fact]
        public void EtiquetaPrograma_WhenAbreviaturaVacia_ShouldFallbackToNombreProgramaTrimmed()
        {
            // Arrange
            var prog = new Programa { Id = 2, NombrePrograma = "  ESPECIALIZACION EN GESTION  ", Abreviatura = "   " };

            // Act
            string etiqueta = AgTipoProgramaCatalog.EtiquetaPrograma(prog);

            // Assert
            etiqueta.Should().Be("ESPECIALIZACION EN GESTION");
        }

        [Fact]
        public void EtiquetaPrograma_WhenProgramaNull_ShouldReturnEmpty()
        {
            // Act
            string etiqueta = AgTipoProgramaCatalog.EtiquetaPrograma(null);

            // Assert
            etiqueta.Should().BeEmpty();
        }

        [Fact]
        public void Filtrar_WhenProgramasInactivos_ShouldIncludeThemRegardlessOfActivoFlag()
        {
            // Arrange (Fase 1466: el catálogo AG ya no filtra por activo)
            var programas = new List<Programa>
            {
                new Programa { Id = 1, NombrePrograma = "APT-53", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 9, NombrePrograma = "BEE-505", Modalidad = "POSGRADO", Activo = 0 },
                new Programa { Id = 21, NombrePrograma = "MDH-409", Modalidad = "POSGRADO", Activo = 0 }
            };

            // Act
            var posgrados = AgTipoProgramaCatalog.Filtrar(programas, AgTipoProgramaCatalog.Posgrado);

            // Assert
            posgrados.Should().HaveCount(2)
                .And.OnlyContain(p => p.Modalidad == "POSGRADO" && p.Activo == 0);
        }
    }
}
