using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class ViaticoConflictAnalyzerTests
    {
        // ─── DetectarBloquesCruzados ──────────────────────────────────────────────

        [Fact]
        public void DetectarBloquesCruzados_WhenInputIsNull_ShouldReturnEmpty()
        {
            // Act
            var result = ViaticoConflictAnalyzer.DetectarBloquesCruzados(null);

            // Assert
            result.Should().BeEmpty();
        }

        [Fact]
        public void DetectarBloquesCruzados_WhenSingleBloque_ShouldReturnEmpty()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 1, 1), SedeCiudad = "Sede A" };
            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var input = new List<(string Sede, ViaticoBloque Bloque)> { ("Sede A", b1) };

            // Act
            var result = ViaticoConflictAnalyzer.DetectarBloquesCruzados(input);

            // Assert
            result.Should().BeEmpty();
        }

        [Fact]
        public void DetectarBloquesCruzados_WhenNoConflicts_ShouldReturnEmpty()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 1, 1), SedeCiudad = "Sede A" };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 1, 3), SedeCiudad = "Sede B" };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                ("Sede A", b1),
                ("Sede B", b2)
            };

            // Act
            var result = ViaticoConflictAnalyzer.DetectarBloquesCruzados(input);

            // Assert
            result.Should().BeEmpty();
        }

        [Fact]
        public void DetectarBloquesCruzados_WhenOverlapInDifferentSedes_ShouldReturnBothBloques()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 2, 1), SedeCiudad = "Sede A" };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 2, 1), SedeCiudad = "Sede B" };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                ("Sede A", b1),
                ("Sede B", b2)
            };

            // Act
            var result = ViaticoConflictAnalyzer.DetectarBloquesCruzados(input);

            // Assert
            result.Should().HaveCount(2);
            result.Should().Contain(b1);
            result.Should().Contain(b2);
        }

        [Fact]
        public void DetectarBloquesCruzados_WhenOverlapInSameSede_ShouldNotReturnConflicts()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 3, 1), SedeCiudad = "Sede A" };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 3, 1), SedeCiudad = "Sede A" };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                ("Sede A", b1),
                ("Sede A", b2)
            };

            // Act
            var result = ViaticoConflictAnalyzer.DetectarBloquesCruzados(input);

            // Assert
            result.Should().BeEmpty();
        }

        // ─── ObtenerDetalleCruce ─────────────────────────────────────────────────

        [Fact]
        public void ObtenerDetalleCruce_WhenConflictExists_ShouldReturnInformativeText()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 4, 1), SedeCiudad = "Sede A", Horas = 4 };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 4, 1), SedeCiudad = "Sede B", Horas = 2 };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                ("Sede A", b1),
                ("Sede B", b2)
            };

            // Act
            var detail = ViaticoConflictAnalyzer.ObtenerDetalleCruce(b1, input, "Sede A");

            // Assert
            detail.Should().Contain("📍 SEDE B");
            detail.Should().Contain("01/04/2026");
        }

        [Fact]
        public void ObtenerDetalleCruce_WhenNoConflictExists_ShouldReturnEmpty()
        {
            // Arrange — bloques en fechas que NO se solapan
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 6, 1), SedeCiudad = "Sede A" };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 6, 5), SedeCiudad = "Sede B" };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                ("Sede A", b1),
                ("Sede B", b2)
            };

            // Act
            var detail = ViaticoConflictAnalyzer.ObtenerDetalleCruce(b1, input, "Sede A");

            // Assert
            detail.Should().BeEmpty();
        }

        [Fact]
        public void ObtenerDetalleCruce_WhenSedeActualIsNull_ShouldUseFallbackText()
        {
            // Arrange
            var s1 = new SesionListItem { FechaSesion = new DateTime(2026, 7, 1), SedeCiudad = null };
            var s2 = new SesionListItem { FechaSesion = new DateTime(2026, 7, 1), SedeCiudad = "Sede B" };

            var b1 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s1 }).First();
            var b2 = ViaticoBloque.AgruparConsecutivos(new List<SesionListItem> { s2 }).First();

            var input = new List<(string Sede, ViaticoBloque Bloque)>
            {
                (null, b1),
                ("Sede B", b2)
            };

            // Act — sedeActual null no debe lanzar NullReferenceException (R6)
            var detail = ViaticoConflictAnalyzer.ObtenerDetalleCruce(b1, input, null);

            // Assert
            detail.Should().Contain("SIN SEDE");
        }
    }
}
