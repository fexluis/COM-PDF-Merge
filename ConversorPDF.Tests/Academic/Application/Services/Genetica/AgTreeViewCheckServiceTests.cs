using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgTreeViewCheckServiceTests
    {
        [Fact]
        public void ExpandirCombinaciones_WhenValidSelections_ShouldReturnExactCartesianProduct()
        {
            // Arrange
            var sut = new AgTreeViewCheckService();
            var periodos = new[] { 1 };
            var sedes = new[] { 2 };
            var programas = new[] { 3 };
            var semestres = new[] { "I", "II" };
            var grupos = new[] { "A", "B" };
            var modalidades = new[] { "Presencial", "Plataforma" };

            // Act: 1 * 1 * 1 * 2 * 2 * 2 = 8 combinaciones
            var resultado = sut.ExpandirCombinaciones(periodos, sedes, programas, semestres, grupos, modalidades);

            // Assert
            resultado.Should().HaveCount(8);
            resultado.Should().Contain(c => c.Semestre == "I" && c.Grupo == "A" && c.Modalidad == "Presencial");
            resultado.Should().Contain(c => c.Semestre == "II" && c.Grupo == "B" && c.Modalidad == "Plataforma");
        }

        [Fact]
        public void ExpandirCombinaciones_WhenAnyLevelEmpty_ShouldReturnEmptyList()
        {
            // Arrange
            var sut = new AgTreeViewCheckService();

            // Act
            var resultado = sut.ExpandirCombinaciones(
                new[] { 1 },
                new[] { 2 },
                new int[0], // programas vacíos
                new[] { "I" },
                new[] { "A" },
                new[] { "Presencial" });

            // Assert
            resultado.Should().BeEmpty();
        }
    }
}
