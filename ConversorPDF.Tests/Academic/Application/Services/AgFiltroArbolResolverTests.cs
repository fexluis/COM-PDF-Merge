using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgFiltroArbolResolverTests
    {
        [Fact]
        public void FiltrarPeriodos_WhenIdsVacio_ShouldReturnAllActivePeriodos()
        {
            // Arrange
            var periodos = new List<Periodo>
            {
                new Periodo { Id = 1, PeriodoNombre = "2024-1", Activo = true },
                new Periodo { Id = 2, PeriodoNombre = "2024-2", Activo = true },
                new Periodo { Id = 3, PeriodoNombre = "2023-2", Activo = false }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarPeriodos(periodos, new List<int>());

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(p => p.Activo);
        }

        [Fact]
        public void FiltrarPeriodos_WhenIdsEspecificos_ShouldReturnMatchingPeriodos()
        {
            // Arrange
            var periodos = new List<Periodo>
            {
                new Periodo { Id = 1, PeriodoNombre = "2024-1", Activo = true },
                new Periodo { Id = 2, PeriodoNombre = "2024-2", Activo = true }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarPeriodos(periodos, new[] { 2 });

            // Assert
            res.Should().ContainSingle()
                .Which.Id.Should().Be(2);
        }

        [Fact]
        public void FiltrarSedes_WhenIdsVacio_ShouldReturnAllActiveSedes()
        {
            // Arrange
            var sedes = new List<Sede>
            {
                new Sede { Id = 10, Ciudad = "Bogota", Activo = 1 },
                new Sede { Id = 20, Ciudad = "Medellin", Activo = 1 },
                new Sede { Id = 30, Ciudad = "Cali", Activo = 0 }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarSedes(sedes, new List<int>());

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(s => s.Activo == 1);
        }

        [Fact]
        public void FiltrarProgramas_WhenFiltroTipoYIds_ShouldIntersectConditions()
        {
            // Arrange
            var programas = new List<Programa>
            {
                new Programa { Id = 101, NombrePrograma = "Pregrado A", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 102, NombrePrograma = "Pregrado B", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 103, NombrePrograma = "Posgrado C", Modalidad = "POSGRADO", Activo = 1 }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarProgramas(programas, "PREGRADO", new[] { 101 });

            // Assert
            res.Should().ContainSingle()
                .Which.Id.Should().Be(101);
        }

        [Fact]
        public void FiltrarProgramas_WhenIdsVacio_ShouldReturnAllProgramsMatchingTipo()
        {
            // Arrange
            var programas = new List<Programa>
            {
                new Programa { Id = 101, NombrePrograma = "Pregrado A", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 102, NombrePrograma = "Pregrado B", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 103, NombrePrograma = "Posgrado C", Modalidad = "POSGRADO", Activo = 1 }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarProgramas(programas, "PREGRADO", new List<int>());

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(p => p.Modalidad == "PREGRADO");
        }

        [Fact]
        public void FiltrarProgramas_WhenTipoTodos_ShouldReturnAllActivePrograms()
        {
            // Arrange
            var programas = new List<Programa>
            {
                new Programa { Id = 101, NombrePrograma = "Pregrado A", Modalidad = "PREGRADO", Activo = 1 },
                new Programa { Id = 103, NombrePrograma = "Posgrado C", Modalidad = "POSGRADO", Activo = 1 }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarProgramas(programas, "TODOS", new List<int>());

            // Assert
            res.Should().HaveCount(2);
        }

        [Fact]
        public void FiltrarProgramas_WhenNoProgramsMatchIds_ShouldReturnEmpty()
        {
            // Arrange
            var programas = new List<Programa>
            {
                new Programa { Id = 101, NombrePrograma = "Pregrado A", Modalidad = "PREGRADO", Activo = 1 }
            };

            // Act
            var res = AgFiltroArbolResolver.FiltrarProgramas(programas, "PREGRADO", new[] { 999 });

            // Assert
            res.Should().BeEmpty();
        }
    }
}
