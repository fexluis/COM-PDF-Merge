using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgPerfilAlcanceResolverTests
    {
        private static AgContextoSeleccionDto Dto(int idPeriodo, int idSede, int idPrograma)
        {
            return new AgContextoSeleccionDto
            {
                IdPeriodo = idPeriodo,
                IdSede = idSede,
                IdPrograma = idPrograma,
                Semestre = "VI",
                Grupo = "A",
                Modalidad = "Presencial"
            };
        }

        [Fact]
        public void Resolver_WhenPerfilCubreVariasSedesActivas_ShouldIncludeAllOfThem()
        {
            // Arrange (Fase 1580): el perfil guardado abarca 3 sedes activas
            var nodos = new List<AgContextoSeleccionDto>
            {
                Dto(3, 11, 1), Dto(3, 12, 1), Dto(3, 13, 1), Dto(3, 13, 2)
            };
            var periodos = new List<Periodo> { new Periodo { Id = 3, Activo = true } };
            var sedes = new List<Sede>
            {
                new Sede { Id = 11, Ciudad = "CALI", Activo = 1 },
                new Sede { Id = 12, Ciudad = "BOGOTA", Activo = 1 },
                new Sede { Id = 13, Ciudad = "MEDELLIN", Activo = 1 }
            };
            var programas = new List<Programa>
            {
                new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" },
                new Programa { Id = 2, Modalidad = "PREGRADO VIRTUAL" }
            };

            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(nodos, periodos, sedes, programas);

            // Assert
            alcance.Sedes.Should().BeEquivalentTo(new[] { 11, 12, 13 });
            alcance.Programas.Should().BeEquivalentTo(new[] { 1, 2 });
            alcance.Periodos.Should().BeEquivalentTo(new[] { 3 });
            alcance.CombosFueraDeCatalogo.Should().Be(0);
        }

        [Fact]
        public void Resolver_WhenSedeOProgramaNoExisteEnCatalogo_ShouldCountCombosFueraDeCatalogo()
        {
            // Arrange: una sede desactivada/inexistente y un programa eliminado
            var nodos = new List<AgContextoSeleccionDto>
            {
                Dto(3, 13, 1),
                Dto(3, 99, 1),
                Dto(3, 13, 77)
            };
            var periodos = new List<Periodo> { new Periodo { Id = 3, Activo = true } };
            var sedes = new List<Sede> { new Sede { Id = 13, Ciudad = "MEDELLIN", Activo = 1 } };
            var programas = new List<Programa> { new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" } };

            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(nodos, periodos, sedes, programas);

            // Assert
            alcance.Sedes.Should().BeEquivalentTo(new[] { 13 });
            alcance.Programas.Should().BeEquivalentTo(new[] { 1 });
            alcance.CombosFueraDeCatalogo.Should().Be(2);
        }

        [Fact]
        public void Resolver_WhenProgramasSonMixtos_ShouldReturnTipoTodos()
        {
            // Arrange
            var nodos = new List<AgContextoSeleccionDto> { Dto(3, 13, 1), Dto(3, 13, 2) };
            var programas = new List<Programa>
            {
                new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" },
                new Programa { Id = 2, Modalidad = "POSGRADO ESPECIALIZACION" }
            };

            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(nodos, new List<Periodo>(), new List<Sede>(), programas);

            // Assert
            alcance.TipoPrograma.Should().Be(AgTipoProgramaCatalog.Todos);
        }

        [Fact]
        public void Resolver_WhenProgramasSonSoloPosgrado_ShouldReturnTipoPosgrado()
        {
            // Arrange
            var nodos = new List<AgContextoSeleccionDto> { Dto(3, 13, 2) };
            var programas = new List<Programa> { new Programa { Id = 2, Modalidad = "POSGRADO ESPECIALIZACION" } };

            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(nodos, new List<Periodo>(), new List<Sede>(), programas);

            // Assert
            alcance.TipoPrograma.Should().Be(AgTipoProgramaCatalog.Posgrado);
        }

        [Fact]
        public void Resolver_WhenProgramasSonSoloPregrado_ShouldReturnTipoPregrado()
        {
            // Arrange
            var nodos = new List<AgContextoSeleccionDto> { Dto(3, 13, 1) };
            var programas = new List<Programa> { new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" } };

            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(nodos, new List<Periodo>(), new List<Sede>(), programas);

            // Assert
            alcance.TipoPrograma.Should().Be(AgTipoProgramaCatalog.Pregrado);
        }

        [Fact]
        public void Resolver_WhenNodosVacio_ShouldReturnEmptyScope()
        {
            // Act
            var alcance = AgPerfilAlcanceResolver.Resolver(
                Enumerable.Empty<AgContextoSeleccionDto>(),
                new List<Periodo>(),
                new List<Sede>(),
                new List<Programa>());

            // Assert
            alcance.Periodos.Should().BeEmpty();
            alcance.Sedes.Should().BeEmpty();
            alcance.Programas.Should().BeEmpty();
            alcance.TipoPrograma.Should().Be(AgTipoProgramaCatalog.Todos);
            alcance.CombosFueraDeCatalogo.Should().Be(0);
        }
    }
}
