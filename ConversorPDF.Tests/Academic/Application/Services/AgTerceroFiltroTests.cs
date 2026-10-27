using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgTerceroFiltroTests
    {
        private readonly List<AgTerceroSeleccionDto> _sample = new List<AgTerceroSeleccionDto>
        {
            new AgTerceroSeleccionDto
            {
                Id = 1,
                Nombre = "Carlos Perez",
                Cedula = "12345",
                Departamento = "Cundinamarca",
                Residencia = "Bogota",
                TieneAsignaturasHabilitadas = true,
                Activo = true
            },
            new AgTerceroSeleccionDto
            {
                Id = 2,
                Nombre = "Maria Lopez",
                Cedula = "67890",
                Departamento = "Antioquia",
                Residencia = "Medellin",
                TieneAsignaturasHabilitadas = false,
                Activo = true
            },
            new AgTerceroSeleccionDto
            {
                Id = 3,
                Nombre = "Pedro Gomez",
                Cedula = "11223",
                Departamento = "Cundinamarca",
                Residencia = "Soacha",
                TieneAsignaturasHabilitadas = true,
                Activo = false
            }
        };

        [Fact]
        public void Filtrar_WhenTerritorialSpecified_ShouldReturnOnlyMatchingTerritorial()
        {
            // Act
            var res = AgTerceroFiltro.Filtrar(_sample, "Cundinamarca", null, false, true);

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(t => t.Departamento == "Cundinamarca");
        }

        [Fact]
        public void Filtrar_WhenBusquedaMatchesCedula_ShouldReturnMatched()
        {
            // Act
            var res = AgTerceroFiltro.Filtrar(_sample, null, "67890", false, true);

            // Assert
            res.Should().ContainSingle()
                .Which.Nombre.Should().Be("Maria Lopez");
        }

        [Fact]
        public void Filtrar_WhenSoloHabilitadasTrue_ShouldExcludeNonHabilitadas()
        {
            // Act
            var res = AgTerceroFiltro.Filtrar(_sample, null, null, soloHabilitadas: true, mostrarInactivos: true);

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(t => t.TieneAsignaturasHabilitadas);
        }

        [Fact]
        public void Filtrar_WhenMostrarInactivosFalse_ShouldExcludeInactivos()
        {
            // Act
            var res = AgTerceroFiltro.Filtrar(_sample, null, null, soloHabilitadas: false, mostrarInactivos: false);

            // Assert
            res.Should().HaveCount(2)
                .And.OnlyContain(t => t.Activo);
        }

        [Fact]
        public void Filtrar_WhenTerceroTieneSubunidadCompuesta_ShouldCoincidirConCadaParte()
        {
            var origen = new List<AgTerceroSeleccionDto>
            {
                new AgTerceroSeleccionDto { Id = 151, Nombre = "FERNANDO FONTAL BUENO", Cedula = "1",
                    Departamento = "BOLIVAR,VALLE", Subunidad = "BOLIVAR,VALLE", Activo = true },
                new AgTerceroSeleccionDto { Id = 152, Nombre = "SOLO BOLIVAR", Cedula = "2",
                    Departamento = "BOLIVAR", Subunidad = "BOLIVAR", Activo = true }
            };

            AgTerceroFiltro.Filtrar(origen, "VALLE", null, false, true).Should().ContainSingle()
                .Which.Id.Should().Be(151);

            AgTerceroFiltro.Filtrar(origen, "BOLIVAR", null, false, true).Should().HaveCount(2);
        }

        [Fact]
        public void Filtrar_WhenFiltroEsTodasLasTerritoriales_ShouldDevolverTodos()
        {
            var res = AgTerceroFiltro.Filtrar(_sample, AgTerritorialCatalog.Todas, null, false, true);

            res.Should().HaveCount(3);
        }

        [Fact]
        public void Filtrar_WhenBusquedaCoincideSubunidad_ShouldReturnMatched()
        {
            var origen = new List<AgTerceroSeleccionDto>
            {
                new AgTerceroSeleccionDto { Id = 7, Nombre = "ANA", Cedula = "9",
                    Departamento = "BOLIVAR", Subunidad = "VALLE", Activo = true },
                new AgTerceroSeleccionDto { Id = 8, Nombre = "LUIS", Cedula = "10",
                    Departamento = "BOLIVAR", Subunidad = "BOLIVAR", Activo = true }
            };

            var res = AgTerceroFiltro.Filtrar(origen, null, "VALLE", false, true);

            res.Should().ContainSingle().Which.Id.Should().Be(7);
        }

        [Fact]
        public void Filtrar_WhenSoloHabilitadasYSoloTieneOtras_ShouldExcluir()
        {
            var origen = new List<AgTerceroSeleccionDto>
            {
                new AgTerceroSeleccionDto { Id = 20, Nombre = "LOCAL", Cedula = "20", Activo = true,
                    HabilitacionesLocales = 2, HabilitacionesOtras = 0 },
                new AgTerceroSeleccionDto { Id = 21, Nombre = "OTRAS", Cedula = "21", Activo = true,
                    HabilitacionesLocales = 0, HabilitacionesOtras = 3 }
            };

            var res = AgTerceroFiltro.Filtrar(origen, null, null, soloHabilitadas: true, mostrarInactivos: true);

            res.Should().ContainSingle().Which.Id.Should().Be(20);
            origen[1].TotalHabilitaciones.Should().Be(3);
            origen[1].TieneAsignaturasHabilitadas.Should().BeFalse();
        }
    }
}
