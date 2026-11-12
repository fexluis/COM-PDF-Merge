using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    /// <summary>
    /// Pruebas del selector de corridas del MODO AG (Fase 1495).
    /// </summary>
    public class AgCorridaSelectorTests
    {
        private static CorridaStagingItem Corrida(
            long id, int idPeriodo, string periodo, int idPrograma, string programa,
            string subunidad, int sesiones, int estado = 0, long? idIndividuoOptimo = 1)
        {
            return new CorridaStagingItem
            {
                IdCorrida = id,
                IdPeriodo = idPeriodo,
                PeriodoNombre = periodo,
                IdPrograma = idPrograma,
                ProgramaNombre = programa,
                Subunidad = subunidad,
                FechaInicioCorrida = "2026-11-09 09:31:00",
                EstadoCorrida = estado,
                TotalVinculaciones = sesiones,
                TotalSesiones = sesiones,
                IdIndividuoOptimo = idIndividuoOptimo
            };
        }

        private static List<CorridaStagingItem> FixtureReal()
        {
            // Reproduce el vault real: #46 (BOLIVAR, Error, sin staging), #44 (VALLE, 244 ses),
            // #37 (VALLE, sin staging), #1 (periodo anterior, 145 ses).
            return new List<CorridaStagingItem>
            {
                Corrida(46, 3, "2027-01", 20, "MAESTRIA", "BOLIVAR", 0, estado: 4),
                Corrida(44, 3, "2027-01", 1, "ADMINISTRACION PUBLICA TERRITORIAL", "VALLE", 244, estado: 2),
                Corrida(37, 3, "2027-01", 2, "OTRO PROGRAMA", "VALLE", 0, estado: 2),
                Corrida(1, 2, "2026-02", 1, "ADMINISTRACION PUBLICA TERRITORIAL", "VALLE", 145, estado: 2)
            };
        }

        [Fact]
        public void ResolverSeleccionInicial_PrefiereCorridaEnUso_SobreMayorId()
        {
            var elegida = AgCorridaSelector.ResolverSeleccionInicial(FixtureReal(), 44);

            elegida.IdCorrida.Should().Be(44);
        }

        [Fact]
        public void ResolverSeleccionInicial_SinCorridaEnUso_EligePrimeraConStaging()
        {
            var elegida = AgCorridaSelector.ResolverSeleccionInicial(FixtureReal(), null);

            elegida.IdCorrida.Should().Be(44, "la #46 no tiene staging y no debe preseleccionarse");
        }

        [Fact]
        public void ResolverSeleccionInicial_SinStaging_DevuelveTodasLasCorridas()
        {
            var sinStaging = new List<CorridaStagingItem>
            {
                Corrida(46, 3, "2027-01", 20, "MAESTRIA", "BOLIVAR", 0, estado: 4),
                Corrida(37, 3, "2027-01", 2, "OTRO", "VALLE", 0, estado: 2)
            };

            var elegida = AgCorridaSelector.ResolverSeleccionInicial(sinStaging, null);

            elegida.IdCorrida.Should().Be(0);
            elegida.Etiqueta.Should().Be(AgCorridaSelector.EtiquetaTodasLasCorridas);
        }

        [Fact]
        public void ResolverSeleccionInicial_CorridaEnUsoInexistente_CaeEnLaPrimeraConStaging()
        {
            var elegida = AgCorridaSelector.ResolverSeleccionInicial(FixtureReal(), 999);

            elegida.IdCorrida.Should().Be(44);
        }

        [Fact]
        public void FormatearEtiqueta_ConStaging_ContienePeriodoProgramaSubunidadYSesiones()
        {
            var etiqueta = AgCorridaSelector.FormatearEtiqueta(
                Corrida(44, 3, "2027-01", 1, "ADMINISTRACION PUBLICA TERRITORIAL", "VALLE", 244, estado: 2));

            etiqueta.Should().Be("#44 · 2027-01 · ADMINISTRACION PUBLICA TERRITORIAL · VALLE · 244 ses");
        }

        [Fact]
        public void FormatearEtiqueta_SinStaging_MarcaSinSesionesYEstadoError()
        {
            var etiqueta = AgCorridaSelector.FormatearEtiqueta(
                Corrida(46, 3, "2027-01", 20, "MAESTRIA", "BOLIVAR", 0, estado: 4));

            etiqueta.Should().Contain("(sin sesiones)");
            etiqueta.Should().Contain("Error");
            etiqueta.Should().NotContain("· ·");
        }

        [Fact]
        public void FormatearEtiqueta_CamposVacios_NoDejaSeparadoresHuerfanos()
        {
            var etiqueta = AgCorridaSelector.FormatearEtiqueta(
                new CorridaStagingItem { IdCorrida = 7, TotalSesiones = 3, IdIndividuoOptimo = 1 });

            etiqueta.Should().Be("#7 · 3 ses");
        }

        [Fact]
        public void FormatearEtiqueta_SinIndividuoOptimo_AgregaSufijoSinResultadoAG()
        {
            var etiqueta = AgCorridaSelector.FormatearEtiqueta(
                Corrida(48, 3, "2027-01", 1, "ADMINISTRACION PUBLICA TERRITORIAL", "VALLE", 180, estado: 0, idIndividuoOptimo: null));

            etiqueta.Should().Be("#48 · 2027-01 · ADMINISTRACION PUBLICA TERRITORIAL · VALLE · 180 ses · sin resultado AG");
        }

        [Fact]
        public void PrepararParaCombo_AnteponeComodinYPueblaEtiquetas()
        {
            var items = AgCorridaSelector.PrepararParaCombo(FixtureReal()).ToList();

            items.Should().HaveCount(5);
            items[0].IdCorrida.Should().Be(0);
            items[0].Etiqueta.Should().Be(AgCorridaSelector.EtiquetaTodasLasCorridas);
            items.Skip(1).Should().OnlyContain(c => !string.IsNullOrWhiteSpace(c.Etiqueta));
            items[1].Etiqueta.Should().StartWith("#46");
        }

        [Fact]
        public void PrepararParaCombo_ColeccionNula_DevuelveSoloComodin()
        {
            var items = AgCorridaSelector.PrepararParaCombo(null).ToList();

            items.Should().HaveCount(1);
            items[0].IdCorrida.Should().Be(0);
        }

        private static readonly (int? IdPeriodo, int? IdPrograma)[] EsperadoPeriodoYPrograma =
        {
            (2, 1), (2, null), (null, null)
        };

        [Fact]
        public void CascadaFiltros_ConPeriodoYPrograma_DevuelveTresIntentos()
        {
            AgCorridaSelector.CascadaFiltros(2, 1).Should().Equal(EsperadoPeriodoYPrograma);
        }

        [Fact]
        public void CascadaFiltros_SinPrograma_DevuelveDosIntentos()
        {
            AgCorridaSelector.CascadaFiltros(2, null)
                .Should().Equal(new (int?, int?)[] { (2, null), (null, null) });
        }

        [Fact]
        public void CascadaFiltros_SinPeriodo_DevuelveDosIntentos()
        {
            AgCorridaSelector.CascadaFiltros(null, 5)
                .Should().Equal(new (int?, int?)[] { (null, 5), (null, null) });
        }

        [Fact]
        public void CascadaFiltros_TodoNulo_DevuelveUnIntento()
        {
            var unico = AgCorridaSelector.CascadaFiltros(null, null).Should().ContainSingle().Subject;
            unico.IdPeriodo.Should().BeNull();
            unico.IdPrograma.Should().BeNull();
        }

        [Theory]
        [InlineData(2, 1)]
        [InlineData(2, null)]
        [InlineData(null, 5)]
        [InlineData(null, null)]
        public void CascadaFiltros_SiempreTerminaAmplio(int? idPeriodo, int? idPrograma)
        {
            var ultimo = AgCorridaSelector.CascadaFiltros(idPeriodo, idPrograma).Last();
            ultimo.IdPeriodo.Should().BeNull();
            ultimo.IdPrograma.Should().BeNull();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoActivaPresente_DebeSeleccionarActivaYSinCargaInicial()
        {
            // Arrange
            var corridas = FixtureReal();
            long idActiva = 44;

            // Act
            var seleccion = AgCorridaSelector.ResolverSeleccionModuloR(corridas, idActiva);

            // Assert
            seleccion.Should().NotBeNull();
            seleccion.ItemSeleccionado.Should().NotBeNull();
            seleccion.ItemSeleccionado.IdCorrida.Should().Be(44);
            seleccion.IndiceSeleccionado.Should().Be(1);
            seleccion.RequiereCargaInicial.Should().BeFalse();
            seleccion.CorridaActivaNoEncontrada.Should().BeFalse();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoActivaAusenteConStagingDisponible_NoDebeSeleccionarNiCargarYMarcaNoEncontrada()
        {
            // Arrange
            var corridas = FixtureReal();
            long idActivaInexistente = 999;

            // Act
            var seleccion = AgCorridaSelector.ResolverSeleccionModuloR(corridas, idActivaInexistente);

            // Assert
            seleccion.Should().NotBeNull();
            seleccion.ItemSeleccionado.Should().BeNull();
            seleccion.IndiceSeleccionado.Should().Be(-1);
            seleccion.RequiereCargaInicial.Should().BeFalse();
            seleccion.CorridaActivaNoEncontrada.Should().BeTrue();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoActivaAusenteConSoloCabeceras_NoDebeSeleccionarNiCargarYMarcaNoEncontrada()
        {
            // Arrange
            var soloCabeceras = new List<CorridaStagingItem>
            {
                Corrida(46, 3, "2027-01", 20, "MAESTRIA", "BOLIVAR", 0, estado: 4),
                Corrida(37, 3, "2027-01", 2, "OTRO", "VALLE", 0, estado: 2)
            };
            long idActivaInexistente = 999;

            // Act
            var seleccion = AgCorridaSelector.ResolverSeleccionModuloR(soloCabeceras, idActivaInexistente);

            // Assert
            seleccion.Should().NotBeNull();
            seleccion.ItemSeleccionado.Should().BeNull();
            seleccion.IndiceSeleccionado.Should().Be(-1);
            seleccion.RequiereCargaInicial.Should().BeFalse();
            seleccion.CorridaActivaNoEncontrada.Should().BeTrue();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoSinActivaYConStagingDisponible_DebeSeleccionarPrimeraConStagingYRequiereCargaInicial()
        {
            // Arrange
            var corridas = FixtureReal();

            // Act
            var seleccion = AgCorridaSelector.ResolverSeleccionModuloR(corridas, null);

            // Assert
            seleccion.Should().NotBeNull();
            seleccion.ItemSeleccionado.Should().NotBeNull();
            seleccion.ItemSeleccionado.IdCorrida.Should().Be(44);
            seleccion.IndiceSeleccionado.Should().Be(1);
            seleccion.RequiereCargaInicial.Should().BeTrue();
            seleccion.CorridaActivaNoEncontrada.Should().BeFalse();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoSinActivaYSoloCabeceras_NoDebeSeleccionarNiCargar()
        {
            // Arrange
            var soloCabeceras = new List<CorridaStagingItem>
            {
                Corrida(46, 3, "2027-01", 20, "MAESTRIA", "BOLIVAR", 0, estado: 4),
                Corrida(37, 3, "2027-01", 2, "OTRO", "VALLE", 0, estado: 2)
            };

            // Act
            var seleccion = AgCorridaSelector.ResolverSeleccionModuloR(soloCabeceras, null);

            // Assert
            seleccion.Should().NotBeNull();
            seleccion.ItemSeleccionado.Should().BeNull();
            seleccion.IndiceSeleccionado.Should().Be(-1);
            seleccion.RequiereCargaInicial.Should().BeFalse();
            seleccion.CorridaActivaNoEncontrada.Should().BeFalse();
        }

        [Fact]
        public void ResolverSeleccionModuloR_CuandoListaVaciaONula_NoDebeSeleccionarNiCargar()
        {
            // Arrange
            var vacia = new List<CorridaStagingItem>();

            // Act
            var seleccionVacia = AgCorridaSelector.ResolverSeleccionModuloR(vacia, null);
            var seleccionNula = AgCorridaSelector.ResolverSeleccionModuloR(null, null);

            // Assert
            seleccionVacia.Should().NotBeNull();
            seleccionVacia.ItemSeleccionado.Should().BeNull();
            seleccionVacia.IndiceSeleccionado.Should().Be(-1);
            seleccionVacia.RequiereCargaInicial.Should().BeFalse();
            seleccionVacia.CorridaActivaNoEncontrada.Should().BeFalse();

            seleccionNula.Should().NotBeNull();
            seleccionNula.ItemSeleccionado.Should().BeNull();
            seleccionNula.IndiceSeleccionado.Should().Be(-1);
            seleccionNula.RequiereCargaInicial.Should().BeFalse();
            seleccionNula.CorridaActivaNoEncontrada.Should().BeFalse();
        }
    }
}
