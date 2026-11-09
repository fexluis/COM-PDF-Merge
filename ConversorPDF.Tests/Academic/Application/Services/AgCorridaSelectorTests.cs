using System.Collections.Generic;
using System.Linq;
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
            string subunidad, int sesiones, int estado = 0)
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
                TotalSesiones = sesiones
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
                new CorridaStagingItem { IdCorrida = 7, TotalSesiones = 3 });

            etiqueta.Should().Be("#7 · 3 ses");
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
    }
}
