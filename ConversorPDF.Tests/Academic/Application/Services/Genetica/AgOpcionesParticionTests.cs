using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgOpcionesParticionTests
    {
        [Fact]
        public void MarVie_Cabe_NoEscala_YParticionOrigen()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.Accion.Should().Be("quedarse");
            opcion.CarrilesTexto.Should().Be("2,3|4,5");
            opcion.Carriles.Should().HaveCount(2);
            opcion.Carriles[0].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 2, 3 });
            opcion.Carriles[1].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 4, 5 });
        }

        [Fact]
        public void MarVie_Cabe_ParticionEs_2_3_y_4_5()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.CarrilesTexto.Should().Be("2,3|4,5");
            opcion.Carriles[0].Select(r => r.Dia).Should().Equal(2, 3);
            opcion.Carriles[1].Select(r => r.Dia).Should().Equal(4, 5);
        }

        [Fact]
        public void MarVie_NoCabe_Escala_LunVie_Con_2_3_y_1_4_5()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("LUN-VIE", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.CarrilesTexto.Should().Be("2,3|1,4,5");
            opcion.Carriles[0].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 2, 3 });
            opcion.Carriles[1].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 1, 4, 5 });
        }

        [Fact]
        public void LunJue_NoCabe_Escala_LunVie_Con_1_2_5_y_3_4()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "1,2", "3,4" };

            var opcion = AgOpcionesParticion.Resolver("LUN-VIE", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.CarrilesTexto.Should().Be("1,2,5|3,4");
            opcion.Carriles[0].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 1, 2, 5 });
            opcion.Carriles[1].Select(r => r.Dia).Should().BeEquivalentTo(new[] { 3, 4 });
        }

        [Fact]
        public void YaEnLunVie_MarVie_UsaOpcionUnica_SinEnumerar()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "2,3" };

            var opcion = AgOpcionesParticion.Resolver("LUN-VIE", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.Accion.Should().Be("quedarse");
            opcion.CarrilesTexto.Should().Be("2,3|1,4,5");
        }

        [Fact]
        public void SinPatrones_SinOpciones_ComportamientoActual()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, new List<string>());

            opcion.Should().BeNull();
        }

        [Fact]
        public void CuatroCincoSeisSiete_Usa_5_6M_y_4_6T_7()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.TARDE, 4.0),
                new AgRanura(5, AgFranjaJornada.TARDE, 4.0),
                new AgRanura(6, AgFranjaJornada.MANANA, 4.0),
                new AgRanura(6, AgFranjaJornada.TARDE, 4.0),
                new AgRanura(7, AgFranjaJornada.MANANA, 4.0)
            };
            var patrones = new List<string> { "5,6", "6,7" };

            var opcion = AgOpcionesParticion.Resolver("4,5,6,7", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.CarrilesTexto.Should().Be("5,6M|4,6T,7");
            opcion.Carriles.Should().HaveCount(2);
            opcion.Carriles[0].Should().HaveCount(2);
            opcion.Carriles[1].Should().HaveCount(3);
        }

        [Fact]
        public void OchoHoras_SinFranjaTarde_NoProduceSemilla()
        {
            // Sábado 8h solo tiene ranura MANANA; el carril 5,6M|4,6T,7 exige 6T -> tokens no resolubles -> null
            var ranuras8h = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.MANANA, 8.0),
                new AgRanura(5, AgFranjaJornada.MANANA, 8.0),
                new AgRanura(6, AgFranjaJornada.MANANA, 8.0),
                new AgRanura(7, AgFranjaJornada.MANANA, 8.0)
            };
            var patrones = new List<string> { "5,6", "6,7" };

            var opcion = AgOpcionesParticion.Resolver("4,5,6,7", ranuras8h, patrones);

            opcion.Should().BeNull();
        }

        [Fact]
        public void PatronTerciario_3_4_5_NoClasificaMarVie()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "3,4,5" };

            var origen = AgOpcionesParticion.DetectarOrigen(patrones);
            origen.Should().Be(AgOpcionesParticion.Origen.Ninguno);

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            opcion.Should().BeNull();
        }

        [Fact]
        public void Mixto_1_2_y_2_3_NoAplicaCatalogo()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "1,2", "2,3" };

            var origen = AgOpcionesParticion.DetectarOrigen(patrones);
            origen.Should().Be(AgOpcionesParticion.Origen.Ninguno);

            var opcion = AgOpcionesParticion.Resolver("LUN-VIE", ranuras, patrones);
            opcion.Should().BeNull();
        }

        [Fact]
        public void CandidatasEvaluadas_NuncaSuperaDos()
        {
            AgOpcionesParticion.MaxOpcionesDecision.Should().Be(2);
        }

        [Fact]
        public void Resolver_Destino_MarcaAccionEscalar()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("LUN-VIE", ranuras, patrones, "escalar");

            opcion.Should().NotBeNull();
            opcion.Accion.Should().Be("escalar");
            opcion.CarrilesTexto.Should().Be("2,3|1,4,5");
        }

        [Fact]
        public void Resolver_AccionInvalida_LanzaArgumentOutOfRangeException()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "1,2" };

            Action act = () => AgOpcionesParticion.Resolver("LUN-JUE", ranuras, patrones, "invalida");
            act.Should().Throw<ArgumentOutOfRangeException>();
        }
    }
}
