using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgVentanaSolapeTests
    {
        [Fact]
        public void SeSolapan_FechasDisjuntas_DevuelveFalse()
        {
            var iniA = new DateTime(2027, 2, 2);
            var finA = new DateTime(2027, 4, 5);
            var iniB = new DateTime(2027, 4, 6);
            var finB = new DateTime(2027, 4, 19);

            bool solapan = AgVentanaSolape.SeSolapan(iniA, finA, iniB, finB);

            solapan.Should().BeFalse();
        }

        [Fact]
        public void SeSolapan_BordeMismoDia_DevuelveTrue()
        {
            var iniA = new DateTime(2027, 2, 2);
            var finA = new DateTime(2027, 4, 5);
            var iniB = new DateTime(2027, 4, 5);
            var finB = new DateTime(2027, 4, 19);

            bool solapan = AgVentanaSolape.SeSolapan(iniA, finA, iniB, finB);

            solapan.Should().BeTrue();
        }

        [Fact]
        public void SeSolapan_FechasNulas_DevuelveTrueConservador()
        {
            var iniA = new DateTime(2027, 2, 2);
            var finA = new DateTime(2027, 4, 5);

            AgVentanaSolape.SeSolapan(iniA, finA, null, null).Should().BeTrue();
            AgVentanaSolape.SeSolapan(null, null, iniA, finA).Should().BeTrue();
            AgVentanaSolape.SeSolapan(null, null, null, null).Should().BeTrue();
        }

        private class ElementoConVentana
        {
            public string Nombre { get; set; }
            public DateTime? Ini { get; set; }
            public DateTime? Fin { get; set; }
        }

        [Fact]
        public void AgruparPorSolape_TresOleadasConsecutivas_ProduceTresGruposDeUno()
        {
            var oleadas = new List<ElementoConVentana>
            {
                new ElementoConVentana { Nombre = "O1", Ini = new DateTime(2027, 2, 2), Fin = new DateTime(2027, 3, 15) },
                new ElementoConVentana { Nombre = "O2", Ini = new DateTime(2027, 3, 16), Fin = new DateTime(2027, 4, 15) },
                new ElementoConVentana { Nombre = "O3", Ini = new DateTime(2027, 4, 16), Fin = new DateTime(2027, 5, 15) }
            };

            var grupos = AgVentanaSolape.AgruparPorSolape(oleadas, x => x.Ini, x => x.Fin);

            grupos.Should().HaveCount(3);
            grupos[0].Should().HaveCount(1);
            grupos[1].Should().HaveCount(1);
            grupos[2].Should().HaveCount(1);
        }

        [Fact]
        public void AgruparPorSolape_DosSolapadas_ProduceUnGrupoDeDos()
        {
            var elementos = new List<ElementoConVentana>
            {
                new ElementoConVentana { Nombre = "A1", Ini = new DateTime(2027, 2, 2), Fin = new DateTime(2027, 3, 20) },
                new ElementoConVentana { Nombre = "A2", Ini = new DateTime(2027, 3, 10), Fin = new DateTime(2027, 4, 15) }
            };

            var grupos = AgVentanaSolape.AgruparPorSolape(elementos, x => x.Ini, x => x.Fin);

            grupos.Should().HaveCount(1);
            grupos[0].Should().HaveCount(2);
        }
    }
}

