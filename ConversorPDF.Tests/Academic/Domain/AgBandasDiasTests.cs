using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgBandasDiasTests
    {
        private static IReadOnlyList<AgRanura> Ranuras(int[] dias, double horas = 4.0)
            => dias.SelectMany(d => d == 6
                    ? new[] { new AgRanura(6, AgFranjaJornada.MANANA, horas), new AgRanura(6, AgFranjaJornada.TARDE, horas) }
                    : new[] { new AgRanura(d, AgFranjaJornada.MANANA, horas) })
                .ToList();

        [Fact]
        public void Particiones_CuandoVentanaDe3Dias_DevuelveCuatroParticionesContiguas()
        {
            var partes = AgBandasDias.Particiones(new[] { 4, 5, 6 });

            partes.Should().HaveCount(4);                       // 2^(3-1)
            partes.Should().OnlyContain(p => p.SelectMany(b => b).Distinct().Count() == 3);
            partes.SelectMany(p => p).Should().OnlyContain(b => b.Count > 0);
        }

        [Fact]
        public void Particiones_CuandoVentanaVacia_DevuelveVacio()
        {
            AgBandasDias.Particiones(new int[0]).Should().BeEmpty();
        }

        [Fact]
        public void RanurasDe_CuandoVieSab_ReparteTresRanurasEnDosBandas()
        {
            var ranuras = Ranuras(new[] { 5, 6 });

            AgBandasDias.RanurasDe(ranuras, new[] { 5 }).Should().HaveCount(1);
            AgBandasDias.RanurasDe(ranuras, new[] { 6 }).Should().HaveCount(2);
        }

        [Fact]
        public void Resolver_CuandoCapacidadSuficiente_MaximizaBandasYMakespan()
        {
            var ranuras = Ranuras(new[] { 2, 3, 4, 5 });
            var demanda = new[] { 16, 16, 16, 16 };

            var asign = AgBandasDias.Resolver(new[] { 2, 3, 4, 5 }, ranuras, demanda, null, 16);

            asign.Factible.Should().BeTrue();
            asign.Bandas.Should().HaveCount(4);                 // paralelismo máximo, desperdicio 0
            asign.Makespan.Should().Be(16);
            asign.Ocupacion.Should().BeApproximately(1.0, 0.001);
            asign.BandaPorIndice.Distinct().Should().HaveCount(4);   // determinista, orden interno irrelevante
        }

        [Fact]
        public void Resolver_CuandoDemandaExcedeCapacidad_ReportaExcesoPorBanda()
        {
            var ranuras = Ranuras(new[] { 5, 6 });               // 3 ranuras/semana
            var asign = AgBandasDias.Resolver(new[] { 5, 6 }, ranuras, new[] { 6, 6 }, null, 2);

            asign.Factible.Should().BeFalse();
            asign.ExcesoPorBanda.Sum().Should().Be(6);           // 12 sesiones vs 6 de capacidad
        }

        [Fact]
        public void Resolver_CuandoPatronCruzaBandas_RespetaPatronEnBandaAlojadora()
        {
            var ranuras = Ranuras(new[] { 4, 5, 6 });
            var demanda = new[] { 3, 3 };

            var asign = AgBandasDias.Resolver(new[] { 4, 5, 6 }, ranuras, demanda, new[] { "4,6", "4,6" }, 12);

            asign.Bandas.Should().HaveCount(1);
            asign.ViolacionesPatron.Should().Be(0);             // el patrón manda sobre el paralelismo
        }
    }
}
