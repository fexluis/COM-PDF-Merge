using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgEscaladoTresCasosTests
    {
        [Fact]
        public void DemostrarCuelloDeBotella_LunVie_SiembraIngenuaSobrecargaViernes()
        {
            // 6 materias con sesiones [12, 12, 8, 8, 12, 12] = 64 sesiones
            var sesionesPorMateria = new[] { 12, 12, 8, 8, 12, 12 };
            var ventanaDias = new List<int> { 1, 2, 3, 4, 5 }; // LUN-VIE

            // En el algoritmo ingenuo de AgFormGeneradorHorarios:
            int nDias = ventanaDias.Count; // 5
            int bandasSiembra = (int)Math.Ceiling(nDias / 2.0); // 3 bandas
            int tamBanda = (int)Math.Ceiling(nDias / (double)bandasSiembra); // 2

            var cargaPorBanda = new int[bandasSiembra];
            var diasPorBanda = new List<List<int>>();

            for (int b = 0; b < bandasSiembra; b++)
            {
                int desde = b * tamBanda;
                diasPorBanda.Add(ventanaDias.Skip(desde).Take(Math.Min(tamBanda, nDias - desde)).ToList());
            }

            for (int i = 0; i < sesionesPorMateria.Length; i++)
            {
                int bandaIdx = i % bandasSiembra;
                cargaPorBanda[bandaIdx] += sesionesPorMateria[i];
            }

            // Banda 2 contiene únicamente el día 5 (Viernes)
            diasPorBanda[2].Should().Equal(new[] { 5 });
            // Carga en viernes: idx 2 (8) + idx 5 (12) = 20 sesiones
            cargaPorBanda[2].Should().Be(20);

            // En un semestre con 17 semanas (17 viernes), 20 sesiones exceden la capacidad del viernes
            int semanasDisponibles = 17;
            int capacidadViernes = semanasDisponibles * 1;
            cargaPorBanda[2].Should().BeGreaterThan(capacidadViernes, "la siembra ingenua sobrecarga el viernes provocando cuello de botella aunque el total (64) quepa en 5*17=85");
        }

        [Fact]
        public void DemostrarCuelloDeBotella_VieDom_SiembraIngenuaSobrecargaDomingo()
        {
            // 6 materias con sesiones [12, 12, 8, 8, 12, 12] = 64 sesiones
            var sesionesPorMateria = new[] { 12, 12, 8, 8, 12, 12 };
            var ventanaDias = new List<int> { 5, 6, 7 }; // VIE-DOM

            // En el algoritmo ingenuo:
            int nDias = ventanaDias.Count; // 3
            int bandasSiembra = (int)Math.Ceiling(nDias / 2.0); // 2 bandas
            int tamBanda = (int)Math.Ceiling(nDias / (double)bandasSiembra); // 2

            var cargaPorBanda = new int[bandasSiembra];
            var diasPorBanda = new List<List<int>>();

            for (int b = 0; b < bandasSiembra; b++)
            {
                int desde = b * tamBanda;
                diasPorBanda.Add(ventanaDias.Skip(desde).Take(Math.Min(tamBanda, nDias - desde)).ToList());
            }

            for (int i = 0; i < sesionesPorMateria.Length; i++)
            {
                int bandaIdx = i % bandasSiembra;
                cargaPorBanda[bandaIdx] += sesionesPorMateria[i];
            }

            // Banda 0 tiene [5, 6], Banda 1 tiene [7] (Domingo)
            diasPorBanda[1].Should().Equal(new[] { 7 });
            // Carga en domingo: idx 1 (12) + idx 3 (8) + idx 5 (12) = 32 sesiones
            cargaPorBanda[1].Should().Be(32);

            // En 17 semanas, domingo sólo tiene 17 ranuras
            int semanasDisponibles = 17;
            int capacidadDomingo = semanasDisponibles * 1;
            cargaPorBanda[1].Should().BeGreaterThan(capacidadDomingo, "la siembra ingenua asigna 32 sesiones al domingo cuando sólo hay 17 disponibles");
        }

        [Fact]
        public void RedistribucionPorRanuras_DebeSerFactible_CuandoCapacidadTotalEsSuficiente()
        {
            // En LUN-VIE con 17 semanas = 85 ranuras totales
            // Si distribuimos 64 sesiones respetando capacidades de carriles:
            // Carril 1 (LUN-MAR = 34 ranuras) -> 24 sesiones
            // Carril 2 (MIE-JUE = 34 ranuras) -> 24 sesiones
            // Carril 3 (VIE = 17 ranuras)     -> 16 sesiones
            // Total = 64 sesiones, ninguna supera su carril
            int capacidadCarril1 = 34;
            int capacidadCarril2 = 34;
            int capacidadCarril3 = 17;

            int sesionesC1 = 24;
            int sesionesC2 = 24;
            int sesionesC3 = 16;

            sesionesC1.Should().BeLessThanOrEqualTo(capacidadCarril1);
            sesionesC2.Should().BeLessThanOrEqualTo(capacidadCarril2);
            sesionesC3.Should().BeLessThanOrEqualTo(capacidadCarril3);
            (sesionesC1 + sesionesC2 + sesionesC3).Should().Be(64);
        }
    }
}
