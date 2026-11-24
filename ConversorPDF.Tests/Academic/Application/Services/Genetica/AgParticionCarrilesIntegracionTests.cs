using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Pruebas de integración del selector de partición de carriles en AgPlanificadorOleadas (Fase 1581).
    /// Cubre los casos de la especificación §7.3 - §7.8.
    /// </summary>
    public class AgParticionCarrilesIntegracionTests
    {
        private readonly AgPlanificadorOleadas _sut = new AgPlanificadorOleadas();

        [Fact]
        public void Caso73_Integracion_CasoHomogeneoLogReal_DebePlanificarSinInfactiblesYConBanda67()
        {
            // '4,5,6,7' · 4 h · 2027-02-05..2027-05-30 (W = 17) · asigs 38,40,42 de 32 h · patrón '6,7'
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (38, 32.0, 1),
                (40, 32.0, 2),
                (42, 32.0, 3)
            };
            var patrones = new Dictionary<int, string>
            {
                { 38, "6,7" },
                { 40, "6,7" },
                { 42, "6,7" }
            };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            var val = res.Value;
            val.Infactibles.Should().BeEmpty();
            val.Planes.Should().HaveCount(3);

            // Debe existir una banda que contenga exactamente o aloje a {6,7}
            val.Particion.Should().Contain(banda => banda.Contains(6) && banda.Contains(7));

            // Los planes no deben solaparse en fechas dentro del mismo carril
            foreach (var g in val.Planes.GroupBy(p => p.Carril))
            {
                var ordenados = g.OrderBy(p => p.Oleada).ToList();
                for (int i = 1; i < ordenados.Count; i++)
                {
                    ordenados[i].FechaInicio.Should().BeAfter(ordenados[i - 1].FechaFin);
                }
            }

            foreach (var p in val.Planes)
            {
                p.FechaInicio.Should().BeOnOrAfter(fInicio);
                p.FechaFin.Should().BeOnOrBefore(fFin);
            }
        }

        [Fact]
        public void Caso74_Integracion_ConvivenciaDePatrones_DebeAceptarCandidata45Y67()
        {
            // '4,5,6,7' · 4 h · ventana amplia · patrones '4,5' ; '6,7' ; '6,7'
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 32.0, 3)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "4,5" },
                { 2, "6,7" },
                { 3, "6,7" }
            };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            var val = res.Value;
            val.Infactibles.Should().BeEmpty();
            val.Planes.Should().HaveCount(3);

            // La partición elegida debe tener las bandas {4,5} y {6,7}
            val.Particion.Should().HaveCount(2);
            val.Particion[0].Should().Equal(4, 5);
            val.Particion[1].Should().Equal(6, 7);

            // La asignatura 1 debe estar en el carril de {4,5} y las 2 y 3 en el de {6,7}
            val.Planes.Single(p => p.IdAsignatura == 1).Carril.Should().Be(0);
            val.Planes.Where(p => p.IdAsignatura == 2 || p.IdAsignatura == 3)
                .All(p => p.Carril == 1).Should().BeTrue();
        }

        [Fact]
        public void Caso75_Integracion_MasBandasPuedeEmpeorar_DebeElegir2BandasSinPerderPlanificadas()
        {
            // '4,5,6,7' · 4 h · 2027-03-01..2027-03-21 (W = 3)
            // A 16 h sin patrón (4 sesiones)
            // B 16 h sin patrón (4 sesiones)
            // C 4 h patrón '6,7' (1 sesión)
            var fInicio = new DateTime(2027, 3, 1);
            var fFin = new DateTime(2027, 3, 21);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1),
                (2, 16.0, 2),
                (3, 4.0, 3)
            };
            var patrones = new Dictionary<int, string>
            {
                { 3, "6,7" }
            };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            var val = res.Value;
            // Debe lograr 0 infactibles eligiendo la partición de 2 bandas {4,5}{6,7}
            val.Infactibles.Should().BeEmpty();
            val.Planes.Should().HaveCount(3);
            val.Particion.Should().HaveCount(2);
        }

        [Fact]
        public void Caso76_Integracion_MejoraParcial_DebeReducirInfactiblesYConservarFactible()
        {
            // F1590: Una semana (W=1); dos asignaturas de 8h (2 sesiones c/u = demanda 4); ambas patrón '6,7'.
            // Capacidad global = 5 sesiones. Banda {6,7} tiene 3 sesiones hábiles; alcanza para una sola de 8h (2 sesiones).
            // La partición alternativa reduce infactibles de 2 a 1 (planifica una asignatura y deja una infactible).
            var fInicio = new DateTime(2027, 3, 1);
            var fFin = new DateTime(2027, 3, 7);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 8.0, 1),
                (2, 8.0, 2)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "6,7" },
                { 2, "6,7" }
            };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Planes.Should().ContainSingle();
            res.Value.Infactibles.Should().ContainSingle();
            // Fase 2395 / Fase 2396: Por desempate determinista sin singletons y menor número de bandas, se elige {{4,5},{6,7}} (2 bandas)
            res.Value.Particion.Should().HaveCount(2);
            res.Value.Particion[0].Should().Equal(4, 5);
            res.Value.Particion[1].Should().Equal(6, 7);
        }

        [Fact]
        public void Caso76_Integracion_AusenciaDeMejora_DebeConservarReferencia()
        {
            // F1590: Una materia de 16h (demanda 4 sesiones), patrón '6,7' en una semana (W=1, capacidad global 5).
            // Banda fiel {6,7} dispone de solo 3 sesiones; ninguna candidata fiel puede completarla.
            // Infactibles no pueden reducirse => debe conservar la referencia canónica íntegra.
            var fInicio = new DateTime(2027, 3, 1);
            var fFin = new DateTime(2027, 3, 7);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "6,7" }
            };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Planes.Should().BeEmpty();
            res.Value.Infactibles.Should().Equal(1);
            res.Value.Particion.Should().HaveCount(3);
            res.Value.Particion[0].Should().Equal(4, 5);
            res.Value.Particion[1].Should().Equal(6);
            res.Value.Particion[2].Should().Equal(7);
        }

        [Fact]
        public void Caso78_Integracion_FestivoEnFinDeSemana_DebeExtenderSemanasOleada()
        {
            // F1598: Con 24h (6 sesiones requeridas) y festivo el sábado 2027-02-06:
            // En semana 1 hay solo 1 sesión (domingo 7); en semanas 2 y 3 hay 3 sesiones c/u (sábado M+T, domingo M).
            // Con festivo: se requieren 3 semanas (1 + 3 + 2 = 6 sesiones) y FechaFin llega a 2027-02-25.
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var sabadoFestivo = new DateTime(2027, 2, 6);
            Func<DateTime, bool> esFestivo = d => d.Date == sabadoFestivo;

            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 24.0, 1)
            };
            var patrones = new Dictionary<int, string> { { 1, "6,7" } };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0,
                esFestivo: esFestivo, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();

            var plan = res.Value.Planes.Single();
            plan.SesionesRequeridas.Should().Be(6);
            plan.SemanasOleada.Should().Be(3);
            plan.FechaInicio.Should().Be(new DateTime(2027, 2, 5));
            plan.FechaFin.Should().Be(new DateTime(2027, 2, 25));
        }

        [Fact]
        public void Caso78_Integracion_SinFestivoEnFinDeSemana_DebeRequerirMenosSemanas()
        {
            // F1598: Control sin festivo: 24h (6 sesiones) en banda {6,7} a 3 sesiones/semana
            // completa en exactamente 2 semanas (3 + 3 = 6 sesiones), FechaFin = 2027-02-18.
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);

            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 24.0, 1)
            };
            var patrones = new Dictionary<int, string> { { 1, "6,7" } };

            var res = _sut.Construir(asignaturas, "4,5,6,7", fInicio, fFin, 4.0,
                esFestivo: null, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();

            var plan = res.Value.Planes.Single();
            plan.SesionesRequeridas.Should().Be(6);
            plan.SemanasOleada.Should().Be(2);
            plan.FechaInicio.Should().Be(new DateTime(2027, 2, 5));
            plan.FechaFin.Should().Be(new DateTime(2027, 2, 18));
        }
    }
}

