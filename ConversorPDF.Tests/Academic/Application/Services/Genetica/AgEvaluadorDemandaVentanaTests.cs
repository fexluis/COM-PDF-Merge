using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgEvaluadorDemandaVentanaTests
    {
        [Fact]
        public void Evaluar_WhenDemandaExcedeCapacidadGlobal_ShouldIndicarSaturacion()
        {
            var contexto = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2026, 2, 1),
                    FechaFin = new DateTime(2026, 5, 30) // 17 semanas
                }
            };

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-01",
                Estado = AgEstadoAula.Disponible,
                DiasVentana = "5,6", // VIE-SAB (17 vie + 34 sab = 51 ranuras)
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            // 4 asignaturas de 64h = 4 * 16 = 64 sesiones > 51 ranuras
            for (int i = 1; i <= 4; i++)
            {
                var asig = new Asignatura { Id = i, Nombre = $"Materia {i}", IntHoraria = 64 };
                contexto.Asignaturas.Add(asig);
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = 64.0
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }
            contexto.ConstruirIndices();

            var eval = AgEvaluadorDemandaVentana.Evaluar(
                aula.DiasVentana, aula.Jornada, 4.0, contexto, aula,
                contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);

            string sesInfo = string.Join("; ", (eval.Sesiones ?? new List<AgSesionPlanificada>()).GroupBy(s => s.IdAsignatura).Select(g => $"Asig{g.Key}:{g.Count()}(dias:{string.Join(",", g.Select(x => x.Ranura.Dia).Distinct())})"));
            eval.HaySaturacion.Should().BeTrue($"sesInfo={sesInfo}, cabe={eval.CabeEnPlazo}, pend={eval.DemandaPendiente}");
            eval.DemandaTotal.Should().Be(64);
            eval.CapacidadTotal.Should().Be(51);
        }

        [Fact]
        public void Evaluar_WhenDemandaCabeGlobalPeroSaturaCarrilDomingo_ShouldIndicarSaturacionCarriles()
        {
            var contexto = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2026, 2, 1),
                    FechaFin = new DateTime(2026, 5, 30) // 17 semanas
                }
            };

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-01",
                Estado = AgEstadoAula.Disponible,
                DiasVentana = "5,6,7", // VIE-DOM (51 vie-sab + 17 dom = 68 ranuras)
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            // 6 materias con sesiones [12, 12, 8, 8, 12, 12] = 64 sesiones (256h)
            // Sembradas alternadamente por GenerarEncabezadosAulasAsync (banda 0=[5,6], banda 1=[7]):
            // Las asignaturas impares reciben patrón "7" (Domingo) sumando 12 + 8 + 12 = 32 sesiones.
            // Capacidad domingo = 17 ranuras. Demanda domingo (32) > 17 produce saturación de carril.
            int[] horas = { 48, 48, 32, 32, 48, 48 };
            for (int i = 0; i < horas.Length; i++)
            {
                int id = i + 1;
                var asig = new Asignatura { Id = id, Nombre = $"Materia {id}", IntHoraria = horas[i] };
                contexto.Asignaturas.Add(asig);
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = id,
                    HorasObjetivo = horas[i],
                    DiasPatron = (i % 2 == 1) ? "7" : "5,6",
                    OrigenPatron = "FIJO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, id)] = aa;
            }
            contexto.ConstruirIndices();

            var eval = AgEvaluadorDemandaVentana.Evaluar(
                aula.DiasVentana, aula.Jornada, 4.0, contexto, aula,
                contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);

            eval.HaySaturacion.Should().BeTrue();
            eval.SaturacionCarriles.Should().BeTrue();
            eval.DemandaTotal.Should().Be(64);
            eval.CapacidadTotal.Should().Be(68);

            // Escalador no debe proponer JUE-DOM: no reduce la cola dominical y rompe el patrón 5,6
            var escalador = new AgEscaladorRanuraAula(() => null);
            var decisiones = escalador.Evaluar(contexto);

            decisiones.Should().BeEmpty("JUE-DOM no reduce la saturación dominical ni conserva el patrón 5,6");
        }
    }
}
