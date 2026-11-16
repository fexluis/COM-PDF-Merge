using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Contrato de respaldo del empaquetado: cuando una asignatura pendiente no puede empaquetarse
    /// (por ejemplo, un día de patrón sin ningún hueco en el calendario), el empaquetado debe devolver
    /// false sin mutar el estado y el relleno previo debe seguir cubriendo al resto.
    /// </summary>
    public class AgPlanificadorRanurasFallbackTests
    {
        [Fact]
        public void CompletarReferencia_WhenPendingSubjectHasPatronDayWithoutSlot_ShouldFillWithGreedyWithoutHalfPack()
        {
            // La asignatura 2 es MENOR y su patrón incluye el día 3, que no existe en las ranuras.
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 8, 1),
                (2, 4, 3)
            };
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, "TARDE", 4),
                new AgRanura(2, "TARDE", 4)
            };
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 2, 9);
            var patrones = new Dictionary<int, string> { { 2, "1,2,3" } };

            var resultado = AgPlanificadorRanuras.CompletarReferencia(demandas, ranuras, ini, fin,
                4, null, patrones, null, CancellationToken.None);

            resultado.IsSuccess.Should().BeTrue();

            // La asignatura empaquetable conserva su cuota completa y sus días permitidos.
            var mayor = resultado.Value.Planes.Single(p => p.IdAsignatura == 1);
            mayor.Sesiones.Should().HaveCount(2);
            mayor.Sesiones.Should().OnlyContain(s => s.Ranura.Dia == 1 || s.Ranura.Dia == 2);

            // La restringida a un día inexistente no inventa sesiones: el respaldo no la fuerza.
            resultado.Value.Planes.Any(p => p.IdAsignatura == 2).Should().BeFalse(
                "una asignatura con día de patrón inexistente no puede empaquetarse");

            // Una sola sesión por día: el respaldo no deja huecos duplicados ni medio empaquetados.
            var todas = resultado.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            todas.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja)
                .Should().OnlyHaveUniqueItems();
            todas.Should().OnlyContain(s => s.Fecha >= ini && s.Fecha <= fin);
        }

        [Fact]
        public void CompletarReferencia_WhenPatternDayExistsInRanurasButNotInWindow_ShouldStillFillAvailableSlots()
        {
            // Día 5 existe en las ranuras pero no aparece en la ventana de dos días.
            // El empaquetado no puede colocar esa asignatura, pero el relleno previo sí cubre el día 2.
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4, 1)
            };
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, "TARDE", 4),
                new AgRanura(2, "TARDE", 4),
                new AgRanura(5, "TARDE", 4)
            };
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 2, 2);
            var patrones = new Dictionary<int, string> { { 1, "2,5" } };

            var resultado = AgPlanificadorRanuras.CompletarReferencia(demandas, ranuras, ini, fin,
                4, null, patrones, null, CancellationToken.None);

            resultado.IsSuccess.Should().BeTrue();
            resultado.Value.Planes.SelectMany(p => p.Sesiones).Should().HaveCount(1,
                "el relleno previo debe cubrir el martes disponible de la ventana");
        }

        [Fact]
        public void CompletarReferencia_WhenAllPendingSubjectsPack_ShouldKeepCountsAndAllowedDays()
        {
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 12, 1),
                (2, 8, 2)
            };
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, "TARDE", 4),
                new AgRanura(2, "TARDE", 4),
                new AgRanura(3, "TARDE", 4)
            };
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 2, 28);
            var patrones = new Dictionary<int, string> { { 2, "2,3" } };

            var resultado = AgPlanificadorRanuras.CompletarReferencia(demandas, ranuras, ini, fin,
                4, null, patrones, null, CancellationToken.None);

            resultado.IsSuccess.Should().BeTrue();
            resultado.Value.Infactibles.Should().BeEmpty();
            resultado.Value.Planes.Single(p => p.IdAsignatura == 1).Sesiones.Should().HaveCount(3);
            resultado.Value.Planes.Single(p => p.IdAsignatura == 2).Sesiones.Should().HaveCount(2)
                .And.OnlyContain(s => s.Ranura.Dia == 2 || s.Ranura.Dia == 3);

            var todas = resultado.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            todas.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja)
                .Should().OnlyHaveUniqueItems();
        }
    }
}
