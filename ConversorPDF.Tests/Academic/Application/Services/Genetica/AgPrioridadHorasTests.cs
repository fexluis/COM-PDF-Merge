using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPrioridadHorasTests
    {
        public static IEnumerable<object[]> CasosOrden()
        {
            foreach (string ruta in new[] { "oleadas", "ranuras", "completado" })
            {
                yield return new object[] { ruta, new[] { 4, 12, 8 }, new[] { 1, 9, 2 }, new[] { 2, 3, 1 } };
                yield return new object[] { ruta, new[] { 8, 8, 8 }, new[] { 3, 1, 2 }, new[] { 2, 3, 1 } };
                yield return new object[] { ruta, new[] { 8, 8, 8 }, new[] { 2, -1, 0 }, new[] { 2, 3, 1 } };
                yield return new object[] { ruta, new[] { 8, 8, 8 }, new[] { 1, 1, 1 }, new[] { 1, 2, 3 } };
            }
        }

        [Theory]
        [MemberData(nameof(CasosOrden))]
        public void Planificar_WhenDemandasDesordenadas_ShouldAsignarPorHorasYDesempates(
            string ruta, int[] horas, int[] curricular, int[] esperadas)
        {
            var demandas = Enumerable.Range(1, 3)
                .Select(id => (IdAsignatura: id, Horas: (double)horas[id - 1], OrdenCurricular: curricular[id - 1])).ToList();
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 3, 28);
            var ranuras = new List<AgRanura> { new AgRanura(1, "TARDE", 4) };
            // A smaller restricted subject must not override hours-first priority in hole filling.
            var patrones = new Dictionary<int, string> { { 1, "1" } };
            List<string> firma = null;
            foreach (var entrada in new[] { demandas, demandas.AsEnumerable().Reverse().ToList(),
                new[] { demandas[1], demandas[0], demandas[2] }.ToList() })
            {
                AgPlanOleadasResultado plan;
                if (ruta == "oleadas")
                {
                    var resultado = new AgPlanificadorOleadas().Construir(entrada, "1", ini, fin, 4);
                    resultado.IsSuccess.Should().BeTrue();
                    plan = resultado.Value;
                }
                else if (ruta == "ranuras")
                {
                    plan = AgPlanificadorRanuras.Construir(entrada, ranuras, ini, fin, 4, null, null);
                }
                else
                {
                    var resultado = AgPlanificadorRanuras.CompletarReferencia(entrada, ranuras,
                        ini, fin, 4, null, patrones, null, CancellationToken.None);
                    resultado.IsSuccess.Should().BeTrue();
                    plan = resultado.Value;
                    plan.Planes.Single(p => p.IdAsignatura == 1).Sesiones.Should()
                        .OnlyContain(s => s.Ranura.Dia == 1);
                }

                plan.Infactibles.Should().BeEmpty();
                plan.Planes.OrderBy(p => p.FechaInicio).Select(p => p.IdAsignatura).Should().Equal(esperadas);
                plan.Planes.OrderBy(p => p.Oleada).Select(p => p.IdAsignatura).Should().Equal(esperadas);
                foreach (var demanda in demandas)
                {
                    plan.Planes.Single(p => p.IdAsignatura == demanda.IdAsignatura).Sesiones
                        .Should().HaveCount((int)(demanda.Horas / 4));
                }
                var sesiones = plan.Planes.SelectMany(p => p.Sesiones).ToList();
                sesiones.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja).Should().OnlyHaveUniqueItems();
                sesiones.Should().OnlyContain(s => s.Fecha >= ini && s.Fecha <= fin);
                var actual = sesiones.OrderBy(s => s.Fecha)
                    .Select(s => s.IdAsignatura + ":" + s.Fecha.ToString("yyyyMMdd") + ":" + s.Ranura.Franja).ToList();
                if (firma != null) { actual.Should().Equal(firma); }
                firma = actual;
            }
        }

        [Fact]
        public void Construir_WhenPatronMenorCompiteConAutomatica_ShouldAsignarCarrilPorHorasAntesDeSembrar()
        {
            // Arrange: dos carriles paralelos; el lunes final permite la cola del primero.
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4, 1), (2, 8, 9), (3, 8, 10)
            };
            var ini = new DateTime(2027, 2, 1);
            var fin = ini.AddDays(7);
            var patrones = new Dictionary<int, string> { { 1, "1,2" } };
            AgPlanOleadasResultado anterior = null;
            foreach (var entrada in new[] { demandas, demandas.AsEnumerable().Reverse().ToList(),
                new[] { demandas[1], demandas[0], demandas[2] }.ToList() })
            {
                // Act
                var resultado = new AgPlanificadorOleadas().Construir(entrada, "1,2,3,4",
                    ini, fin, 4, diasPatronPorAsignatura: patrones);
                var evaluacion = new AgPlanificadorOleadas().ConstruirEvaluacion(entrada, "1,2,3,4",
                    ini, fin, 4, diasPatronPorAsignatura: patrones);

                // Assert: la automática mayor ocupa primero el carril 0; la menor espera en él.
                resultado.IsSuccess.Should().BeTrue();
                evaluacion.IsSuccess.Should().BeTrue();
                resultado.Value.Infactibles.Should().BeEmpty();
                var mayor = resultado.Value.Planes.Single(p => p.IdAsignatura == 2);
                var menor = resultado.Value.Planes.Single(p => p.IdAsignatura == 1);
                mayor.Carril.Should().Be(0);
                mayor.Oleada.Should().Be(0);
                mayor.Sesiones.Select(s => s.Fecha).Should().Equal(ini, ini.AddDays(1));
                menor.Carril.Should().Be(0);
                menor.Oleada.Should().Be(1);
                menor.Sesiones.Select(s => s.Fecha).Should().Equal(fin);
                menor.Sesiones.Should().OnlyContain(s => s.Ranura.Dia == 1 || s.Ranura.Dia == 2);
                var paralela = resultado.Value.Planes.Single(p => p.IdAsignatura == 3);
                paralela.Carril.Should().Be(1);
                paralela.Oleada.Should().Be(0);
                paralela.Sesiones.Select(s => s.Fecha).Should().Equal(ini.AddDays(2), ini.AddDays(3));
                paralela.FechaInicio.Should().BeBefore(menor.FechaInicio);
                var sesiones = resultado.Value.Planes.SelectMany(p => p.Sesiones).ToList();
                sesiones.Should().HaveCount(5);
                sesiones.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja)
                    .Should().OnlyHaveUniqueItems();
                resultado.Value.Planes.Should().BeEquivalentTo(evaluacion.Value.Planes);
                if (anterior != null) { resultado.Value.Planes.Should().BeEquivalentTo(anterior.Planes); }
                anterior = resultado.Value;
            }
            demandas.Select(d => d.Horas).Should().Equal(4, 8, 8);
        }

        [Fact]
        public void CompletarReferencia_WhenMayorCargaPendiente_ShouldPreservarReferenciaYRestricciones()
        {
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4, 1), (2, 12, 3), (3, 8, 2)
            };
            var lunes = new AgRanura(1, "TARDE", 4);
            var martes = new AgRanura(2, "TARDE", 4);
            var ranuras = new List<AgRanura> { lunes, martes };
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 2, 23);
            var protegidas = new List<AgSesionPlanificada> { new AgSesionPlanificada(1, ini, lunes) };
            var referencia = new AgPlanOleadasResultado(new List<AgPlanOleada>
            {
                new AgPlanOleada(1, 0, 0, new[] { 1 }, ini, ini, 1, 1, 1,
                    new[] { lunes }, protegidas)
            }, new List<int> { 2, 3 });
            var patrones = new Dictionary<int, string> { { 1, "1" }, { 3, "1" } };

            var resultado = AgPlanificadorRanuras.CompletarReferencia(demandas, ranuras, ini, fin,
                4, null, patrones, referencia, CancellationToken.None);

            resultado.IsSuccess.Should().BeTrue();
            resultado.Value.Infactibles.Should().BeEmpty();
            resultado.Value.Planes.Single(p => p.IdAsignatura == 1).Sesiones.Should().BeEquivalentTo(protegidas);
            resultado.Value.Planes.Single(p => p.IdAsignatura == 2).Sesiones.Select(s => s.Fecha)
                .Should().Equal(ini.AddDays(1), ini.AddDays(7), ini.AddDays(8));
            resultado.Value.Planes.Single(p => p.IdAsignatura == 3).Sesiones.Select(s => s.Fecha)
                .Should().Equal(ini.AddDays(14), ini.AddDays(21));
            referencia.Planes.Single().Sesiones.Should().BeEquivalentTo(protegidas);
            referencia.Infactibles.Should().Equal(2, 3);
        }

        [Fact]
        public void Demandas_WhenHorasObjetivoDifierenDelCatalogo_ShouldOrdenarSoloSemestreYAulaActuales()
        {
            var contexto = new AgContexto();
            var aula = new AgAula { Id = 1, IdPrograma = 1, Semestre = "III" };
            contexto.Aulas.Add(aula);
            contexto.Aulas.Add(new AgAula { Id = 2, IdPrograma = 1, Semestre = "III" });
            for (int id = 1; id <= 4; id++)
            {
                contexto.Asignaturas.Add(new Asignatura
                {
                    Id = id, IdPrograma = 1, Semestre = id == 4 ? "IV" : "III", IntHoraria = 100 - id * 4
                });
                contexto.AulaAsignaturas.Add(new AgAulaAsignatura
                {
                    IdAula = 1, IdAsignatura = id, HorasObjetivo = id * 4, OrdenCurricular = id
                });
            }
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 2, IdAsignatura = 1, HorasObjetivo = 200, OrdenCurricular = 1
            });
            contexto.ConstruirIndices();

            var demandas = AgEvaluadorDemandaVentana.Demandas(contexto, aula);

            demandas.Select(d => d.IdAsignatura).Should().Equal(3, 2, 1);
            demandas.Select(d => d.Horas).Should().Equal(12, 8, 4);
            contexto.AulaAsignaturas.Should().HaveCount(5);
        }
    }
}
