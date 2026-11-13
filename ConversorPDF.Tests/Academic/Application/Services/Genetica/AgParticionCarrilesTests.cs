using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgParticionCarrilesTests
    {
        [Fact]
        public void Seleccionar_WhenVentanaORanurasVacias_ShouldThrowArgumentException()
        {
            var canonicas = new List<IReadOnlyList<AgRanura>>();
            var evaluadorLlamado = false;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                evaluadorLlamado = true;
                return CrearEvaluacionSintetica(bandas, 0, 0, new List<int>(), new List<int>());
            };

            Action act = () => AgParticionCarriles.Seleccionar(
                new List<int>(),
                new List<AgRanura>(),
                new List<string> { "6,7" },
                canonicas,
                evaluar);

            act.Should().Throw<ArgumentException>();
            evaluadorLlamado.Should().BeFalse("no debe invocarse el evaluador ante entradas inválidas");
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("0")]
        [InlineData("9,9")]
        public void Seleccionar_WhenPatronesNoValidos_ShouldReturnSinCambio(string patronInvalido)
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>> { ranuras };
            var patrones = patronInvalido == null ? null : new List<string> { patronInvalido };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5 },
                ranuras,
                patrones,
                canonicas,
                bandas => CrearEvaluacionSintetica(bandas, 0, 0, new List<int> { 1 }, new List<int>()));

            res.Cambiada.Should().BeFalse();
            res.Referencia.Should().BeSameAs(res.Elegida);
        }

        [Fact]
        public void Seleccionar_WhenNoAlojablesRefEsCeroPeroSinInfactibles_ShouldReturnSinCambio()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>> { ranuras };
            var patrones = new List<string> { "4,5" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5 },
                ranuras,
                patrones,
                canonicas,
                bandas => CrearEvaluacionSintetica(bandas, 0, 0, new List<int> { 1 }, new List<int>()));

            res.Cambiada.Should().BeFalse();
            res.Referencia.Should().BeSameAs(res.Elegida);
        }

        [Fact]
        public void Seleccionar_WhenMejoraEstructuralPeroIgualNumeroInfactibles_ShouldRechazar()
        {
            // Ventana 4,5,6,7. Patrón "6,7".
            // Canónica: [4,5], [6,7] no: canónica [4,5]|[6]|[7] tiene noAloj=1.
            // Alternativa: [4,5]|[6,7] tiene noAloj=0 pero evaluación da los mismos infactibles.
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] },
                new List<AgRanura> { ranuras[3] }
            };
            var patrones = new List<string> { "6,7" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    // Si es canónica: noAloj=1, infactibles=1 (asig 1)
                    if (bandas.Count == 3)
                        return CrearEvaluacionSintetica(bandas, 1, 1, new List<int>(), new List<int> { 1 });
                    // Candidata alternativa: noAloj=0, pero sigue dando infactibles=1
                    return CrearEvaluacionSintetica(bandas, 0, 1, new List<int>(), new List<int> { 1 });
                });

            res.Cambiada.Should().BeFalse();
            res.Referencia.Should().BeSameAs(res.Elegida);
        }

        [Fact]
        public void Seleccionar_WhenMenosInfactiblesPeroPierdeUnaPlanificada_ShouldRechazar()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] },
                new List<AgRanura> { ranuras[3] }
            };
            var patrones = new List<string> { "6,7" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    if (bandas.Count == 3)
                    {
                        // Referencia tiene planificada la asig 1 y 2
                        return CrearEvaluacionSintetica(bandas, 1, 2, new List<int> { 1, 2 }, new List<int> { 3, 4 });
                    }
                    // Alternativa baja infactibles a 1, pero solo planifica {2, 3} (pierde la asig 1)
                    return CrearEvaluacionSintetica(bandas, 0, 1, new List<int> { 2, 3 }, new List<int> { 4 });
                });

            res.Cambiada.Should().BeFalse();
            res.Referencia.Should().BeSameAs(res.Elegida);
        }

        [Fact]
        public void Seleccionar_WhenPierdeUnAlojablePrevio_ShouldRechazar()
        {
            // Patrones: 1:"4,5", 2:"6,7".
            // Canónica aloja "4,5". Si una partición alternativa extraña alojara "6,7" pero rompiera "4,5"
            // debe ser descartada por condición (b).
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] },
                new List<AgRanura> { ranuras[3] }
            };
            var patrones = new List<string> { "4,5", "6,7" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    // Si la banda rompe 4,5 (ej. {4}|{5,6,7}), noAlojables podría variar pero no debe admitirse
                    int noAloj = AgParticionCarriles.ContarPatronesNoAlojables(bandas, patrones);
                    int infact = noAloj > 0 ? 1 : 0;
                    var alojables = AgParticionCarriles.ObtenerIndicesAlojables(bandas, patrones);
                    return CrearEvaluacionSintetica(bandas, noAloj, infact, new List<int> { 1, 2 }, new List<int>(), patronesAlojables: alojables);
                });

            // Con {4,5}|{6,7} preserva 4,5 y aloja 6,7 (noAloj=0), por lo que esa sí gana.
            res.Cambiada.Should().BeTrue();
            res.Elegida.PatronesNoAlojables.Should().Be(0);
        }

        [Fact]
        public void Seleccionar_WhenDosCandidatasValidas_ShouldSeguirOrdenDesempateExacto()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] },
                new List<AgRanura> { ranuras[3] }
            };
            var patrones = new List<string> { "6,7" };

            // Dos particiones que alojan 6,7:
            // P1: {4}{5}{6,7} (3 bandas)
            // P2: {4,5}{6,7} (2 bandas)
            // Si ambas dan 0 infactibles, 0 noAlojables, mismo makespan=9:
            // Desempate: ThenByDescending(Bandas) -> gana P1 con 3 bandas!
            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    int noAloj = AgParticionCarriles.ContarPatronesNoAlojables(bandas, patrones);
                    int infact = noAloj > 0 ? 1 : 0;
                    var alojables = AgParticionCarriles.ObtenerIndicesAlojables(bandas, patrones);
                    return CrearEvaluacionSintetica(bandas, noAloj, infact, new List<int> { 1 }, new List<int>(), makespan: 9, patronesAlojables: alojables);
                });

            res.Cambiada.Should().BeTrue();
            res.Elegida.Bandas.Should().Be(3);
            res.CandidatasEvaluadas.Should().BeGreaterThan(0, "debe contar las evaluaciones reales");
        }

        [Fact]
        public void Seleccionar_ConSabadoMananaYTarde_ConservaAmbasRanuras()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.MANANA, 4.0),
                new AgRanura(6, AgFranjaJornada.TARDE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] }
            };
            var patrones = new List<string> { "6" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 5, 6 },
                ranuras,
                patrones,
                canonicas,
                bandas => CrearEvaluacionSintetica(bandas, 0, 0, new List<int> { 1 }, new List<int>()));

            // La canónica ya aloja 6 en ambos carriles (noAloj=0) => sin cambio
            res.Cambiada.Should().BeFalse();
        }

        [Fact]
        public void Seleccionar_WhenIgualInfactiblesPeroMejorMakespan_ShouldAceptar()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2] },
                new List<AgRanura> { ranuras[3] }
            };
            var patrones = new List<string> { "6,7" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    if (bandas.Count == 3)
                        return CrearEvaluacionSintetica(bandas, 1, 1, new List<int> { 1 }, new List<int> { 2 }, makespan: 5);
                    // Alternativa: igual número de infactibles (1) pero mejor makespan (3)
                    return CrearEvaluacionSintetica(bandas, 0, 1, new List<int> { 1 }, new List<int> { 2 }, makespan: 3);
                });

            res.Cambiada.Should().BeTrue();
            res.Elegida.Makespan.Should().Be(3);
        }

        [Fact]
        public void Seleccionar_WhenSinPatronesValidosPeroConInfactibles_ShouldEvaluarParticionesAlternativas()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] },
                new List<AgRanura> { ranuras[4] }
            };
            // Sin patrones explícitos (todos automáticos)
            var patrones = new List<string> { "", "" };

            var res = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                bandas =>
                {
                    // Canónica tiene 1 infactible por cuello de botella
                    if (bandas.Count == 3 && bandas[2].Count == 1)
                        return CrearEvaluacionSintetica(bandas, 0, 1, new List<int> { 1 }, new List<int> { 2 });
                    // Alternativa balanceada resuelve el infactible
                    return CrearEvaluacionSintetica(bandas, 0, 0, new List<int> { 1, 2 }, new List<int>());
                });

            res.Cambiada.Should().BeTrue();
            res.Elegida.Resultado.Infactibles.Count.Should().Be(0);
        }

        private static AgEvaluacionParticion CrearEvaluacionSintetica(
            IReadOnlyList<IReadOnlyList<AgRanura>> bandas,
            int patronesNoAlojables,
            int infactiblesCount,
            List<int> idsPlanes,
            List<int> idsInfactibles,
            int makespan = 5,
            IReadOnlyCollection<int> patronesAlojables = null)
        {
            var planes = idsPlanes.Select(id => new AgPlanOleada(
                id, 0, 0, new List<int> { 4 }, DateTime.Today, DateTime.Today.AddDays(7), 1, 1, 2, new List<AgRanura>()
            )).ToList();

            var infactList = idsInfactibles.ToList();
            if (infactList.Count == 0 && infactiblesCount > 0)
            {
                infactList = Enumerable.Range(100, infactiblesCount).ToList();
            }

            var resultado = new AgPlanOleadasResultado(
                planes,
                infactList,
                infactList.Select(i => $"Infactible {i}").ToList(),
                bandas.Select(b => (IReadOnlyList<int>)b.Select(r => r.Dia).Distinct().ToList()).ToList(),
                string.Empty,
                0,
                50.0);

            var clave = string.Join("|", bandas.Select(b =>
                string.Join(".", b.Select(r => r.Dia).Distinct().OrderBy(d => d))));

            int maxRanuras = bandas.Count > 0 ? bandas.Max(b => b.Count) : 0;

            return new AgEvaluacionParticion(
                resultado,
                patronesNoAlojables,
                patronesAlojables ?? new List<int>(),
                makespan,
                bandas.Count,
                maxRanuras,
                clave);
        }
    }
}

