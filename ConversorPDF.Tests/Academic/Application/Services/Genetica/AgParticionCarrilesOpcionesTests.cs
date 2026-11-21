using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgParticionCarrilesOpcionesTests
    {
        [Fact]
        public void Seleccionar_ConOpciones_EvaluaSoloOpcionesYMaximoDosCandidatas()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = opcion != null ? new List<AgOpcionParticion> { opcion } : new List<AgOpcionParticion>();

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    // Referencia canónica con 1 infactible
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 99 });
                    return new AgEvaluacionParticion(resRef, 1, new List<int>(), 14, bandas.Count, 2, "ref");
                }
                // Candidata con 0 infactibles (mejora estricta)
                var resMejora = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resMejora, 0, new List<int> { 0, 1 }, 10, bandas.Count, 2, "2,3|4,5");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.CandidatasEvaluadas.Should().BeLessOrEqualTo(2);
            seleccion.Cambiada.Should().BeTrue();
            seleccion.Elegida.Clave.Should().Be("2,3|4,5");
        }

        [Fact]
        public void Seleccionar_SinMejora_DevuelveReferencia()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };

            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = opcion != null ? new List<AgOpcionParticion> { opcion } : new List<AgOpcionParticion>();

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    // Referencia ya óptima (0 infactibles, makespan 8)
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                    return new AgEvaluacionParticion(resRef, 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "ref");
                }
                // Candidata con peor makespan (sin mejora)
                var resPeor = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resPeor, 0, new List<int> { 0, 1 }, 12, bandas.Count, 2, "2,3|4,5");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeFalse();
            seleccion.Elegida.Should().Be(seleccion.Referencia);
        }

        [Fact]
        public void Seleccionar_SinOpciones_MantieneComportamientoPrevio()
        {
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };

            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                var resOleadas = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resOleadas, 0, new List<int>(), 10, bandas.Count, 2, "clave");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                new List<string>(),
                canonicas,
                evaluar,
                opciones: null);

            seleccion.Cambiada.Should().BeFalse();
        }

        [Fact]
        public void Seleccionar_ExponeOrigenYOpcionTexto()
        {
            // FASE 2220 / FASE 2221: Selección exitosa de catálogo expone origen=catalogo y OpcionTexto
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };
            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = new List<AgOpcionParticion> { opcion };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 99 });
                    return new AgEvaluacionParticion(resRef, 1, new List<int>(), 14, bandas.Count, 2, "ref");
                }
                var resMejora = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resMejora, 0, new List<int> { 0, 1 }, 10, bandas.Count, 2, "2,3|4,5");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeTrue();
            seleccion.Origen.Should().Be("catalogo");
            seleccion.OpcionTexto.Should().Be("2,3|4,5");
        }

        [Fact]
        public void Seleccionar_CatalogoSinCambio_ReportaOrigenCatalogo()
        {
            // FASE 2220 / FASE 2221: Catálogo sin mejora reporta cambiada=false y origen=catalogo
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };
            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = new List<AgOpcionParticion> { opcion };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                    return new AgEvaluacionParticion(resRef, 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "ref");
                }
                var resPeor = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resPeor, 0, new List<int> { 0, 1 }, 12, bandas.Count, 2, "2,3|4,5");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeFalse();
            seleccion.Origen.Should().Be("catalogo_fallback");
            seleccion.OpcionTexto.Should().Be("2,3|4,5");
        }

        [Fact]
        public void Catalogo_4567_OpcionTexto_ConservaFranjas()
        {
            // FASE 2224 / FASE 2227: OpcionTexto conserva las franjas (5,6M|4,6T,7)
            var ranuras = new List<AgRanura>
            {
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.MANANA, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.TARDE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3], ranuras[4] }
            };
            var patrones = new List<string> { "5,6M", "4,6T,7" };
            var opcion = AgOpcionesParticion.Resolver("JUE-DOM", ranuras, patrones);
            var opciones = new List<AgOpcionParticion> { opcion };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 99 });
                    return new AgEvaluacionParticion(resRef, 1, new List<int>(), 14, bandas.Count, 2, "ref");
                }
                var resMejora = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resMejora, 0, new List<int> { 0, 1 }, 10, bandas.Count, 2, "5,6|4,6,7");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 4, 5, 6, 7 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeTrue();
            seleccion.Origen.Should().Be("catalogo");
            seleccion.OpcionTexto.Should().Be("5,6M|4,6T,7");
        }

        [Fact]
        public void CatalogoSinMejora_EjecutaEnumeracion()
        {
            // FASE 2223 / FASE 2227: Catálogo sin mejora ejecuta fallback contiguo de Particiones()
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "1,2", "3,4" };
            var opcionInerte = new AgOpcionParticion("quedarse", "1,2|3,4", canonicas);
            var opciones = new List<AgOpcionParticion> { opcionInerte };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1) // Referencia
                {
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1, 2 });
                    return new AgEvaluacionParticion(resRef, 1, new List<int>(), 10, bandas.Count, 2, "ref");
                }
                if (llamadas == 2) // Opción catálogo: no mejora infactibles
                {
                    var resCat = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1, 2 });
                    return new AgEvaluacionParticion(resCat, 1, new List<int>(), 10, bandas.Count, 2, "cat");
                }
                // Enumeración contigua: encuentra mejora
                var resEnum = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resEnum, 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "particion_enum");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2, 3, 4 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeTrue();
            seleccion.Origen.Should().Be("catalogo_fallback");
            llamadas.Should().BeGreaterThan(2);
        }

        [Fact]
        public void CatalogoSinMejora_RegistraOpcionIntentada()
        {
            // FASE 2225 / FASE 2227: Catálogo sin mejora reporta OpcionTexto con la opción probada
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };
            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = new List<AgOpcionParticion> { opcion };

            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resRef, 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "ref");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeFalse();
            seleccion.OpcionTexto.Should().Be("2,3|4,5");
        }

        [Fact]
        public void CatalogoFallback_OpcionTextoEsLaIntentada()
        {
            // FASE 2228 / FASE 2231: En fallback OpcionTexto conserva la opción intentada ("2,3|4,5"), no la clave contigua
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "1,2", "3,4" };
            var opcionInerte = new AgOpcionParticion("quedarse", "2,3|4,5", canonicas);
            var opciones = new List<AgOpcionParticion> { opcionInerte };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                    return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1 }), 1, new List<int>(), 10, bandas.Count, 2, "ref");
                if (llamadas == 2)
                    return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1 }), 1, new List<int>(), 10, bandas.Count, 2, "cat");
                return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()), 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "1.2|3.4");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2, 3, 4 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeTrue();
            seleccion.Origen.Should().Be("catalogo_fallback");
            seleccion.OpcionTexto.Should().Be("2,3|4,5");
        }

        [Fact]
        public void CatalogoFallback_ParticionElegidaEsLaContigua()
        {
            // FASE 2228 / FASE 2231: En fallback ParticionElegidaTexto expone la clave de la partición contigua
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "1,2", "3,4" };
            var opcionInerte = new AgOpcionParticion("quedarse", "2,3|4,5", canonicas);
            var opciones = new List<AgOpcionParticion> { opcionInerte };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1) return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1 }), 1, new List<int>(), 10, bandas.Count, 2, "ref");
                if (llamadas == 2) return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1 }), 1, new List<int>(), 10, bandas.Count, 2, "cat");
                return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()), 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "1.2|3.4");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2, 3, 4 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.ParticionElegidaTexto.Should().Be("1.2|3.4");
        }

        [Fact]
        public void CatalogoFallback_OpcionesEvaluadasEsUno()
        {
            // FASE 2229 / FASE 2231: OpcionesEvaluadas cuenta las opciones de catálogo evaluadas independientemente de contiguas
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "1,2", "3,4" };
            var opcionInerte = new AgOpcionParticion("quedarse", "2,3|4,5", canonicas);
            var opciones = new List<AgOpcionParticion> { opcionInerte };

            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()), 0, new List<int> { 0, 1 }, 8, bandas.Count, 2, "clave");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2, 3, 4 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.OpcionesEvaluadas.Should().Be(1);
        }

        [Fact]
        public void SinPatronesSinInfactibles_NoEnumera()
        {
            // FASE 2230 / FASE 2231: Sin patrones y sin infactibles ni catálogo, salida temprana sin enumeración
            var ranuras = new List<AgRanura>
            {
                new AgRanura(1, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0] },
                new List<AgRanura> { ranuras[1] }
            };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                return new AgEvaluacionParticion(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()), 0, new List<int>(), 8, bandas.Count, 2, "ref");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 1, 2 },
                ranuras,
                new List<string>(),
                canonicas,
                evaluar,
                opciones: null);

            seleccion.Cambiada.Should().BeFalse();
            seleccion.CandidatasEvaluadas.Should().Be(0);
            llamadas.Should().Be(1);
        }

        [Fact]
        public void Seleccionar_CatalogoExitoso_CandidatasEvaluadasEsMayorACero()
        {
            // FASE 2232: En catálogo exitoso CandidatasEvaluadas reporta las opciones evaluadas (> 0)
            var ranuras = new List<AgRanura>
            {
                new AgRanura(2, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(3, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0)
            };
            var canonicas = new List<IReadOnlyList<AgRanura>>
            {
                new List<AgRanura> { ranuras[0], ranuras[1] },
                new List<AgRanura> { ranuras[2], ranuras[3] }
            };
            var patrones = new List<string> { "2,3", "4,5" };
            var opcion = AgOpcionesParticion.Resolver("MAR-VIE", ranuras, patrones);
            var opciones = new List<AgOpcionParticion> { opcion };

            int llamadas = 0;
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar = bandas =>
            {
                llamadas++;
                if (llamadas == 1)
                {
                    var resRef = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 99 });
                    return new AgEvaluacionParticion(resRef, 1, new List<int>(), 14, bandas.Count, 2, "ref");
                }
                var resMejora = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>());
                return new AgEvaluacionParticion(resMejora, 0, new List<int> { 0, 1 }, 10, bandas.Count, 2, "2,3|4,5");
            };

            var seleccion = AgParticionCarriles.Seleccionar(
                new List<int> { 2, 3, 4, 5 },
                ranuras,
                patrones,
                canonicas,
                evaluar,
                opciones: opciones);

            seleccion.Cambiada.Should().BeTrue();
            seleccion.Origen.Should().Be("catalogo");
            seleccion.CandidatasEvaluadas.Should().Be(1);
            seleccion.OpcionesEvaluadas.Should().Be(1);
        }
    }
}
