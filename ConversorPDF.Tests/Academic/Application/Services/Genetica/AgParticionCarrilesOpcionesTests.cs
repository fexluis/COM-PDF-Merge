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
    }
}
