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
    public class AgRandomSeedTests
    {
        [Fact]
        public void Seleccion_WhenSameSeed_ShouldYieldIdenticalSelectionSequence()
        {
            // Arrange
            const int seed = 12345;
            var sut1 = new AgOperadorSeleccion(seed);
            var sut2 = new AgOperadorSeleccion(seed);
            var sutDifferent = new AgOperadorSeleccion(99999);

            var poblacion = new List<AgIndividuo>();
            for (int i = 0; i < 20; i++)
            {
                poblacion.Add(new AgIndividuo { Fitness = i * 1.5 });
            }

            // Act
            var seleccion1 = Enumerable.Range(0, 10).Select(_ => sut1.SeleccionarPorTorneo(poblacion, 3).Fitness).ToList();
            var seleccion2 = Enumerable.Range(0, 10).Select(_ => sut2.SeleccionarPorTorneo(poblacion, 3).Fitness).ToList();
            var seleccionDiff = Enumerable.Range(0, 10).Select(_ => sutDifferent.SeleccionarPorTorneo(poblacion, 3).Fitness).ToList();

            // Assert
            seleccion1.Should().Equal(seleccion2);
            seleccion1.Should().NotEqual(seleccionDiff);
        }

        [Fact]
        public void Cruce_WhenSameSeed_ShouldProduceIdenticalOffspring()
        {
            // Arrange
            const int seed = 42;
            var sut1 = new AgOperadorCruce(seed);
            var sut2 = new AgOperadorCruce(seed);
            var sutDifferent = new AgOperadorCruce(8888);

            var p1 = CrearIndividuoPrueba(1);
            var p2 = CrearIndividuoPrueba(2);

            // Act
            var (h1_a, h2_a) = sut1.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: null);
            var (h1_b, h2_b) = sut2.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: null);
            var (h1_diff, _) = sutDifferent.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: null);

            // Assert
            h1_a.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero))
                .Should().Equal(h1_b.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero)));
            h2_a.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero))
                .Should().Equal(h2_b.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero)));
        }

        [Fact]
        public void AsignadorTercero_WhenSameSeed_ShouldMakeIdenticalChoices()
        {
            // Arrange
            const int seed = 777;
            var sut1 = new AgAsignadorTercero(seed);
            var sut2 = new AgAsignadorTercero(seed);

            var contexto = new AgContexto();
            contexto.PoolPorAsignatura[100] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 1, Prioridad = 1, Activo = 1 },
                new TerceroAsignatura { IdTercero = 2, Prioridad = 1, Activo = 1 },
                new TerceroAsignatura { IdTercero = 3, Prioridad = 1, Activo = 1 },
                new TerceroAsignatura { IdTercero = 4, Prioridad = 1, Activo = 1 }
            };

            var elecciones1 = new List<int>();
            var elecciones2 = new List<int>();

            for (int i = 0; i < 20; i++)
            {
                var g1 = new AgGen { IdAsignatura = 100 };
                var g2 = new AgGen { IdAsignatura = 100 };
                sut1.AsignarDocente(g1, contexto);
                sut2.AsignarDocente(g2, contexto);
                elecciones1.Add(g1.IdTercero);
                elecciones2.Add(g2.IdTercero);
            }

            // Assert
            elecciones1.Should().Equal(elecciones2);
        }

        [Fact]
        public void Mutacion_WhenSameSeed_ShouldYieldIdenticalResults()
        {
            // Arrange
            const int seed = 54321;
            var sut1 = new AgOperadorMutacion(seed);
            var sut2 = new AgOperadorMutacion(seed);

            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, Activo = 1 });
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 101, Activo = 1 },
                new TerceroAsignatura { IdTercero = 102, Activo = 1 }
            };

            var ind1 = CrearIndividuoPrueba(10);
            var ind2 = CrearIndividuoPrueba(10);

            // Act
            sut1.Mutar(ind1, tasaMutacion: 0.8, contexto);
            sut2.Mutar(ind2, tasaMutacion: 0.8, contexto);

            // Assert
            ind1.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero))
                .Should().Equal(ind2.Genes.Select(g => (g.IdBloqueHorario, g.IdTercero)));
        }

        [Fact]
        public void FabricaIndividuos_WhenSeeded_ShouldAcceptSeedConstructor()
        {
            // Arrange & Act
            var fabrica1 = new AgFabricaIndividuos(1234);
            var fabrica2 = new AgFabricaIndividuos(1234, new AgPlanificadorPatronSemanal());

            // Assert
            fabrica1.Should().NotBeNull();
            fabrica2.Should().NotBeNull();
        }

        private static AgIndividuo CrearIndividuoPrueba(int seedOffset)
        {
            var ind = new AgIndividuo();
            for (int i = 1; i <= 6; i++)
            {
                ind.Genes.Add(new AgGen
                {
                    Id = i,
                    IdAsignatura = 10,
                    IdAula = 1,
                    IdTercero = 100 + i + seedOffset,
                    IdBloqueHorario = (i % 2) + 1,
                    FechaInicioEstimada = new DateTime(2026, 3, 1),
                    FechaFinEstimada = new DateTime(2026, 4, 1),
                    HorasSemanales = 4,
                    EsFijo = 0,
                    EsVacante = 0,
                    ClaveGrupo = "GRUPO_A"
                });
            }
            return ind;
        }
    }
}
