using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgHuellaAptitudTests
    {
        [Fact]
        public void CalcularHuellaEstructural_WhenGenesHaveSameProperties_ShouldProduceIdenticalHash()
        {
            // Arrange
            var g1 = new AgGen
            {
                IdAsignatura = 10,
                IdAula = 20,
                IdBloqueHorario = 30,
                IdTercero = 100,
                EsVacante = 0,
                DiaSemana = 1,
                EsFijo = 0,
                EsFijoTemporal = false,
                IdPrograma = 2,
                IdSede = 5,
                ClaveGrupo = "GRP-A",
                FechaInicioEstimada = new DateTime(2026, 2, 1),
                FechaFinEstimada = new DateTime(2026, 6, 30),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 7), new DateTime(2026, 2, 14) }
            };

            var g2 = g1.Clonar();

            // Act
            long h1 = g1.CalcularHuellaEstructural();
            long h2 = g2.CalcularHuellaEstructural();

            // Assert
            h1.Should().Be(h2);
            h1.Should().NotBe(0L);
        }

        [Theory]
        [InlineData(1, 2, 0, 0, 0, 0)]
        [InlineData(0, 0, 10, 20, 0, 0)]
        [InlineData(0, 0, 0, 0, 101, 102)]
        public void CalcularHuellaEstructural_WhenGeneAttributesDiffer_ShouldProduceDifferentHash(
            int asig1, int asig2,
            int aula1, int aula2,
            int doc1, int doc2)
        {
            // Arrange
            var g1 = new AgGen
            {
                IdAsignatura = asig1 > 0 ? asig1 : 1,
                IdAula = aula1 > 0 ? aula1 : 1,
                IdTercero = doc1 > 0 ? doc1 : 100,
                IdBloqueHorario = 5,
                DiaSemana = 1
            };

            var g2 = new AgGen
            {
                IdAsignatura = asig2 > 0 ? asig2 : 1,
                IdAula = aula2 > 0 ? aula2 : 1,
                IdTercero = doc2 > 0 ? doc2 : 100,
                IdBloqueHorario = 5,
                DiaSemana = 1
            };

            // Act
            long h1 = g1.CalcularHuellaEstructural();
            long h2 = g2.CalcularHuellaEstructural();

            // Assert
            h1.Should().NotBe(h2);
        }

        [Fact]
        public void CalcularHuellaEstructural_WhenIndividuoHasEmptyGenes_ShouldReturnZero()
        {
            // Arrange
            var ind = new AgIndividuo();

            // Act
            long huella = ind.CalcularHuellaEstructural();

            // Assert
            huella.Should().Be(0L);
            ind.HuellaEstructural.Should().Be(0L);
        }

        [Fact]
        public void CalcularHuellaEstructural_WhenTwoIndividuosHaveIdenticalGenes_ShouldProduceIdenticalHash()
        {
            // Arrange
            var ind1 = new AgIndividuo();
            ind1.Genes.Add(new AgGen { IdAsignatura = 1, IdAula = 10, IdTercero = 101, IdBloqueHorario = 1 });
            ind1.Genes.Add(new AgGen { IdAsignatura = 2, IdAula = 20, IdTercero = 102, IdBloqueHorario = 2 });

            var ind2 = ind1.Clonar();

            // Act
            long h1 = ind1.CalcularHuellaEstructural();
            long h2 = ind2.CalcularHuellaEstructural();

            // Assert
            h1.Should().Be(h2);
            ind1.HuellaEstructural.Should().Be(h1);
            ind2.HuellaEstructural.Should().Be(h2);
        }

        [Fact]
        public void CalcularHuellaEstructural_WhenOneGeneDiffersInIndividuo_ShouldProduceDifferentHash()
        {
            // Arrange
            var ind1 = new AgIndividuo();
            ind1.Genes.Add(new AgGen { IdAsignatura = 1, IdAula = 10, IdTercero = 101, IdBloqueHorario = 1 });
            ind1.Genes.Add(new AgGen { IdAsignatura = 2, IdAula = 20, IdTercero = 102, IdBloqueHorario = 2 });

            var ind2 = ind1.Clonar();
            ind2.Genes[1].IdTercero = 999; // modificar un docente

            // Act
            long h1 = ind1.CalcularHuellaEstructural();
            long h2 = ind2.CalcularHuellaEstructural();

            // Assert
            h1.Should().NotBe(h2);
        }

        [Fact]
        public void EvaluarConCache_WhenIdenticalStructureEvaluated_ShouldProduceIdenticalFitnessBitForBit()
        {
            // Arrange
            var validador = new AgValidadorReglas();
            var evaluador = new AgEvaluadorAptitud(validador);
            var contexto = new AgContexto();

            var indOriginal = new AgIndividuo();
            indOriginal.Genes.Add(new AgGen { IdAsignatura = 1, IdAula = 10, IdTercero = 101, IdBloqueHorario = 1 });
            evaluador.EvaluarAptitud(indOriginal, contexto);

            var indCopia = indOriginal.Clonar();
            indCopia.Estado = AgEstadoIndividuo.NoEvaluado;
            indCopia.Fitness = 0.0;
            indCopia.HardViolations = 0;
            indCopia.SoftViolations = 0.0;

            // Act
            evaluador.EvaluarAptitud(indCopia, contexto);

            // Assert
            indCopia.Fitness.Should().Be(indOriginal.Fitness);
            indCopia.HardViolations.Should().Be(indOriginal.HardViolations);
            indCopia.SoftViolations.Should().Be(indOriginal.SoftViolations);
            indCopia.Estado.Should().Be(indOriginal.Estado);
            indCopia.CalcularHuellaEstructural().Should().Be(indOriginal.CalcularHuellaEstructural());
        }

        [Fact]
        public void Clonar_WhenIndividuoHasHuellaEstructural_ShouldPreserveHuella()
        {
            // Arrange
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { IdAsignatura = 1, IdAula = 10, IdTercero = 101, IdBloqueHorario = 1 });
            ind.CalcularHuellaEstructural();

            // Act
            var clon = ind.Clonar();

            // Assert
            clon.HuellaEstructural.Should().Be(ind.HuellaEstructural);
            clon.CalcularHuellaEstructural().Should().Be(ind.HuellaEstructural);
        }
    }
}
