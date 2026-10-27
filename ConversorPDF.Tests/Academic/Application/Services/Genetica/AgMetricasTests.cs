using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgMetricasTests
    {
        [Fact]
        public void Recalcular_WhenZeroViolationsAndNoVacancies_ShouldSetEstadoFactible()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo { HardViolations = 0 };
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                EsVacante = 0,
                FechaInicioEstimada = new DateTime(2026, 2, 15),
                FechaFinEstimada = new DateTime(2026, 4, 15)
            });

            // Act
            AgMetricas.Recalcular(individuo, contexto);

            // Assert
            individuo.Estado.Should().Be(AgEstadoIndividuo.Factible);
            individuo.Genes[0].DentroDeVentana.Should().Be(1);
        }

        [Fact]
        public void Recalcular_WhenVacanteExists_ShouldSetEstadoInfactible()
        {
            // Arrange
            var contexto = new AgContexto();
            var individuo = new AgIndividuo { HardViolations = 0 };
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 0,
                EsVacante = 1
            });

            // Act
            AgMetricas.Recalcular(individuo, contexto);

            // Assert
            individuo.Estado.Should().Be(AgEstadoIndividuo.Infactible);
        }
    }
}
