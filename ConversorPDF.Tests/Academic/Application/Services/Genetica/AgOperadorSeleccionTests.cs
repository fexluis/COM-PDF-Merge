using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgOperadorSeleccionTests
    {
        [Fact]
        public void SeleccionarPorTorneo_WhenEmptyPopulation_ShouldThrowArgumentException()
        {
            // Arrange
            var sut = new AgOperadorSeleccion();

            // Act
            Action act = () => sut.SeleccionarPorTorneo(new List<AgIndividuo>());

            // Assert
            act.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void SeleccionarPorTorneo_WhenTournamentCoversPopulation_ShouldSelectHighestFitness()
        {
            // Arrange
            var sut = new AgOperadorSeleccion(new Random(42));
            var poblacion = new List<AgIndividuo>
            {
                new AgIndividuo { Id = 1, Fitness = 10.0 },
                new AgIndividuo { Id = 2, Fitness = 95.0 },
                new AgIndividuo { Id = 3, Fitness = 40.0 }
            };

            // Act: tamaño torneo suficientemente amplio garantiza evaluar a los mejores aspirantes
            var elegido = sut.SeleccionarPorTorneo(poblacion, tamanoTorneo: 20);

            // Assert
            elegido.Should().NotBeNull();
            elegido.Fitness.Should().Be(95.0);
            elegido.Id.Should().Be(2);
        }
    }
}
