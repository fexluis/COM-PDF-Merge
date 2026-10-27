using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class CalculadoraViaticaSesionTests
    {
        [Fact]
        public void DeterminarViatica_WhenModalidadVirtual_ShouldReturnFalse()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 10);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>();

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "VIRTUAL", "BOGOTA", "MEDELLIN", rangos);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public void DeterminarViatica_WhenModalidadPlataforma_ShouldReturnFalse()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 10);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>();

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "PLATAFORMA", "BOGOTA", "MEDELLIN", rangos);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public void DeterminarViatica_WhenSameCity_ShouldReturnFalse()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 10);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>();

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "PRESENCIAL", "Bogotá D.C.", "BOGOTA", rangos);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public void DeterminarViatica_WhenPresencialWithDesplazamiento_ShouldReturnTrue()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 10);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>();

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "PRESENCIAL", "CALI", "BOGOTA", rangos);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public void DeterminarViatica_WhenSemipresencialInsidePresencialRange_ShouldReturnTrue()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 15);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>
            {
                (new DateTime(2026, 3, 1), new DateTime(2026, 3, 20))
            };

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "SEMIPRESENCIAL", "CALI", "BOGOTA", rangos);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public void DeterminarViatica_WhenMixtaInsidePresencialRange_ShouldReturnTrue()
        {
            // Arrange
            var fecha = new DateTime(2026, 3, 15);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>
            {
                (new DateTime(2026, 3, 1), new DateTime(2026, 3, 20))
            };

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "MIXTA", "CALI", "BOGOTA", rangos);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public void DeterminarViatica_WhenSemipresencialOutsidePresencialRange_ShouldReturnFalse()
        {
            // Arrange
            var fecha = new DateTime(2026, 4, 15);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>
            {
                (new DateTime(2026, 3, 1), new DateTime(2026, 3, 20))
            };

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "SEMIPRESENCIAL", "CALI", "BOGOTA", rangos);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public void DeterminarViatica_WhenMixtaOutsidePresencialRange_ShouldReturnFalse()
        {
            // Arrange
            var fecha = new DateTime(2026, 4, 15);
            var rangos = new List<(DateTime Inicio, DateTime Fin)>
            {
                (new DateTime(2026, 3, 1), new DateTime(2026, 3, 20))
            };

            // Act
            bool result = CalculadoraViaticaSesion.DeterminarViatica(fecha, "MIXTA", "CALI", "BOGOTA", rangos);

            // Assert
            result.Should().BeFalse();
        }
    }
}
