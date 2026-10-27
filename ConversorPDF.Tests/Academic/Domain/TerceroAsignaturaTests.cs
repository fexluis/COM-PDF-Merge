using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class TerceroAsignaturaTests
    {
        [Fact]
        public void Semestre_WhenAssignedWithSpacesOrLowercase_ShouldNormalizeTrimAndUpper()
        {
            // Arrange
            var entity = new TerceroAsignatura();

            // Act
            entity.Semestre = "  iv  ";

            // Assert
            entity.Semestre.Should().Be("IV");
        }

        [Fact]
        public void Semestre_WhenAssignedNull_ShouldDefaultToEmptyString()
        {
            // Arrange
            var entity = new TerceroAsignatura();

            // Act
            entity.Semestre = null;

            // Assert
            entity.Semestre.Should().BeEmpty();
        }

        [Fact]
        public void Activo_WhenInstantiated_ShouldDefaultToOne()
        {
            // Arrange & Act
            var entity = new TerceroAsignatura();

            // Assert
            entity.Activo.Should().Be(1);
        }

        [Fact]
        public void EstablecerFechaUltimaAsignacion_WhenValidDateProvided_ShouldFormatAsIsoString()
        {
            // Arrange
            var entity = new TerceroAsignatura();
            var fecha = new DateTime(2026, 5, 21);

            // Act
            entity.EstablecerFechaUltimaAsignacion(fecha);

            // Assert
            entity.FechaUltimaAsignacion.Should().Be("2026-05-21");
            entity.TryObtenerFechaUltimaAsignacion(out var fechaParseada).Should().BeTrue();
            fechaParseada.Should().Be(new DateTime(2026, 5, 21));
        }

        [Fact]
        public void EstablecerFechaUltimaAsignacion_WhenNullProvided_ShouldSetEmptyString()
        {
            // Arrange
            var entity = new TerceroAsignatura();

            // Act
            entity.EstablecerFechaUltimaAsignacion(null);

            // Assert
            entity.FechaUltimaAsignacion.Should().BeEmpty();
            entity.TryObtenerFechaUltimaAsignacion(out _).Should().BeFalse();
        }

        [Fact]
        public void TryObtenerFechaUltimaAsignacion_WhenInvalidFormat_ShouldReturnFalse()
        {
            // Arrange
            var entity = new TerceroAsignatura { FechaUltimaAsignacion = "21/05/2026" }; // Formato no ISO

            // Act
            bool result = entity.TryObtenerFechaUltimaAsignacion(out _);

            // Assert
            result.Should().BeFalse();
        }
    }
}
