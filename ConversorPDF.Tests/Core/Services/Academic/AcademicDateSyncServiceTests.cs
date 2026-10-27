using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AcademicDateSyncServiceTests
    {
        private readonly AcademicDateSyncService _sut;

        public AcademicDateSyncServiceTests()
        {
            _sut = new AcademicDateSyncService();
        }

        [Fact]
        public void CalculateAdjustment_WhenSesionesEmpty_ShouldReturnNoAdjustment()
        {
            // Arrange
            var sesiones = new List<Sesion>();
            var start = DateTime.Today;
            var end = DateTime.Today.AddDays(30);

            // Act
            var result = _sut.CalculateAdjustment(sesiones, start, end);

            // Assert
            result.NeedsAdjustment.Should().BeFalse();
        }

        [Fact]
        public void CalculateAdjustment_WhenDatesMatchExactly_ShouldReturnNoAdjustment()
        {
            // Arrange
            var start = new DateTime(2026, 4, 1);
            var end = new DateTime(2026, 4, 15);
            var sesiones = new List<Sesion>
            {
                new Sesion { FechaSesion = start },
                new Sesion { FechaSesion = end }
            };

            // Act
            var result = _sut.CalculateAdjustment(sesiones, start, end);

            // Assert
            result.NeedsAdjustment.Should().BeFalse();
        }

        [Fact]
        public void CalculateAdjustment_WhenStartDiffers_ShouldReturnAdjustment()
        {
            // Arrange
            var currentStart = new DateTime(2026, 4, 1);
            var currentEnd = new DateTime(2026, 4, 10);
            var realStart = new DateTime(2026, 4, 5); // Later than current
            
            var sesiones = new List<Sesion>
            {
                new Sesion { FechaSesion = realStart },
                new Sesion { FechaSesion = currentEnd }
            };

            // Act
            var result = _sut.CalculateAdjustment(sesiones, currentStart, currentEnd);

            // Assert
            result.NeedsAdjustment.Should().BeTrue();
            result.CalculatedStart.Should().Be(realStart);
            result.Messages.Should().Contain(m => m.Contains("Inicio ajustada"));
        }

        [Fact]
        public void CalculateAdjustment_WhenEndExceedsLimit_ShouldSetExceededFlag()
        {
            // Arrange
            var currentStart = new DateTime(2026, 4, 1);
            var currentEnd = new DateTime(2026, 4, 10);
            var realEnd = new DateTime(2026, 4, 15); // Exceeds limit
            
            var sesiones = new List<Sesion>
            {
                new Sesion { FechaSesion = currentStart },
                new Sesion { FechaSesion = realEnd }
            };

            // Act
            var result = _sut.CalculateAdjustment(sesiones, currentStart, currentEnd);

            // Assert
            result.NeedsAdjustment.Should().BeTrue();
            result.CalculatedEnd.Should().Be(realEnd);
            result.ExceededOriginalLimit.Should().BeTrue();
            result.Messages.Should().Contain(m => m.Contains("SUPERADO"));
        }
    }
}
