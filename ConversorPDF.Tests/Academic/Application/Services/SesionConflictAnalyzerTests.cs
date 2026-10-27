using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class SesionConflictAnalyzerTests
    {
        [Fact]
        public void Analizar_WhenNoOverlapSameCity_ShouldReturnNinguno()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "10:00 AM A 12:00 M", SedeCiudad = "CALI" }
            };
            var salvedades = new List<SalvedadConflicto>();

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            result[1].Should().Be(EstadoConflicto.Ninguno);
            result[2].Should().Be(EstadoConflicto.Ninguno);
        }

        [Fact]
        public void Analizar_WhenDifferentCitySameTime_ShouldReturnConflictoReal()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "9:00 AM A 11:00 AM", SedeCiudad = "BUGA" }
            };
            var salvedades = new List<SalvedadConflicto>();

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            result[1].Should().Be(EstadoConflicto.ConflictoReal);
            result[2].Should().Be(EstadoConflicto.ConflictoReal);
        }

        [Fact]
        public void Analizar_WhenDifferentCityDifferentTime_ShouldReturnRevisarViabilidad()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "2:00 PM A 4:00 PM", SedeCiudad = "BUGA" }
            };
            var salvedades = new List<SalvedadConflicto>();

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            result[1].Should().Be(EstadoConflicto.RevisarViabilidad);
            result[2].Should().Be(EstadoConflicto.RevisarViabilidad);
        }

        [Fact]
        public void Analizar_WhenDifferentCityDifferentTimeButSaved_ShouldReturnSalvado()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "2:00 PM A 4:00 PM", SedeCiudad = "BUGA" }
            };
            var salvedades = new List<SalvedadConflicto>
            {
                new SalvedadConflicto { IdSesionA = 1, IdSesionB = 2 }
            };

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            result[1].Should().Be(EstadoConflicto.Salvado);
            result[2].Should().Be(EstadoConflicto.Salvado);
        }

        [Fact]
        public void Analizar_WhenVirtualSession_ShouldNotHaveConflict()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "BUGA", Modalidad = "Virtual" }
            };
            var salvedades = new List<SalvedadConflicto>();

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            result[1].Should().Be(EstadoConflicto.Ninguno);
            result.Should().NotContainKey(2); // Virtual sessions are ignored
        }

        [Fact]
        public void Analizar_WhenMultipleSessionsOnSameDay_ShouldReturnWorstState()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "2:00 PM A 4:00 PM", SedeCiudad = "BUGA" },
                new SesionListItem { IdSesion = 3, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "BOGOTA" }
            };
            var salvedades = new List<SalvedadConflicto>
            {
                new SalvedadConflicto { IdSesionA = 1, IdSesionB = 2 } // 1 y 2 estan salvadas
            };

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            // 1 vs 2: Salvado
            // 1 vs 3: ConflictoReal
            // 2 vs 3: RevisarViabilidad
            result[1].Should().Be(EstadoConflicto.ConflictoReal);
            result[2].Should().Be(EstadoConflicto.RevisarViabilidad);
            result[3].Should().Be(EstadoConflicto.ConflictoReal);
        }

        [Fact]
        public void Analizar_WhenBoundaryHoursMatch_ShouldNotOverlap()
        {
            // Arrange
            var sesiones = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "8:00 AM A 10:00 AM", SedeCiudad = "CALI" },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 5, 23), HorarioSesion = "10:00 AM A 12:00 M", SedeCiudad = "BUGA" }
            };
            var salvedades = new List<SalvedadConflicto>();

            // Act
            var result = SesionConflictAnalyzer.Analizar(sesiones, salvedades);

            // Assert
            // No hay cruce de horarios, pero distinta sede
            result[1].Should().Be(EstadoConflicto.RevisarViabilidad);
            result[2].Should().Be(EstadoConflicto.RevisarViabilidad);
        }
    }
}
