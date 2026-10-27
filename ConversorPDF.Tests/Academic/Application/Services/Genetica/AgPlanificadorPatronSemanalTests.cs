using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPlanificadorPatronSemanalTests
    {
        private readonly AgPlanificadorPatronSemanal _sut = new AgPlanificadorPatronSemanal();

        [Fact]
        public void Construir_When48HorasY2DiasY6Semanas_ShouldGenerate12SesionesYHistogramaConsistente()
        {
            // Arrange
            var diasK = new List<int> { 4, 5 }; // Jueves, Viernes
            int totalHoras = 48;
            var ini = new DateTime(2026, 8, 13);
            var fin = new DateTime(2026, 9, 25);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            var patron = result.Value;
            patron.FechasValidas.Count.Should().Be(12);
            patron.DiasSemana.Should().BeEquivalentTo(diasK);
            patron.SesionesPorSemana.Sum().Should().Be(12);
        }

        [Fact]
        public void Construir_When32HorasY2Dias_ShouldGenerate8Sesiones()
        {
            // Arrange
            var diasK = new List<int> { 2, 3 }; // Martes, Miércoles
            int totalHoras = 32;
            var ini = new DateTime(2026, 8, 11);
            var fin = new DateTime(2026, 9, 15);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.FechasValidas.Count.Should().Be(8);
            result.Value.SesionesPorSemana.Sum().Should().Be(8);
        }

        [Fact]
        public void Construir_When64HorasY2Dias_ShouldGenerate16Sesiones()
        {
            // Arrange
            var diasK = new List<int> { 5, 6 }; // Viernes, Sábado
            int totalHoras = 64;
            var ini = new DateTime(2026, 8, 14);
            var fin = new DateTime(2026, 10, 15);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.FechasValidas.Count.Should().Be(16);
            result.Value.SesionesPorSemana.Sum().Should().Be(16);
        }

        [Fact]
        public void Construir_WhenVentanaIncompleta_ShouldReturnFailure()
        {
            // Arrange
            var diasK = new List<int> { 5 }; // Solo Viernes (1 día)
            int totalHoras = 48; // 12 sesiones
            var ini = new DateTime(2026, 8, 14);
            var fin = new DateTime(2026, 9, 4); // Solo 4 viernes

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Infactible");
        }

        [Fact]
        public void Construir_WhenDiasKVacio_ShouldReturnFailure()
        {
            // Arrange
            var diasK = new List<int>();
            int totalHoras = 32;
            var ini = new DateTime(2026, 8, 14);
            var fin = new DateTime(2026, 10, 14);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("especificaron días");
        }

        [Fact]
        public void Construir_When64HorasY4DiasY5Semanas_ShouldSatisfyInvariantes()
        {
            // Arrange (Fase 424: Invariantes canónicos Sum()==S, Max()<=nCap, orden ascendente)
            var diasK = new List<int> { 2, 3, 4, 5 };
            int totalHoras = 64;
            var ini = new DateTime(2026, 8, 11);
            var fin = new DateTime(2026, 9, 14);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.SesionesPorSemana.Sum().Should().Be(16);
            result.Value.SesionesPorSemana.All(s => s <= 4).Should().BeTrue();
            result.Value.SesionesPorSemana.Should().BeInAscendingOrder();
            result.Value.SesionesPorSemana.Max().Should().BeLessOrEqualTo(5);
            result.Value.FechasValidas.Count.Should().Be(16);
        }

        [Fact]
        public void Construir_When48HorasY3DiasY5Semanas_ShouldSatisfyInvariantes()
        {
            // Arrange (Fase 424 / Fase 430: Invariantes canónicos Sum()==S, Max()<=nCap, orden ascendente)
            var diasK = new List<int> { 2, 3, 4 };
            int totalHoras = 48;
            var ini = new DateTime(2026, 8, 11);
            var fin = new DateTime(2026, 9, 14);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.SesionesPorSemana.Sum().Should().Be(12);
            result.Value.SesionesPorSemana.All(s => s <= 3).Should().BeTrue();
            result.Value.SesionesPorSemana.Should().BeInAscendingOrder();
            result.Value.SesionesPorSemana.Max().Should().BeLessOrEqualTo(5);
            result.Value.FechasValidas.Count.Should().Be(12);
        }

        [Fact]
        public void Construir_When32HorasY2DiasY5Semanas_ShouldSatisfyInvariantes()
        {
            // Arrange (Fase 424 / Fase 430: 32h / 2 días / 5 sem => Sum==8, Max<=2, orden ascendente)
            var diasK = new List<int> { 2, 3 };
            int totalHoras = 32;
            var ini = new DateTime(2026, 8, 11);
            var fin = new DateTime(2026, 9, 14);

            // Act
            var result = _sut.Construir(diasK, totalHoras, ini, fin);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.SesionesPorSemana.Sum().Should().Be(8);
            result.Value.SesionesPorSemana.All(s => s <= 2).Should().BeTrue();
            result.Value.SesionesPorSemana.Should().BeInAscendingOrder();
            result.Value.SesionesPorSemana.Max().Should().BeLessOrEqualTo(5);
            result.Value.FechasValidas.Count.Should().Be(8);
        }

        [Fact]
        public void Construir_WhenVentanaSinDiasHabiles_ShouldReturnFailure()
        {
            // Arrange (Fase 431: Ventana donde todos los días hábiles son festivos o no coinciden con diasK)
            var diasK = new List<int> { 1 }; // Solo Lunes
            int totalHoras = 32;
            var ini = new DateTime(2026, 8, 17); // Lunes festivo
            var fin = new DateTime(2026, 8, 17);

            // Act: Todo día en la ventana es festivo
            var result = _sut.Construir(diasK, totalHoras, ini, fin, esFestivo: d => true);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("no contiene días hábiles");
        }

        [Fact]
        public void DiasKDe_PatronNoCarril_1_3_NormalizaAlCarril()
        {
            // Arrange (Fase 1074): patrón "1,3" no es un carril consecutivo de "LUN-JUE" (1,2,3,4).
            // Debe normalizarse al carril 0 -> [1, 2] o carril 1 -> [3, 4].
            var aa = new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, DiasPatron = "1,3" };
            var aula = new AgAula { Id = 1, DiasVentana = "LUN-JUE" };

            // Act
            var dias0 = AgPlanificadorPatronSemanal.DiasKDe(aa, aula, carril: 0);
            var dias1 = AgPlanificadorPatronSemanal.DiasKDe(aa, aula, carril: 1);

            // Assert
            dias0.Should().Equal(1, 2);
            dias1.Should().Equal(3, 4);
        }

        [Fact]
        public void DiasKDe_TextoNoCarril_1_3_NormalizaAlCarril()
        {
            // Arrange (Fase 1084): Overload textual con patrón "1,3" que no es carril disjunto de "LUN-JUE"
            // Debe normalizarse al carril correspondiente de la ventana.
            var aula = new AgAula { Id = 1, DiasVentana = "LUN-JUE" };

            // Act
            var dias0 = AgPlanificadorPatronSemanal.DiasKDe("1,3", "LUN-JUE", aula, carril: 0);
            var dias1 = AgPlanificadorPatronSemanal.DiasKDe("1,3", "LUN-JUE", aula, carril: 1);

            // Assert
            dias0.Should().Equal(1, 2);
            dias1.Should().Equal(3, 4);
        }
    }
}
