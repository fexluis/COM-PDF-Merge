using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgEvaluadorAptitudTests
    {
        [Fact]
        public void EvaluarAptitud_WhenValidGenesAndZeroViolations_ShouldCalculateCorrectFitnessFormula()
        {
            // Arrange
            var validador = new AgValidadorReglas();
            var sut = new AgEvaluadorAptitud(validador);
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 1,
                IdTercero = 10,
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10),
                EsFijo = 1
            });
            contexto.RegistrarDocenteHabilitado(10, 0, 1, string.Empty);
            contexto.AulasPorId[1] = new AgAula { Id = 1, Codigo = "101", Estado = AgEstadoAula.Disponible };

            // Act
            sut.EvaluarAptitud(individuo, contexto);

            // Assert
            // F = F1 (1) - F2 (0) + 2*F3 (2*1) + F4 (0) + 2*F5 (2*1) = 1 + 2 + 2 = 5.0
            individuo.HardViolations.Should().Be(0);
            individuo.Fitness.Should().Be(5.0);
            individuo.EsFactible.Should().BeTrue();
        }

        [Fact]
        public void EvaluarAptitud_WhenHardViolationsExist_ShouldReturnLowerFitnessAndInfeasible()
        {
            // Arrange
            var validador = new AgValidadorReglas();
            var sut = new AgEvaluadorAptitud(validador);
            var contexto = new AgContexto();
            contexto.Pesos.PenalidadHardBase = 1000.0;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                IdTercero = 999, // Docente no habilitado en pool
                IdPrograma = 1,
                IdAsignatura = 1,
                Semestre = "I"
            });

            // Act
            sut.EvaluarAptitud(individuo, contexto);

            // Assert
            individuo.HardViolations.Should().BeGreaterThan(0);
            individuo.Fitness.Should().BeLessThan(0.0);
            individuo.EsFactible.Should().BeFalse();
        }
    }
}
