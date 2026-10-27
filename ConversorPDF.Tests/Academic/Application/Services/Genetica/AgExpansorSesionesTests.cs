using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgExpansorSesionesTests
    {
        private readonly AgExpansorSesiones _sut = new AgExpansorSesiones();

        [Fact]
        public void Expandir_WhenVentanaIsNull_ShouldReturnFailure()
        {
            // Arrange
            var gen = new AgGen { Id = 1, IdBloqueHorario = 1 };
            var contexto = new AgContexto();

            // Act
            var result = _sut.Expandir(gen, 1, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Ventana semestral no resuelta");
        }

        [Fact]
        public void Expandir_WhenBloqueNotFound_ShouldReturnFailure()
        {
            // Arrange
            var gen = new AgGen { Id = 1, IdBloqueHorario = 99 };
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            // Act
            var result = _sut.Expandir(gen, 1, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Bloque horario 99 no encontrado");
        }

        [Fact]
        public void Expandir_WhenValidGenAndContext_ShouldGenerateSessionsWithCorrectIdGenAndIdIndividuo()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            var bloque = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 1, // Lunes
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4,
                Activo = 1
            };
            contexto.Bloques.Add(bloque);
            contexto.AulasPorId[1] = new AgAula { Id = 1, ModalidadVinculacion = "Presencial" };
            contexto.ConstruirIndices();

            var gen = new AgGen
            {
                Id = 42,
                IdIndividuo = 7,
                IdAsignatura = 1,
                IdTercero = 10,
                IdAula = 1,
                IdBloqueHorario = 1,
                TotalHorasAsignatura = 16,
                FechaInicioEstimada = new DateTime(2026, 2, 2), // Lunes
                FechaFinEstimada = new DateTime(2026, 3, 2)
            };

            // Act
            var result = _sut.Expandir(gen, 100, contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().NotBeEmpty();
            result.Value.Should().OnlyContain(s => s.IdGen == 42 && s.IdIndividuo == 7 && s.IdVinculacion == 100);
        }

        [Fact]
        public void Expandir_WhenHorasGenIsZero_ShouldReturnEmptyWithoutFailing()
        {
            // Arrange: Fase 576 => 0 h no debe fallar por bloque inexistente o ventana
            var gen = new AgGen { Id = 7, HorasGen = 0.0, IdBloqueHorario = 0 };
            var contexto = new AgContexto();

            // Act
            var res = _sut.Expandir(gen, 1, contexto);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Should().BeEmpty();
        }
    }
}
