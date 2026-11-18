using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgQuickWinsTests
    {
        [Fact]
        public void Mutar_WhenTasaZero_ShouldReturnFalseAndNotCallSellar()
        {
            // Arrange
            var sut = new AgOperadorMutacion(new Random(123));
            var individuo = new AgIndividuo { Estado = AgEstadoIndividuo.Factible };
            var contexto = new AgContexto();

            // Act
            bool cambios = sut.Mutar(individuo, tasaMutacion: 0.0, contexto: contexto);

            // Assert
            cambios.Should().BeFalse();
            individuo.Estado.Should().Be(AgEstadoIndividuo.Factible);
        }

        [Fact]
        public void Cruzar_WhenTasaCruceZero_ShouldReturnExactClonesWithPreservedState()
        {
            // Arrange
            var sut = new AgOperadorCruce(new Random(42));
            var padre1 = new AgIndividuo { Estado = AgEstadoIndividuo.Factible, Fitness = 150.0 };
            var padre2 = new AgIndividuo { Estado = AgEstadoIndividuo.Factible, Fitness = 140.0 };
            var gen = new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, IdBloqueHorario = 1 };
            padre1.Genes.Add(gen);
            padre2.Genes.Add(gen.Clonar());

            // Act
            var (hijo1, hijo2) = sut.Cruzar(padre1, padre2, tasaCruce: 0.0, generacion: 1, contexto: new AgContexto());

            // Assert
            hijo1.Fitness.Should().Be(150.0);
            hijo1.Estado.Should().Be(AgEstadoIndividuo.Factible);
            hijo2.Fitness.Should().Be(140.0);
            hijo2.Estado.Should().Be(AgEstadoIndividuo.Factible);
        }

        [Fact]
        public void EvaluarAptitud_WhenIndividualAlreadyEvaluated_ShouldSkipValidationUnlessForced()
        {
            // Arrange
            var mockValidador = new Mock<IAgValidadorReglas>();
            var sut = new AgEvaluadorAptitud(mockValidador.Object);
            var ind = new AgIndividuo
            {
                Estado = AgEstadoIndividuo.Factible,
                Fitness = 200.0
            };
            var contexto = new AgContexto();

            // Act 1: without force
            sut.EvaluarAptitud(ind, contexto);

            // Assert 1: skipped
            mockValidador.Verify(v => v.Evaluar(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()), Times.Never);
            ind.Fitness.Should().Be(200.0);

            // Act 2: with force
            sut.EvaluarAptitud(ind, contexto, forzar: true);

            // Assert 2: executed
            mockValidador.Verify(v => v.Evaluar(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()), Times.Once);
        }

        [Fact]
        public void ValidadorReglas_WhenIncluirDetalleJsonFalse_ShouldProduceDefaultDetalleJson()
        {
            // Arrange
            var sut = new AgValidadorReglas { IncluirDetalleJson = false };
            var contexto = new AgContexto();
            contexto.IdSede = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var bloque = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "07:00", HoraFin = "11:00", Activo = 1 };
            contexto.Bloques.Add(bloque);
            contexto.BloquesPorId[1] = bloque;

            var aula = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            var ind = new AgIndividuo();
            // Dos genes que generan cruce en la misma aula
            var gen1 = new AgGen { Id = 1, IdAula = 1, IdAsignatura = 10, IdBloqueHorario = 1, HorasGen = 4, TotalHorasAsignatura = 4, FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) } };
            var gen2 = new AgGen { Id = 2, IdAula = 1, IdAsignatura = 20, IdBloqueHorario = 1, HorasGen = 4, TotalHorasAsignatura = 4, FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) } };
            ind.Genes.Add(gen1);
            ind.Genes.Add(gen2);

            // Act
            sut.Evaluar(ind, contexto, incluirDetalleJson: false);

            // Assert
            ind.Violaciones.Should().NotBeEmpty();
            ind.Violaciones[0].DetalleJson.Should().Be("{}");
        }

        [Fact]
        public void ValidadorReglas_WhenIncluirDetalleJsonTrue_ShouldProduceSerializedJson()
        {
            // Arrange
            var sut = new AgValidadorReglas { IncluirDetalleJson = true };
            var contexto = new AgContexto();
            contexto.IdSede = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var bloque = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "07:00", HoraFin = "11:00", Activo = 1 };
            contexto.Bloques.Add(bloque);
            contexto.BloquesPorId[1] = bloque;

            var aula = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            var ind = new AgIndividuo();
            var gen1 = new AgGen { Id = 1, IdAula = 1, IdAsignatura = 10, IdBloqueHorario = 1, HorasGen = 4, TotalHorasAsignatura = 4, FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) } };
            var gen2 = new AgGen { Id = 2, IdAula = 1, IdAsignatura = 20, IdBloqueHorario = 1, HorasGen = 4, TotalHorasAsignatura = 4, FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) } };
            ind.Genes.Add(gen1);
            ind.Genes.Add(gen2);

            // Act
            sut.Evaluar(ind, contexto, incluirDetalleJson: true);

            // Assert
            ind.Violaciones.Should().NotBeEmpty();
            ind.Violaciones[0].DetalleJson.Should().Contain("indicesGenes");
            ind.Violaciones[0].DetalleJson.Should().Contain("intervalos");
        }
    }
}
