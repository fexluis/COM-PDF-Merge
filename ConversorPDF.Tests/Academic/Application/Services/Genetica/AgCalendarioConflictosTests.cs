using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgCalendarioConflictosTests
    {
        [Fact]
        public void Crear_WhenGenesAndContextAreValid_ShouldReturnProjectedOccupations()
        {
            // Arrange
            var contexto = CrearContextoValido();
            var gen = new AgGen
            {
                Id = 1,
                IdIndividuo = 10,
                IdAula = 1,
                IdAsignatura = 2,
                IdBloqueHorario = 1,
                HorasGen = 8,
                TotalHorasAsignatura = 8,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 9)
            };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(2);

            var oc1 = result.Value[0];
            oc1.Gen.Should().BeSameAs(gen);
            oc1.Fecha.Should().Be(new DateTime(2026, 2, 2));
            oc1.Inicio.Should().Be(new TimeSpan(8, 0, 0));
            oc1.Fin.Should().Be(new TimeSpan(12, 0, 0));
            oc1.Franja.Should().Be("MANANA");

            var oc2 = result.Value[1];
            oc2.Gen.Should().BeSameAs(gen);
            oc2.Fecha.Should().Be(new DateTime(2026, 2, 9));
            oc2.Inicio.Should().Be(new TimeSpan(8, 0, 0));
            oc2.Fin.Should().Be(new TimeSpan(12, 0, 0));
            oc2.Franja.Should().Be("MANANA");
        }

        [Fact]
        public void Crear_WhenGenHasZeroOrNegativeHours_ShouldSkipGen()
        {
            // Arrange
            var contexto = CrearContextoValido();
            var gen1 = new AgGen { Id = 1, HorasGen = 0, IdBloqueHorario = 999 };
            var gen2 = new AgGen { Id = 2, HorasGen = -4, IdBloqueHorario = 999 };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen1, gen2 }, contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }

        [Fact]
        public void Crear_WhenBloqueNotExists_ShouldReturnBloqueInvalidoFailure()
        {
            // Arrange
            var contexto = CrearContextoValido();
            var gen = new AgGen
            {
                Id = 1,
                IdAula = 5,
                IdAsignatura = 8,
                IdBloqueHorario = 999,
                HorasGen = 4
            };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I6_BLOQUE_INVALIDO: aula=5, asig=8");
        }

        [Fact]
        public void Crear_WhenBloqueNotActive_ShouldReturnBloqueInvalidoFailure()
        {
            // Arrange
            var contexto = CrearContextoValido();
            contexto.BloquesPorId[2] = new AgBloqueHorario
            {
                Id = 2,
                Activo = 0,
                HoraInicio = "08:00",
                HoraFin = "12:00"
            };

            var gen = new AgGen
            {
                Id = 1,
                IdAula = 3,
                IdAsignatura = 4,
                IdBloqueHorario = 2,
                HorasGen = 4
            };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I6_BLOQUE_INVALIDO: aula=3, asig=4");
        }

        [Theory]
        [InlineData("12:00", "08:00")]
        [InlineData("08:00", "08:00")]
        [InlineData("invalido", "12:00")]
        [InlineData("08:00", "invalido")]
        public void Crear_WhenBloqueTimesInvalid_ShouldReturnBloqueInvalidoFailure(string hIni, string hFin)
        {
            // Arrange
            var contexto = CrearContextoValido();
            contexto.BloquesPorId[3] = new AgBloqueHorario
            {
                Id = 3,
                Activo = 1,
                HoraInicio = hIni,
                HoraFin = hFin
            };

            var gen = new AgGen
            {
                Id = 1,
                IdAula = 10,
                IdAsignatura = 20,
                IdBloqueHorario = 3,
                HorasGen = 4
            };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I6_BLOQUE_INVALIDO: aula=10, asig=20");
        }

        [Fact]
        public void Crear_WhenExpansorFails_ShouldPropagateFailure()
        {
            // Arrange
            var contexto = new AgContexto();
            var bloque = new AgBloqueHorario { Id = 1, Activo = 1, HoraInicio = "08:00", HoraFin = "12:00" };
            contexto.BloquesPorId[1] = bloque;

            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                HorasGen = 4
            };

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Ventana semestral no resuelta");
        }

        [Fact]
        public void Crear_WhenSesionHasNoFecha_ShouldReturnSesionSinFechaFailure()
        {
            // Arrange
            var contexto = CrearContextoValido();
            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                HorasGen = 4
            };

            var mockExpansor = new Mock<IAgExpansorSesiones>();
            var sesionSinFecha = new AgSesion { FechaSesion = null };
            mockExpansor.Setup(e => e.Expandir(gen, 0, contexto))
                .Returns(Result.Success(new List<AgSesion> { sesionSinFecha }));

            // Act
            var result = AgCalendarioConflictos.Crear(new[] { gen }, contexto, mockExpansor.Object);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_SESION_SIN_FECHA");
        }

        [Fact]
        public void Crear_WhenArgumentsNull_ShouldThrowArgumentNullException()
        {
            // Arrange
            var contexto = CrearContextoValido();

            // Act & Assert
            Action act1 = () => AgCalendarioConflictos.Crear(null, contexto);
            act1.Should().Throw<ArgumentNullException>();

            Action act2 = () => AgCalendarioConflictos.Crear(new AgGen[0], null);
            act2.Should().Throw<ArgumentNullException>();
        }

        [Fact]
        public void SeCruzan_WhenSameDateAndOverlappingIntervals_ShouldReturnTrue()
        {
            // Arrange
            var a = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(8, 0, 0),
                Fin = new TimeSpan(12, 0, 0)
            };
            var b = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(10, 0, 0),
                Fin = new TimeSpan(14, 0, 0)
            };

            // Act
            bool cruzados = AgCalendarioConflictos.SeCruzan(a, b);

            // Assert
            cruzados.Should().BeTrue();
        }

        [Fact]
        public void SeCruzan_WhenSameDateAndDisjointOrAdjacentIntervals_ShouldReturnFalse()
        {
            // Arrange
            var a = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(8, 0, 0),
                Fin = new TimeSpan(12, 0, 0)
            };
            var adyacente = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(12, 0, 0),
                Fin = new TimeSpan(16, 0, 0)
            };
            var disjunto = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(14, 0, 0),
                Fin = new TimeSpan(18, 0, 0)
            };

            // Act & Assert
            AgCalendarioConflictos.SeCruzan(a, adyacente).Should().BeFalse();
            AgCalendarioConflictos.SeCruzan(adyacente, a).Should().BeFalse();
            AgCalendarioConflictos.SeCruzan(a, disjunto).Should().BeFalse();
            AgCalendarioConflictos.SeCruzan(disjunto, a).Should().BeFalse();
        }

        [Fact]
        public void SeCruzan_WhenDifferentDatesAndSameInterval_ShouldReturnFalse()
        {
            // Arrange
            var a = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(8, 0, 0),
                Fin = new TimeSpan(12, 0, 0)
            };
            var b = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 2),
                Inicio = new TimeSpan(8, 0, 0),
                Fin = new TimeSpan(12, 0, 0)
            };

            // Act
            bool cruzados = AgCalendarioConflictos.SeCruzan(a, b);

            // Assert
            cruzados.Should().BeFalse();
        }

        [Fact]
        public void SeCruzan_WhenEitherOccupationIsNull_ShouldReturnFalse()
        {
            // Arrange
            var a = new AgCalendarioConflictos.Ocupacion
            {
                Fecha = new DateTime(2026, 3, 1),
                Inicio = new TimeSpan(8, 0, 0),
                Fin = new TimeSpan(12, 0, 0)
            };

            // Act & Assert
            AgCalendarioConflictos.SeCruzan(a, null).Should().BeFalse();
            AgCalendarioConflictos.SeCruzan(null, a).Should().BeFalse();
            AgCalendarioConflictos.SeCruzan(null, null).Should().BeFalse();
        }

        [Theory]
        [InlineData(1, 10, "3", "A", "1|10|III|A")]
        [InlineData(2, 20, "VI", "B", "2|20|VI|B")]
        [InlineData(3, 30, "SIN ASIGNAR", " C ", "3|30|SIN ASIGNAR|C")]
        [InlineData(4, 40, null, "D", "4|40||D")]
        [InlineData(5, 50, "  ", null, "5|50||")]
        public void Cohorte_WhenEvaluated_ShouldReturnExpectedFormat(
            int idSede, int idPrograma, string semestre, string grupo, string expected)
        {
            // Arrange
            var gen = new AgGen
            {
                IdSede = idSede,
                IdPrograma = idPrograma,
                Semestre = semestre,
                Grupo = grupo
            };

            // Act
            string cohorte = AgCalendarioConflictos.Cohorte(gen);

            // Assert
            cohorte.Should().Be(expected);
        }

        [Fact]
        public void Cohorte_WhenGenIsNull_ShouldReturnEmptyString()
        {
            // Act
            string cohorte = AgCalendarioConflictos.Cohorte(null);

            // Assert
            cohorte.Should().Be(string.Empty);
        }

        [Theory]
        [InlineData(0, 100, 999, true)]
        [InlineData(1, 100, 999, false)]
        [InlineData(0, 0, 999, false)]
        [InlineData(0, -1, 999, false)]
        [InlineData(0, 999, 999, false)]
        public void DocenteReal_WhenEvaluated_ShouldReturnExpectedResult(
            int esVacante, int idTercero, int idTerceroVacante, bool expected)
        {
            // Arrange
            var gen = new AgGen
            {
                EsVacante = esVacante,
                IdTercero = idTercero
            };
            var contexto = new AgContexto
            {
                IdTerceroVacante = idTerceroVacante
            };

            // Act
            bool result = AgCalendarioConflictos.DocenteReal(gen, contexto);

            // Assert
            result.Should().Be(expected);
        }

        [Fact]
        public void DocenteReal_WhenGenOrContextIsNull_ShouldReturnFalse()
        {
            // Arrange
            var gen = new AgGen { EsVacante = 0, IdTercero = 100 };
            var contexto = new AgContexto { IdTerceroVacante = 999 };

            // Act & Assert
            AgCalendarioConflictos.DocenteReal(null, contexto).Should().BeFalse();
            AgCalendarioConflictos.DocenteReal(gen, null).Should().BeFalse();
            AgCalendarioConflictos.DocenteReal(null, null).Should().BeFalse();
        }

        private static AgContexto CrearContextoValido()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(
                1,
                "2026-1",
                new DateTime(2026, 2, 1),
                new DateTime(2026, 6, 30),
                AgOrigenVentana.Presencial).Value;

            var bloque = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 1,
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4,
                Franja = "MANANA",
                Activo = 1
            };
            contexto.Bloques.Add(bloque);
            contexto.AulasPorId[1] = new AgAula { Id = 1, ModalidadVinculacion = "Presencial" };
            contexto.ConstruirIndices();
            return contexto;
        }
    }
}
