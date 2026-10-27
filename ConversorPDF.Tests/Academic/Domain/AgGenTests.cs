using System;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgGenTests
    {
        [Fact]
        public void AsignarBloque_WhenBloqueProvided_ShouldSynchronizeIdBloqueDiaSemanaAndHorasSemanales()
        {
            // Arrange
            var gen = new AgGen();
            var bloque = new AgBloqueHorario
            {
                Id = 15,
                DiaSemanaNumero = 6, // Sábado
                DuracionHoras = 4.5
            };

            // Act
            gen.AsignarBloque(bloque);

            // Assert
            gen.IdBloqueHorario.Should().Be(15);
            gen.DiaSemana.Should().Be(6);
            gen.HorasSemanales.Should().Be(4.5);
        }

        [Fact]
        public void AsignarBloque_WhenNull_ShouldThrowArgumentNullException()
        {
            // Arrange
            var gen = new AgGen();

            // Act & Assert
            Action act = () => gen.AsignarBloque(null);
            act.Should().Throw<ArgumentNullException>();
        }

        [Fact]
        public void AsignarAula_WhenAulaProvided_ShouldSynchronizeAulaAttributes()
        {
            // Arrange
            var gen = new AgGen();
            var aula = new AgAula
            {
                Id = 42,
                IdPeriodo = 2,
                IdSede = 5,
                IdPrograma = 10,
                Semestre = "III",
                Grupo = "B"
            };

            // Act
            gen.AsignarAula(aula);

            // Assert
            gen.IdAula.Should().Be(42);
            gen.IdPeriodo.Should().Be(2);
            gen.IdSede.Should().Be(5);
            gen.IdPrograma.Should().Be(10);
            gen.Semestre.Should().Be("III");
            gen.Grupo.Should().Be("B");
        }

        [Fact]
        public void AsignarAula_WhenNull_ShouldThrowArgumentNullException()
        {
            // Arrange
            var gen = new AgGen();

            // Act & Assert
            Action act = () => gen.AsignarAula(null);
            act.Should().Throw<ArgumentNullException>();
        }

        [Fact]
        public void HorasGen_Negativo_SeClampaACero()
        {
            // Arrange (Fase 583)
            var gen = new AgGen();

            // Act
            gen.HorasGen = -16.0;

            // Assert
            gen.HorasGen.Should().Be(0.0);
        }
    }
}

