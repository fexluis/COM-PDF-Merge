using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class TerceroAsignaturaKeyTests
    {
        [Fact]
        public void ConstruirClave_WhenGivenParameters_ShouldReturnCanonicalPipeDelimitedString()
        {
            // Arrange
            int idTercero = 10;
            int idPrograma = 2;
            int idAsignatura = 105;
            string semestre = "  iv  ";

            // Act
            string clave = TerceroAsignaturaKey.ConstruirClave(idTercero, idPrograma, idAsignatura, semestre);

            // Assert
            clave.Should().Be("10|2|105|IV");
        }

        [Fact]
        public void Calcular_WhenSameInputs_ShouldReturnDeterministicPositiveId()
        {
            // Arrange
            int idTercero = 10;
            int idPrograma = 2;
            int idAsignatura = 105;
            string semestre = "IV";

            // Act
            long key1 = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre);
            long key2 = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre);

            // Assert
            key1.Should().Be(key2);
            key1.Should().BeGreaterThan(0L);
        }

        [Fact]
        public void Calcular_WhenSemestreCaseOrSpacesVary_ShouldReturnIdenticalId()
        {
            // Arrange
            int idTercero = 10;
            int idPrograma = 2;
            int idAsignatura = 105;

            // Act
            long keyUpper = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, "IV");
            long keyLower = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, "iv");
            long keySpaced = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, "  iv  ");

            // Assert
            keyUpper.Should().Be(keyLower);
            keyUpper.Should().Be(keySpaced);
        }

        [Fact]
        public void Calcular_WhenSemestreIsNullOrEmpty_ShouldReturnIdenticalId()
        {
            // Arrange
            int idTercero = 15;
            int idPrograma = 1;
            int idAsignatura = 50;

            // Act
            long keyNull = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, null);
            long keyEmpty = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, "");
            long keyWhitespace = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, "   ");

            // Assert
            keyNull.Should().Be(keyEmpty);
            keyNull.Should().Be(keyWhitespace);
            keyNull.Should().BeGreaterThan(0L);
        }

        [Fact]
        public void Calcular_WhenDistinctTuples_ShouldProduceDistinctKeysWithoutCollisions()
        {
            // Arrange
            var keys = new HashSet<long>();
            int[] terceros = { 1, 2, 3, 100, 500 };
            int[] programas = { 1, 2, 5 };
            int[] asignaturas = { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 };
            string[] semestres = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "" };

            // Act & Assert
            foreach (var t in terceros)
            {
                foreach (var p in programas)
                {
                    foreach (var a in asignaturas)
                    {
                        foreach (var s in semestres)
                        {
                            long key = TerceroAsignaturaKey.Calcular(t, p, a, s);
                            key.Should().BeGreaterThan(0L, "clave de 63 bits debe ser siempre positiva");
                            keys.Add(key).Should().BeTrue("no deben existir colisiones entre tuplas distintas");
                        }
                    }
                }
            }
        }
    }
}
