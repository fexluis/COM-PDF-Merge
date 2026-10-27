using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgOperadorMutacionCruceFrozenTests
    {
        [Fact]
        public void AgGen_Clonar_WhenEsFijoTemporalIsTrue_ShouldPreserveEsFijoTemporal()
        {
            // Arrange
            var original = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 5,
                IdBloqueHorario = 15,
                EsFijoTemporal = true
            };

            // Act
            var clon = original.Clonar();

            // Assert
            clon.EsFijoTemporal.Should().BeTrue();
            clon.IdAsignatura.Should().Be(original.IdAsignatura);
            clon.IdAula.Should().Be(original.IdAula);
            clon.IdBloqueHorario.Should().Be(original.IdBloqueHorario);
        }

        [Fact]
        public void AgOperadorMutacion_Mutar_WhenGenIsEsFijoTemporal_ShouldNotMutateAulaOrBloque()
        {
            // Arrange
            var sut = new AgOperadorMutacion();
            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 1, Grupo = "A", Semestre = "I", Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 2, Grupo = "A", Semestre = "I", Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 10, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4.0 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 20, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4.0 });

            var individuo = new AgIndividuo { Id = 100 };
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 5,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdBloqueHorario = 10,
                DiaSemana = 1,
                EsFijoTemporal = true
            };
            individuo.Genes.Add(gen);

            // Act - Mutar con tasa 1.0 (100% de probabilidad de mutación)
            sut.Mutar(individuo, tasaMutacion: 1.0, contexto);

            // Assert - Aula y Bloque deben conservarse idénticos
            individuo.Genes[0].IdAula.Should().Be(1);
            individuo.Genes[0].IdBloqueHorario.Should().Be(10);
        }

        [Fact]
        public void AgOperadorCruce_Cruzar_WhenGeneInChildIsEsFijoTemporal_ShouldPreserveItsBlockAndPattern()
        {
            // Arrange
            var sut = new AgOperadorCruce();
            var contexto = new AgContexto();
            contexto.BloquesPorId[10] = new AgBloqueHorario { Id = 10, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1 };
            contexto.BloquesPorId[99] = new AgBloqueHorario { Id = 99, DiaSemanaNumero = 5, DuracionHoras = 4.0, Activo = 1 };

            var padre1 = new AgIndividuo { Id = 1 };
            padre1.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 1,
                IdBloqueHorario = 10,
                DiaSemana = 2,
                DiasPatron = "2",
                EsFijoTemporal = true,
                Estado = AgEstadoGen.Valido
            });

            var padre2 = new AgIndividuo { Id = 2 };
            padre2.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 10,
                IdAula = 1,
                IdBloqueHorario = 99,
                DiaSemana = 5,
                DiasPatron = "5",
                EsFijoTemporal = false,
                Estado = AgEstadoGen.Valido
            });

            // Act
            var hijos = sut.Cruzar(padre1, padre2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            hijos.Hijo1.Should().NotBeNull();
            hijos.Hijo1.Genes[0].EsFijoTemporal.Should().BeTrue();
            hijos.Hijo1.Genes[0].IdBloqueHorario.Should().Be(10, "el hijo derivado de padre1 debe conservar su bloque fijo");
            hijos.Hijo1.Genes[0].DiaSemana.Should().Be(2);
        }
    }
}
