using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgAsignadorTerceroTests
    {
        private readonly AgAsignadorTercero _sut = new AgAsignadorTercero(new Random(42));

        [Fact]
        public void AsignarDocente_WhenNoCandidatesInPool_ShouldMarkAsVacanteSinCandidatos()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 1 };
            var contexto = new AgContexto();

            // Act
            bool asignado = _sut.AsignarDocente(gen, contexto);

            // Assert
            asignado.Should().BeFalse();
            gen.EsVacante.Should().Be(1);
            gen.IdTercero.Should().Be(0);
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_CANDIDATOS.Codigo());
        }

        [Fact]
        public void AsignarDocente_WhenCandidateAvailableInPool_ShouldAssignTeacher()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 5 };
            var contexto = new AgContexto();
            contexto.PoolPorAsignatura[5] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 25, Activo = 1, Prioridad = 1 }
            };

            // Act
            bool asignado = _sut.AsignarDocente(gen, contexto);

            // Assert
            asignado.Should().BeTrue();
            gen.IdTercero.Should().Be(25);
            gen.EsVacante.Should().Be(0);
            gen.MotivoVacante.Should().BeEmpty();
        }

        [Fact]
        public void DiagnosticarMotivo_WhenNoPoolExists_ShouldReturnSinCandidatos()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 99 };
            var contexto = new AgContexto();

            // Act
            var motivo = _sut.DiagnosticarMotivo(gen, contexto);

            // Assert
            motivo.Should().Be(AgMotivoVacante.SIN_CANDIDATOS.Codigo());
        }
    }
}
