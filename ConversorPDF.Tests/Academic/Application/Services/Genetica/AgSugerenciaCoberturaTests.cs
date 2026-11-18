using System;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgSugerenciaCoberturaTests
    {
        [Fact]
        public void Texto_WhenCruceAula_ShouldMencionarAula()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("CRUCE_AULA:101");

            // Assert
            result.Should().Contain("aula");
        }

        [Fact]
        public void Texto_WhenCruceDocente_ShouldMencionarDocente()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("CRUCE_DOCENTE:202");

            // Assert
            result.Should().Contain("docente");
        }

        [Fact]
        public void Texto_WhenIncompleta_ShouldMencionarPool()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("COBERTURA_INCOMPLETA:MAT-1");

            // Assert
            result.Should().Contain("pool docente");
        }

        [Fact]
        public void Texto_WhenExcesiva_ShouldMencionarSobreasignacion()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("COBERTURA_EXCESIVA:MAT-1");

            // Assert
            result.Should().Contain("Sobreasignación");
        }

        [Fact]
        public void Texto_WhenEntradaInvalida_ShouldMencionarEntrada()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("COBERTURA_ENTRADA_INVALIDA");

            // Assert
            result.Should().Contain("Entrada de cobertura inválida");
        }

        [Fact]
        public void Texto_WhenDesconocido_ShouldDevolverGenerico()
        {
            // Act
            var result = AgSugerenciaCobertura.Texto("ERROR_DESCONOCIDO_XYZ");

            // Assert
            result.Should().Be("Revise el diagnóstico de cobertura.");
        }
    }
}
