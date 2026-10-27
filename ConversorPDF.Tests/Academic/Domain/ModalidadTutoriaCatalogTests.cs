using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class ModalidadTutoriaCatalogTests
    {
        [Fact]
        public void Opciones_ShouldContainCanonicOptionsInOrder()
        {
            // Assert
            ModalidadTutoriaCatalog.Opciones.Should().ContainInOrder("Presencial", "Plataforma", "Mixta");
            ModalidadTutoriaCatalog.Opciones.Count.Should().Be(3);
        }

        [Fact]
        public void Pares_ShouldMatchCanonicCodesAndLabels()
        {
            // Assert
            ModalidadTutoriaCatalog.Pares.Should().HaveCount(3);
            ModalidadTutoriaCatalog.Pares[0].Codigo.Should().Be("Presencial");
            ModalidadTutoriaCatalog.Pares[0].Etiqueta.Should().Be("Presencial");
            ModalidadTutoriaCatalog.Pares[1].Codigo.Should().Be("Plataforma");
            ModalidadTutoriaCatalog.Pares[1].Etiqueta.Should().Be("Plataforma");
            ModalidadTutoriaCatalog.Pares[2].Codigo.Should().Be("Mixta");
            ModalidadTutoriaCatalog.Pares[2].Etiqueta.Should().Be("Mixta");
        }

        [Theory]
        [InlineData("Plataforma", "Plataforma")]
        [InlineData("plataforma", "Plataforma")]
        [InlineData("Virtual", "Plataforma")]
        [InlineData("VIRTUAL", "Plataforma")]
        [InlineData(" virtual ", "Plataforma")]
        [InlineData("En línea", "Plataforma")]
        [InlineData("EN LINEA", "Plataforma")]
        [InlineData("E-Learning", "Plataforma")]
        [InlineData("Mixta", "Mixta")]
        [InlineData("mixta", "Mixta")]
        [InlineData("Semipresencial", "Mixta")]
        [InlineData("SEMIPRESENCIAL", "Mixta")]
        [InlineData("SEMI PRESENCIAL", "Mixta")]
        [InlineData("A Distancia", "Mixta")]
        [InlineData("distancia", "Mixta")]
        [InlineData("Presencial", "Presencial")]
        [InlineData("presencial", "Presencial")]
        [InlineData("PRESENCIAL", "Presencial")]
        [InlineData(null, "Presencial")]
        [InlineData("", "Presencial")]
        [InlineData("   ", "Presencial")]
        [InlineData("OtraModalidadDesconocida", "OtraModalidadDesconocida")]
        public void NormalizarTutoria_ShouldConvertVariantsToCanonicCode(string entrada, string esperado)
        {
            // Act
            string resultado = ModalidadTutoriaCatalog.NormalizarTutoria(entrada);

            // Assert
            resultado.Should().Be(esperado);
        }

        [Theory]
        [InlineData("Presencial", "Presencial")]
        [InlineData("Plataforma", "Plataforma")]
        [InlineData("Mixta", "Mixta")]
        [InlineData("Virtual", "Plataforma")]
        [InlineData("Semipresencial", "Mixta")]
        [InlineData("A Distancia", "Mixta")]
        [InlineData(null, "")]
        [InlineData("", "")]
        [InlineData("   ", "")]
        [InlineData("OTRA", "OTRA")]
        public void AEtiqueta_ShouldReturnCorrectLabel(string codigo, string esperado)
        {
            // Act
            string resultado = ModalidadTutoriaCatalog.AEtiqueta(codigo);

            // Assert
            resultado.Should().Be(esperado);
        }

        [Fact]
        public void OrderOfEvaluation_SemipresencialMustBeMixtaNotPresencial()
        {
            // Assert
            ModalidadTutoriaCatalog.EsMixta("Semipresencial").Should().BeTrue();
            ModalidadTutoriaCatalog.EsPresencial("Semipresencial").Should().BeFalse();
            ModalidadTutoriaCatalog.EsPlataforma("Semipresencial").Should().BeFalse();
        }

        [Fact]
        public void EsPlataforma_WhenVirtualOrPlataforma_ShouldReturnTrue()
        {
            // Assert
            ModalidadTutoriaCatalog.EsPlataforma("Virtual").Should().BeTrue();
            ModalidadTutoriaCatalog.EsPlataforma("Plataforma").Should().BeTrue();
            ModalidadTutoriaCatalog.EsPlataforma("Presencial").Should().BeFalse();
            ModalidadTutoriaCatalog.EsPlataforma("Mixta").Should().BeFalse();
        }

        [Fact]
        public void EsPresencial_WhenPresencial_ShouldReturnTrue()
        {
            // Assert
            ModalidadTutoriaCatalog.EsPresencial("Presencial").Should().BeTrue();
            ModalidadTutoriaCatalog.EsPresencial("PRESENCIAL").Should().BeTrue();
            ModalidadTutoriaCatalog.EsPresencial("Plataforma").Should().BeFalse();
            ModalidadTutoriaCatalog.EsPresencial("Virtual").Should().BeFalse();
        }
    }
}
