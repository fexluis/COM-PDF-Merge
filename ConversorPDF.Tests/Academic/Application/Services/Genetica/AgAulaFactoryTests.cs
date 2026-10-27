using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgAulaFactoryTests
    {
        [Fact]
        public void CrearDesdeDto_WhenValidDto_ShouldGenerateConsistentAgAulaWithModalidadInAulaRef()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Presencial");
            var ventana = VentanaSemestre.Crear(2, "2026-02", new DateTime(2026, 8, 15), new DateTime(2026, 12, 12), AgOrigenVentana.Periodos).Value;

            // Act
            var aula = AgAulaFactory.CrearDesdeDto(dto, ventana, "VALLE");

            // Assert
            aula.Should().NotBeNull();
            aula.IdPeriodo.Should().Be(2);
            aula.IdSede.Should().Be(5);
            aula.IdPrograma.Should().Be(10);
            aula.Semestre.Should().Be("III");
            aula.Grupo.Should().Be("A");
            aula.ModalidadVinculacion.Should().Be("Presencial");
            aula.EsVirtual.Should().Be(0);
            aula.AulaRef.Should().Be("2|5|10|III|A|PRESENCIAL");
            aula.VentanaOrigen.Should().Be("Periodos");
            aula.FechaInicio.Should().Be("2026-08-15");
            aula.FechaFin.Should().Be("2026-12-12");
            aula.Subunidad.Should().Be("VALLE");
        }

        [Fact]
        public void CrearDesdeDto_WhenVirtualModality_ShouldSetEsVirtualAndDifferentAulaRef()
        {
            // Arrange
            var dtoPresencial = new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Presencial");
            var dtoVirtual = new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Virtual");

            // Act
            var aulaPresencial = AgAulaFactory.CrearDesdeDto(dtoPresencial);
            var aulaVirtual = AgAulaFactory.CrearDesdeDto(dtoVirtual);

            // Assert
            aulaPresencial.AulaRef.Should().NotBe(aulaVirtual.AulaRef);
            aulaVirtual.EsVirtual.Should().Be(1);
            aulaVirtual.TipoAula.Should().Be("VIRTUAL");
            aulaVirtual.Capacidad.Should().Be(999);
            aulaVirtual.AulaRef.Should().Be("2|5|10|III|A|PLATAFORMA");
            aulaVirtual.ModalidadVinculacion.Should().Be("Plataforma");
        }

        [Fact]
        public void CrearDesdeDto_WhenPlataformaModality_ShouldSetEsVirtualAndPlataformaAulaRef()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Plataforma");

            // Act
            var aula = AgAulaFactory.CrearDesdeDto(dto);

            // Assert
            aula.EsVirtual.Should().Be(1);
            aula.TipoAula.Should().Be("VIRTUAL");
            aula.Capacidad.Should().Be(999);
            aula.AulaRef.Should().Be("2|5|10|III|A|PLATAFORMA");
            aula.ModalidadVinculacion.Should().Be("Plataforma");
        }

        [Fact]
        public void CrearDesdeDto_WhenMixtaModality_ShouldSetMixtaAulaRef()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Mixta");

            // Act
            var aula = AgAulaFactory.CrearDesdeDto(dto);

            // Assert
            aula.EsVirtual.Should().Be(0);
            aula.TipoAula.Should().Be("ESTANDAR");
            aula.AulaRef.Should().Be("2|5|10|III|A|MIXTA");
            aula.ModalidadVinculacion.Should().Be("Mixta");
        }

        [Fact]
        public void ConstruirAulaRef_WhenMixedCaseAndSpaces_ShouldFormatInvariantly()
        {
            // Act
            string refKey = AgAulaFactory.ConstruirAulaRef(1, 2, 3, " iii ", " a ", " virtual ");

            // Assert
            refKey.Should().Be("1|2|3|III|A|PLATAFORMA");
        }

        [Fact]
        public void ConstruirAulaRef_WhenMixtaOrSemipresencial_ShouldFormatWithMixta()
        {
            // Act
            string refKeySemi = AgAulaFactory.ConstruirAulaRef(1, 2, 3, "III", "A", "Semipresencial");
            string refKeyMixta = AgAulaFactory.ConstruirAulaRef(1, 2, 3, "III", "A", "Mixta");

            // Assert
            refKeySemi.Should().Be("1|2|3|III|A|MIXTA");
            refKeyMixta.Should().Be("1|2|3|III|A|MIXTA");
        }

        [Fact]
        public void DtoAndAula_WhenEsVirtualSet_ShouldHaveSynchronizedAulaRef()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto(1, 2, 3, "IV", "B", "Presencial")
            {
                EsVirtual = 1
            };

            // Act
            var aula = AgAulaFactory.CrearDesdeDto(dto);

            // Assert
            dto.ModalidadEfectiva.Should().Be("Plataforma");
            dto.AulaRef.Should().Be("1|2|3|IV|B|PLATAFORMA");
            aula.AulaRef.Should().Be(dto.AulaRef);
            aula.ModalidadVinculacion.Should().Be("Plataforma");
        }

        [Fact]
        public void CrearDesdeDto_WhenNullDto_ShouldThrowArgumentNullException()
        {
            // Act & Assert
            Action act = () => AgAulaFactory.CrearDesdeDto(null);
            act.Should().Throw<ArgumentNullException>();
        }

        [Fact]
        public void ExtraerCohorte_WhenValidAulaRef_ShouldReturnCohorteWithoutIdPrograma()
        {
            // Arrange
            string aulaRef = "3|1|2|I|A|PRESENCIAL";

            // Act
            string cohorte = AgAulaFactory.ExtraerCohorte(aulaRef);

            // Assert
            cohorte.Should().Be("3|1|I|A|PRESENCIAL");
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        [InlineData("3|1|2|I|A")] // Menos de 6 partes
        public void ExtraerCohorte_WhenInvalidOrIncompleteAulaRef_ShouldReturnEmptyString(string aulaRefInvalida)
        {
            // Act
            string cohorte = AgAulaFactory.ExtraerCohorte(aulaRefInvalida);

            // Assert
            cohorte.Should().BeEmpty();
        }
    }
}
