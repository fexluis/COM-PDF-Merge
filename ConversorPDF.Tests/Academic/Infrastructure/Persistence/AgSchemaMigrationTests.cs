using System;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgSchemaMigrationTests
    {
        [Fact]
        public void MigrationV48_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV48_AddTercerosAsignaturas();

            // Assert
            mig.Version.Should().Be(48);
            mig.Name.Should().Be("AddTercerosAsignaturas");
        }

        [Fact]
        public void MigrationV49_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV49_AddAgHorariosTables();

            // Assert
            mig.Version.Should().Be(49);
            mig.Name.Should().Be("AddAgHorariosTables");
        }

        [Fact]
        public void MigrationV50_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV50_AddPeriodoVentana();

            // Assert
            mig.Version.Should().Be(50);
            mig.Name.Should().Be("AddPeriodoVentana");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV51_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV51_PeriodoVentanaRepair();

            // Assert
            mig.Version.Should().Be(51);
            mig.Name.Should().Be("PeriodoVentanaRepair");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV52_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV52_IsolateDemoAulas();

            // Assert
            mig.Version.Should().Be(52);
            mig.Name.Should().Be("IsolateDemoAulas");
            mig.Aditiva.Should().BeFalse();
        }

        [Fact]
        public void MigrationV62_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV62_AddAgAulaVentanasYBloquesSemilla();

            // Assert
            mig.Version.Should().Be(62);
            mig.Name.Should().Be("AddAgAulaVentanasYBloquesSemilla");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV63_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV63_UnificarHorariosMilitar();

            // Assert
            mig.Version.Should().Be(63);
            mig.Name.Should().Be("UnificarHorariosMilitar");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV69_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV69_RevalidarVistasAg();

            // Assert
            mig.Version.Should().Be(69);
            mig.Name.Should().Be("RevalidarVistasAg");
            mig.Aditiva.Should().BeTrue();
        }
    }
}
