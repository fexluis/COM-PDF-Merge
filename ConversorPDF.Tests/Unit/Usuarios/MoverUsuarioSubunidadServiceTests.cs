using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Unit.Usuarios
{
    public class MoverUsuarioSubunidadServiceTests
    {
        private readonly Mock<IUsuarioVaultMover> _vaultMoverMock;
        private readonly Mock<IConfigRepository> _configRepositoryMock;
        private readonly MoverUsuarioSubunidadService _sut;

        public MoverUsuarioSubunidadServiceTests()
        {
            _vaultMoverMock = new Mock<IUsuarioVaultMover>();
            _configRepositoryMock = new Mock<IConfigRepository>();
            _sut = new MoverUsuarioSubunidadService(_vaultMoverMock.Object, _configRepositoryMock.Object);
        }

        [Fact]
        public void Mover_WhenUsuarioNoEsElActual_ShouldNotTouchConfig()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_juan",
                Nombre = "Juan Perez",
                NombreUsuario = "juan",
                Subunidad = "ANTIOQUIA",
                Perfil = "Academica",
                Activo = 1
            };
            const string subunidadDestino = "VALLE";
            const string usuarioActual = "pedro";

            _vaultMoverMock
                .Setup(m => m.Mover(usuario, subunidadDestino))
                .Returns(Result.Success(new MovimientoVaultResultado
                {
                    SubunidadDestino = subunidadDestino,
                    RutaVaultDestino = @"C:\Data\AcademicData_VALLE.db.enc",
                    Insertado = true,
                    Actualizado = false
                }));

            // Act
            var result = _sut.Mover(usuario, subunidadDestino, usuarioActual);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.UsuarioId.Should().Be("usr_juan");
            result.Value.SubunidadDestino.Should().Be("VALLE");
            result.Value.ConfigActualizada.Should().BeFalse();
            result.Value.RequiereReinicio.Should().BeFalse();

            _configRepositoryMock.Verify(c => c.UpdateSubunidad(It.IsAny<string>()), Times.Never);
        }

        [Fact]
        public void Mover_WhenUsuarioEsElActual_ShouldUpdateConfigAndRequireRestart()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_pedro",
                Nombre = "Pedro Gomez",
                NombreUsuario = "pedro",
                Subunidad = "ANTIOQUIA",
                Perfil = "Admin",
                Activo = 1
            };
            const string subunidadDestino = "VALLE";
            const string usuarioActual = " PEDRO ";

            _vaultMoverMock
                .Setup(m => m.Mover(usuario, subunidadDestino))
                .Returns(Result.Success(new MovimientoVaultResultado
                {
                    SubunidadDestino = subunidadDestino,
                    RutaVaultDestino = @"C:\Data\AcademicData_VALLE.db.enc",
                    Insertado = false,
                    Actualizado = true
                }));

            _configRepositoryMock
                .Setup(c => c.UpdateSubunidad(subunidadDestino))
                .Returns(Result.Success());

            // Act
            var result = _sut.Mover(usuario, subunidadDestino, usuarioActual);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.ConfigActualizada.Should().BeTrue();
            result.Value.RequiereReinicio.Should().BeTrue();

            _configRepositoryMock.Verify(c => c.UpdateSubunidad(subunidadDestino), Times.Once);
        }

        [Fact]
        public void Mover_WhenVaultMoverFalla_ShouldPropagateFailureAndNotTouchConfig()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_pedro",
                Nombre = "Pedro Gomez",
                NombreUsuario = "pedro",
                Subunidad = "ANTIOQUIA"
            };
            const string subunidadDestino = "VALLE";
            const string usuarioActual = "pedro";

            _vaultMoverMock
                .Setup(m => m.Mover(usuario, subunidadDestino))
                .Returns(Result.Failure<MovimientoVaultResultado>("Error al acceder al vault destino"));

            // Act
            var result = _sut.Mover(usuario, subunidadDestino, usuarioActual);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Error al acceder al vault destino");
            _configRepositoryMock.Verify(c => c.UpdateSubunidad(It.IsAny<string>()), Times.Never);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Mover_WhenSubunidadDestinoVacia_ShouldFail(string destinoInvalido)
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_juan",
                NombreUsuario = "juan"
            };

            // Act
            var result = _sut.Mover(usuario, destinoInvalido, "pedro");

            // Assert
            result.IsFailure.Should().BeTrue();
            _vaultMoverMock.Verify(m => m.Mover(It.IsAny<Usuario>(), It.IsAny<string>()), Times.Never);
        }

        [Fact]
        public void Mover_WhenUsuarioIdVacio_ShouldFail()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "",
                NombreUsuario = "  "
            };

            // Act
            var result = _sut.Mover(usuario, "VALLE", "pedro");

            // Assert
            result.IsFailure.Should().BeTrue();
            _vaultMoverMock.Verify(m => m.Mover(It.IsAny<Usuario>(), It.IsAny<string>()), Times.Never);
        }
    }
}
