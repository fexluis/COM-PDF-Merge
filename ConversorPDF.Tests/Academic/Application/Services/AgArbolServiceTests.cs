using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgArbolServiceTests
    {
        private readonly Mock<IAgArbolRepository> _repositoryMock;
        private readonly AgArbolService _sut;

        public AgArbolServiceTests()
        {
            _repositoryMock = new Mock<IAgArbolRepository>();
            _repositoryMock.Setup(r => r.ExisteNombreAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success(false));
            _sut = new AgArbolService(_repositoryMock.Object);
        }

        [Fact]
        public async Task GuardarPerfilAsync_WhenNombreIsEmpty_ShouldReturnFailure()
        {
            // Arrange
            var selecciones = new List<AgContextoSeleccionDto>
            {
                new AgContextoSeleccionDto(1, 1, 1, "I", "A", "Presencial")
            };

            // Act
            var result = await _sut.GuardarPerfilAsync(
                "", "desc", 1, null, false, 50, 200, 0.85, 0.10, 2, selecciones, "user", "VALLE", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("nombre");
        }

        [Fact]
        public async Task GuardarPerfilAsync_WhenSeleccionesIsEmpty_ShouldReturnFailure()
        {
            // Act
            var result = await _sut.GuardarPerfilAsync(
                "Mi Perfil", "desc", 1, null, false, 50, 200, 0.85, 0.10, 2, new List<AgContextoSeleccionDto>(), "user", "VALLE", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("al menos una combinación");
        }

        [Fact]
        public async Task GuardarPerfilAsync_WhenValidInput_ShouldDelegateToRepository()
        {
            // Arrange
            var selecciones = new List<AgContextoSeleccionDto>
            {
                new AgContextoSeleccionDto(2, 5, 10, "III", "A", "Presencial")
                {
                    Capacidad = 40,
                    Codigo = "ING-III-A"
                }
            };

            _repositoryMock.Setup(r => r.GuardarAsync(It.IsAny<AgArbol>(), It.IsAny<IReadOnlyCollection<AgArbolNodo>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<long>(42));

            // Act
            var result = await _sut.GuardarPerfilAsync(
                "Perfil Valido", "desc", 2, null, false, 60, 220, 0.88, 0.12, 3, selecciones, "user", "VALLE", CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(42);

            _repositoryMock.Verify(r => r.GuardarAsync(
                It.Is<AgArbol>(a => a.Nombre == "Perfil Valido" && a.Poblacion == 60),
                It.Is<IReadOnlyCollection<AgArbolNodo>>(n => n.Count == 1),
                It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task CargarPerfilAsync_WhenValidId_ShouldReconstructDtosWithAttributes()
        {
            // Arrange
            long id = 10;
            var perfil = new AgArbol { Id = id, Nombre = "Perfil Existente", IdPeriodo = 2 };
            var nodos = new List<AgArbolNodo>
            {
                new AgArbolNodo
                {
                    IdArbol = id,
                    IdPeriodo = 2,
                    IdSede = 3,
                    IdPrograma = 4,
                    Semestre = "IV",
                    Grupo = "B",
                    Modalidad = "Virtual",
                    AulaRef = "2|3|4|IV|B|VIRTUAL",
                    Capacidad = 999,
                    TipoAula = "VIRTUAL",
                    EsVirtual = 1
                }
            };

            _repositoryMock.Setup(r => r.ObtenerAsync(id, It.IsAny<CancellationToken>())).ReturnsAsync(Result.Success(perfil));
            _repositoryMock.Setup(r => r.ObtenerNodosAsync(id, It.IsAny<CancellationToken>())).ReturnsAsync(Result.Success<IReadOnlyCollection<AgArbolNodo>>(nodos));

            // Act
            var result = await _sut.CargarPerfilAsync(id, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Item1.Nombre.Should().Be("Perfil Existente");
            result.Value.Item2.Should().HaveCount(1);

            var dto = result.Value.Item2.First();
            dto.IdPeriodo.Should().Be(2);
            dto.IdSede.Should().Be(3);
            dto.IdPrograma.Should().Be(4);
            dto.Semestre.Should().Be("IV");
            dto.Grupo.Should().Be("B");
            dto.ModalidadEfectiva.Should().Be("Plataforma");
            dto.AulaRef.Should().Be("2|3|4|IV|B|PLATAFORMA");
            dto.EsVirtual.Should().Be(1);
        }

        [Fact]
        public async Task GuardarPerfilAsync_WhenDuplicateSelectionsProvided_ShouldDeduplicateNodesByAulaRef()
        {
            // Arrange - Dos DTOs que producen el mismo AulaRef
            var selecciones = new List<AgContextoSeleccionDto>
            {
                new AgContextoSeleccionDto(1, 2, 3, "I", "A", "Presencial"),
                new AgContextoSeleccionDto(1, 2, 3, "I", "A", "Presencial")
            };

            _repositoryMock.Setup(r => r.GuardarAsync(It.IsAny<AgArbol>(), It.IsAny<IReadOnlyCollection<AgArbolNodo>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<long>(1));

            // Act
            var result = await _sut.GuardarPerfilAsync(
                "Perfil Dedupe", "desc", 1, null, false, 50, 200, 0.85, 0.10, 2, selecciones, "user", "VALLE", CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            _repositoryMock.Verify(r => r.GuardarAsync(
                It.IsAny<AgArbol>(),
                It.Is<IReadOnlyCollection<AgArbolNodo>>(n => n.Count == 1),
                It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task GuardarPerfilAsync_WhenExisteNombreAsyncFails_ShouldReturnFailure()
        {
            // Arrange
            var selecciones = new List<AgContextoSeleccionDto>
            {
                new AgContextoSeleccionDto(1, 1, 1, "I", "A", "Presencial")
            };

            _repositoryMock.Setup(r => r.ExisteNombreAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Failure<bool>("Fallo en BD al validar nombre"));

            // Act
            var result = await _sut.GuardarPerfilAsync(
                "Mi Perfil", "desc", 1, null, false, 50, 200, 0.85, 0.10, 2, selecciones, "user", "VALLE", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("Fallo en BD");
        }
    }
}
