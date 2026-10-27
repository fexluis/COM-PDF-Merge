using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPromotorVinculacionesTests
    {
        [Fact]
        public async Task PromoverAProduccionAsync_WhenCorridaNotApproved_ShouldReturnFailure()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockSalvedad = new Mock<IAgSalvedadRepository>();
            var mockSync = new Mock<ITerceroAsignaturaSyncService>();
            var mockConn = new Mock<IDbConnection>();

            mockRepo.Setup(r => r.GetCorridaByIdAsync(1, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 1, Estado = AgEstadoCorrida.EnProgreso, IdAula = 10 }));

            var sut = new AgPromotorVinculaciones(() => mockConn.Object, mockRepo.Object, mockSalvedad.Object, mockSync.Object);

            // Act
            var result = await sut.PromoverAProduccionAsync(1, 10, "admin", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().ContainEquivalentOf("aprobada");
        }

        [Fact]
        public async Task PromoverAProduccionAsync_WhenVacantGenesExist_ShouldReturnFailure()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockSalvedad = new Mock<IAgSalvedadRepository>();
            var mockSync = new Mock<ITerceroAsignaturaSyncService>();
            var mockConn = new Mock<IDbConnection>();

            mockRepo.Setup(r => r.GetCorridaByIdAsync(1, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 1, Estado = AgEstadoCorrida.Aprobada, IdAula = 10 }));
            mockRepo.Setup(r => r.GetAulaByIdAsync(10, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgAula { Id = 10, Estado = AgEstadoAula.Aprobado }));

            var individuo = new AgIndividuo { Id = 10 };
            individuo.Genes.Add(new AgGen { Id = 1, EsVacante = 1, IdTercero = 0 });

            mockRepo.Setup(r => r.GetIndividuoByIdAsync(10, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(individuo));

            var sut = new AgPromotorVinculaciones(() => mockConn.Object, mockRepo.Object, mockSalvedad.Object, mockSync.Object);

            // Act
            var result = await sut.PromoverAProduccionAsync(1, 10, "admin", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("vacantes");
        }

        [Fact]
        public async Task PromoverAProduccionAsync_WhenVacantGenesExistAndJustificationProvided_ShouldProceedPastVacancyGate()
        {
            // Arrange
            var mockRepo = new Mock<IAgScheduleRepository>();
            var mockSalvedad = new Mock<IAgSalvedadRepository>();
            var mockSync = new Mock<ITerceroAsignaturaSyncService>();
            var mockConn = new Mock<IDbConnection>();

            mockRepo.Setup(r => r.GetCorridaByIdAsync(1, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgCorrida { Id = 1, Estado = AgEstadoCorrida.Aprobada, IdAula = 10 }));
            mockRepo.Setup(r => r.GetAulaByIdAsync(10, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(new AgAula { Id = 10, Estado = AgEstadoAula.Aprobado }));

            var individuo = new AgIndividuo { Id = 10 };
            individuo.Genes.Add(new AgGen { Id = 1, EsVacante = 1, IdTercero = 0 });

            mockRepo.Setup(r => r.GetIndividuoByIdAsync(10, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Success(individuo));

            mockRepo.Setup(r => r.GetVinculacionesStagingAsync(1, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Failure<IReadOnlyCollection<AgVinculacion>>("Comprobación superada"));

            var sut = new AgPromotorVinculaciones(() => mockConn.Object, mockRepo.Object, mockSalvedad.Object, mockSync.Object);

            // Act
            var result = await sut.PromoverAProduccionAsync(1, 10, "admin", "Aprobado por Consejo Académico", CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("No se encontraron vinculaciones en staging para esta corrida.");
        }
    }
}
