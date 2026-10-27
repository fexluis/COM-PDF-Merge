using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class TerceroAsignaturaSyncServiceTests
    {
        [Fact]
        public async Task SincronizarDocenteAsignaturaAsync_WhenInvalidIds_ShouldReturnSuccessWithoutAction()
        {
            // Arrange
            var mockRepo = new Mock<ITerceroAsignaturaRepository>();
            var mockConn = new Mock<IDbConnection>();
            var sut = new TerceroAsignaturaSyncService(() => mockConn.Object, mockRepo.Object);

            // Act
            var result = await sut.SincronizarDocenteAsignaturaAsync(0, 1, 0, "I", "SUB", CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            mockRepo.Verify(r => r.UpsertAsync(It.IsAny<TerceroAsignatura>(), It.IsAny<CancellationToken>()), Times.Never);
        }
    }
}
