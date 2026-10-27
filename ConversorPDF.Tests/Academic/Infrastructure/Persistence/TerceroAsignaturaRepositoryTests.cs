using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class TerceroAsignaturaRepositoryTests
    {
        [Fact]
        public async Task UpsertAsync_WhenItemIsNull_ShouldReturnFailure()
        {
            // Arrange
            var mockConn = new Mock<IDbConnection>();
            var sut = new TerceroAsignaturaRepository(() => mockConn.Object);

            // Act
            var result = await sut.UpsertAsync(null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("nulo");
        }

        [Fact]
        public void Upsert_WhenItemIsNull_ShouldReturnFailure()
        {
            // Arrange
            var mockConn = new Mock<IDbConnection>();
            var sut = new TerceroAsignaturaRepository(() => mockConn.Object);

            // Act
            var result = sut.Upsert(null);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("nulo");
        }
    }
}
