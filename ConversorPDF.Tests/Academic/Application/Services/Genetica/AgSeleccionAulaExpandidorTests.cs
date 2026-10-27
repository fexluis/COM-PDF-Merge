using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgSeleccionAulaExpandidorTests
    {
        [Fact]
        public void ValidarYExpandir_WhenDuplicatesPresent_ShouldDeduplicateByAulaRef()
        {
            // Arrange
            var sut = new AgSeleccionAulaExpandidor();
            var item1 = new AgContextoSeleccionDto(1, 2, 3, "I", "A", "Presencial");
            var item2 = new AgContextoSeleccionDto(1, 2, 3, "I", "A", "Presencial"); // duplicado
            var item3 = new AgContextoSeleccionDto(1, 2, 3, "I", "B", "Presencial");

            // Act
            var res = sut.ValidarYExpandir(new[] { item1, item2, item3 });

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Should().HaveCount(2);
        }

        [Fact]
        public void ValidarYExpandir_WhenExceedsCap_ShouldReturnFailure()
        {
            // Arrange
            var sut = new AgSeleccionAulaExpandidor();
            var items = new List<AgContextoSeleccionDto>();
            for (int i = 0; i < 505; i++)
            {
                items.Add(new AgContextoSeleccionDto(1, 2, i + 1, "I", "A", "Presencial"));
            }

            // Act
            var res = sut.ValidarYExpandir(items, capMaximo: 500);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("superando el límite de seguridad");
        }

        [Fact]
        public void ValidarYExpandir_WhenEmpty_ShouldReturnFailure()
        {
            // Arrange
            var sut = new AgSeleccionAulaExpandidor();

            // Act
            var res = sut.ValidarYExpandir(new AgContextoSeleccionDto[0]);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("No se seleccionó ninguna combinación");
        }
    }
}
