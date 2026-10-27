using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests
{
    public class AgArbolMapperTests
    {
        [Fact]
        public void MapToNodo_WhenGivenDtoWithVentanaYJornada_ShouldMapPropertiesCorrectly()
        {
            // Arrange
            var dto = new AgContextoSeleccionDto
            {
                IdPeriodo = 1,
                IdSede = 2,
                IdPrograma = 3,
                Semestre = "4",
                Grupo = "A",
                Modalidad = "Presencial",
                Jornada = "TARDE",
                HorasDia = 8,
                DiasVentana = "VIE-DOM"
            };

            // Act
            var nodo = AgArbolMapper.Mapear(dto, idArbol: 100, subunidad: "SUB-01");

            // Assert
            nodo.IdArbol.Should().Be(100);
            nodo.IdPeriodo.Should().Be(1);
            nodo.IdSede.Should().Be(2);
            nodo.IdPrograma.Should().Be(3);
            nodo.Semestre.Should().Be("4");
            nodo.Grupo.Should().Be("A");
            nodo.Modalidad.Should().Be("Presencial");
            nodo.AulaRef.Should().Be("1|2|3|4|A|PRESENCIAL");
            nodo.Jornada.Should().Be("TARDE");
            nodo.HorasDia.Should().Be(8);
            nodo.DiasVentana.Should().Be("VIE-DOM");
            nodo.Subunidad.Should().Be("SUB-01");
        }

        [Fact]
        public void MapToDto_WhenGivenNodoWithVentanaYJornada_ShouldMapPropertiesCorrectly()
        {
            // Arrange
            var nodo = new AgArbolNodo
            {
                Id = 10,
                IdArbol = 100,
                IdPeriodo = 5,
                IdSede = 6,
                IdPrograma = 7,
                Semestre = "2",
                Grupo = "B",
                Modalidad = "Virtual",
                AulaRef = "5|6|7|2|B|VIRTUAL",
                Jornada = "NOCHE",
                HorasDia = 4,
                DiasVentana = "LUN-JUE"
            };

            // Act
            var resultDto = AgArbolMapper.Mapear(nodo);

            // Assert
            resultDto.IdPeriodo.Should().Be(5);
            resultDto.IdSede.Should().Be(6);
            resultDto.IdPrograma.Should().Be(7);
            resultDto.Semestre.Should().Be("2");
            resultDto.Grupo.Should().Be("B");
            resultDto.Modalidad.Should().Be("Plataforma");
            resultDto.AulaRef.Should().Be("5|6|7|2|B|PLATAFORMA");
            resultDto.Jornada.Should().Be("NOCHE");
            resultDto.HorasDia.Should().Be(4);
            resultDto.DiasVentana.Should().Be("LUN-JUE");
        }
    }
}
