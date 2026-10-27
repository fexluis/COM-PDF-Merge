using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Mappers
{
    public class AgGenToAgVinculacionMapperTests
    {
        [Fact]
        public void Map_WhenGenProvided_ShouldMapPropertiesCorrectly()
        {
            // Arrange
            var gen = new AgGen
            {
                IdTercero = 101,
                IdPrograma = 5,
                IdSede = 2,
                IdAsignatura = 12,
                IdPeriodo = 20261,
                Semestre = "III",
                Grupo = "B",
                TotalHorasAsignatura = 32.0,
                FechaInicioEstimada = new DateTime(2026, 3, 1),
                FechaFinEstimada = new DateTime(2026, 4, 15)
            };

            // Act
            var result = AgGenToAgVinculacionMapper.Map(gen, idCorrida: 99, idIndividuo: 88, subunidad: "SUB-MED");

            // Assert
            result.Should().NotBeNull();
            result.IdCorrida.Should().Be(99);
            result.IdIndividuo.Should().Be(88);
            result.IdTercero.Should().Be(101);
            result.IdPrograma.Should().Be(5);
            result.Idsede.Should().Be(2);
            result.IdAsignatura.Should().Be(12);
            result.IdPeriodo.Should().Be(20261);
            result.Semestre.Should().Be("III");
            result.Grupo.Should().Be("B");
            result.TotalHoras.Should().Be(32.0);
            result.FechaInicio.Should().Be("2026-03-01");
            result.FechaFin.Should().Be("2026-04-15");
            result.Estado.Should().Be(AgEstadoVinculacion.Borrador);
            result.Subunidad.Should().Be("SUB-MED");
        }

        [Fact]
        public void Map_WhenGenNull_ShouldReturnNull()
        {
            // Act
            var result = AgGenToAgVinculacionMapper.Map(null, 1, 1, "SUB");

            // Assert
            result.Should().BeNull();
        }

        [Fact]
        public void MapGrupo_ConOleada_ConservaFechasDelGen()
        {
            // Arrange (Fase 1108 / Fase 1112): El gen posee sub-ventana de oleada (2026-03-01..2026-04-15)
            // dentro de la ventana semestral (2026-02-01..2026-06-30 con IdPeriodo > 0).
            // MapGrupo debe conservar las fechas de la oleada y no sustituirlas por las del semestre.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(100, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var genes = new List<AgGen>
            {
                new AgGen
                {
                    Id = 1,
                    IdAsignatura = 10,
                    IdTercero = 5,
                    FechaInicioEstimada = new DateTime(2026, 3, 1),
                    FechaFinEstimada = new DateTime(2026, 4, 15),
                    TotalHorasAsignatura = 32.0,
                    HorasGen = 32.0
                }
            };

            // Act
            var result = AgGenToAgVinculacionMapper.MapGrupo(genes, idCorrida: 1, idIndividuo: 1, contexto: contexto);

            // Assert
            result.Should().NotBeNull();
            result.FechaInicio.Should().Be("2026-03-01");
            result.FechaFin.Should().Be("2026-04-15");
            result.IdPeriodo.Should().Be(100);
        }

        [Fact]
        public void MapGrupo_GenesConFechasDistintas_UsaMinMax()
        {
            // Arrange (Fase 1109 / Fase 1112): Grupo multi-día con genes en fechas desfasadas.
            // MapGrupo debe tomar Min(FechaInicio) y Max(FechaFin).
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(100, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var genes = new List<AgGen>
            {
                new AgGen
                {
                    Id = 1,
                    IdAsignatura = 10,
                    FechaInicioEstimada = new DateTime(2026, 3, 1),
                    FechaFinEstimada = new DateTime(2026, 4, 10),
                    HorasGen = 16.0
                },
                new AgGen
                {
                    Id = 2,
                    IdAsignatura = 10,
                    FechaInicioEstimada = new DateTime(2026, 3, 5),
                    FechaFinEstimada = new DateTime(2026, 4, 15),
                    HorasGen = 16.0
                }
            };

            // Act
            var result = AgGenToAgVinculacionMapper.MapGrupo(genes, 1, 1, contexto);

            // Assert
            result.Should().NotBeNull();
            result.FechaInicio.Should().Be("2026-03-01");
            result.FechaFin.Should().Be("2026-04-15");
        }
    }
}
