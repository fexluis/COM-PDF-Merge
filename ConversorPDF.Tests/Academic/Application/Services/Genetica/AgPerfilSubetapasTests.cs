using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPerfilSubetapasTests
    {
        [Fact]
        public void TelemetriaSubetapas_WhenActive_ShouldRecordCountsAndTimes()
        {
            // Arrange
            AgTelemetriaSubetapas.Iniciar();
            try
            {
                var contexto = new AgContexto();
                contexto.IdSede = 1;
                contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

                var bloque = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "07:00", HoraFin = "11:00", Activo = 1 };
                contexto.Bloques.Add(bloque);
                contexto.BloquesPorId[1] = bloque;

                var aula = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible };
                contexto.Aulas.Add(aula);
                contexto.AulasPorId[1] = aula;

                contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
                {
                    new TerceroAsignatura { IdTercero = 101, IdAsignatura = 10, Activo = 1, Prioridad = 1 }
                };

                var ind = new AgIndividuo();
                var gen = new AgGen
                {
                    Id = 1,
                    IdAula = 1,
                    IdAsignatura = 10,
                    IdBloqueHorario = 1,
                    IdTercero = 101,
                    HorasSemanales = 4,
                    HorasGen = 4,
                    TotalHorasAsignatura = 48,
                    FechaInicioEstimada = new DateTime(2026, 2, 2),
                    FechaFinEstimada = new DateTime(2026, 5, 30),
                    ClaveGrupo = "G1"
                };
                ind.Genes.Add(gen);

                // Act
                bool disp = AgAsignadorTercero.DisponibleGrupo(101, gen, ind, contexto);
                var cal = AgCalendarioConflictos.Crear(ind.Genes, contexto);
                AgCoherenciaGrupo.Sellar(ind, contexto);

                // Assert
                AgTelemetriaSubetapas.LlamadasDisponibleGrupo.Should().BeGreaterThan(0);
                AgTelemetriaSubetapas.CalendariosCreados.Should().BeGreaterThan(0);
                AgTelemetriaSubetapas.ConteoSellados.Should().BeGreaterThan(0);
                AgTelemetriaSubetapas.TiempoSelladoMs.Should().BeGreaterThanOrEqualTo(0);
                AgTelemetriaSubetapas.TiempoDisponibleGrupoMs.Should().BeGreaterThanOrEqualTo(0);
            }
            finally
            {
                AgTelemetriaSubetapas.Detener();
            }
        }

        [Fact]
        public void TelemetriaSubetapas_WhenInactive_ShouldNotAccumulate()
        {
            // Arrange
            AgTelemetriaSubetapas.Detener();
            AgTelemetriaSubetapas.Reiniciar();

            var contexto = new AgContexto();
            var ind = new AgIndividuo();
            var gen = new AgGen { Id = 1, IdAsignatura = 10 };
            ind.Genes.Add(gen);

            // Act
            bool disp = AgAsignadorTercero.DisponibleGrupo(101, gen, ind, contexto);
            AgCoherenciaGrupo.Sellar(ind, contexto);

            // Assert
            AgTelemetriaSubetapas.LlamadasDisponibleGrupo.Should().Be(0);
            AgTelemetriaSubetapas.ConteoSellados.Should().Be(0);
            AgTelemetriaSubetapas.TiempoSelladoTicks.Should().Be(0);
        }
    }
}
