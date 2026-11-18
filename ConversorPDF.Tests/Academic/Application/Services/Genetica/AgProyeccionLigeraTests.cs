using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgProyeccionLigeraTests
    {
        private AgContexto CrearContexto()
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

            return contexto;
        }

        [Fact]
        public void ProyectarFechas_WhenCalled_ShouldMatchExpandirDatesExactly()
        {
            // Arrange
            var contexto = CrearContexto();
            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                HorasGen = 8,
                TotalHorasAsignatura = 48,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 5, 30),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2), new DateTime(2026, 2, 9) }
            };

            var expansor = new AgExpansorSesiones();

            // Act
            var resFechas = AgExpansorSesiones.ProyectarFechas(gen, contexto);
            var resSesiones = expansor.Expandir(gen, 0, contexto);

            // Assert
            resFechas.IsSuccess.Should().BeTrue();
            resSesiones.IsSuccess.Should().BeTrue();

            var fechasLigeras = resFechas.Value;
            var fechasCompletas = resSesiones.Value.Select(s => s.FechaSesion.Value.Date).ToList();

            fechasLigeras.Should().BeEquivalentTo(fechasCompletas);
        }

        [Fact]
        public void CalendarioConflictos_WithLightweightProjection_ShouldMatchFullProjection()
        {
            // Arrange
            var contexto = CrearContexto();
            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 1,
                IdAsignatura = 20,
                IdBloqueHorario = 1,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) }
            };

            // Act
            var ocupaciones = AgCalendarioConflictos.Crear(new[] { gen1, gen2 }, contexto);

            // Assert
            ocupaciones.IsSuccess.Should().BeTrue();
            ocupaciones.Value.Should().HaveCount(2);
            ocupaciones.Value[0].Fecha.Should().Be(new DateTime(2026, 2, 2));
            ocupaciones.Value[0].Inicio.Should().Be(TimeSpan.FromHours(7));
            ocupaciones.Value[0].Fin.Should().Be(TimeSpan.FromHours(11));
            ocupaciones.Value[1].Fecha.Should().Be(new DateTime(2026, 2, 9));
        }

        [Fact]
        public void ProyectarFechas_WhenErrorConditions_ShouldMatchExpandirErrorStrings()
        {
            // Arrange
            var contexto = CrearContexto();
            var genFuera = new AgGen
            {
                Id = 99,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                FechaInicioEstimada = new DateTime(2025, 1, 1),
                FechaFinEstimada = new DateTime(2025, 1, 10),
                FechasGen = new List<DateTime> { new DateTime(2025, 1, 5) }
            };

            var expansor = new AgExpansorSesiones();

            // Act
            var resFechas = AgExpansorSesiones.ProyectarFechas(genFuera, contexto);
            var resSesiones = expansor.Expandir(genFuera, 0, contexto);

            // Assert
            resFechas.IsFailure.Should().BeTrue();
            resSesiones.IsFailure.Should().BeTrue();
            resFechas.Error.Should().Be(resSesiones.Error);
        }
    }
}
