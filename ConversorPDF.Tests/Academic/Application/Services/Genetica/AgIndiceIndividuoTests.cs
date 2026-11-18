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
    public class AgIndiceIndividuoTests
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

            contexto.RegistrarDocenteHabilitado(101, 0, 10, "");
            contexto.RegistrarDocenteHabilitado(102, 0, 10, "");

            return contexto;
        }

        [Fact]
        public void ProyeccionIndividuo_WhenCreated_ShouldIndexGenesAndGroups()
        {
            // Arrange
            var contexto = CrearContexto();
            var ind = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 101,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G1",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            ind.Genes.Add(gen);

            // Act
            var proyeccion = new AgProyeccionIndividuo(ind, contexto);

            // Assert
            var grupo = proyeccion.ObtenerGrupo(gen);
            grupo.Should().ContainSingle().Which.Should().BeSameAs(gen);
        }

        [Fact]
        public void ProyeccionIndividuo_DisponibleDocente_ShouldMatchDisponibleGrupoEquivalence()
        {
            // Arrange
            var contexto = CrearContexto();
            var ind = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 101,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G1",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            ind.Genes.Add(gen);

            var proyeccion = new AgProyeccionIndividuo(ind, contexto);

            // Act
            bool dispGrupo101 = AgAsignadorTercero.DisponibleGrupo(101, gen, ind, contexto);
            bool dispProy101 = proyeccion.DisponibleDocente(101, gen);

            bool dispGrupo102 = AgAsignadorTercero.DisponibleGrupo(102, gen, ind, contexto);
            bool dispProy102 = proyeccion.DisponibleDocente(102, gen);

            // Assert
            dispProy101.Should().Be(dispGrupo101);
            dispProy102.Should().Be(dispGrupo102);
            dispProy101.Should().BeTrue();
            dispProy102.Should().BeTrue();
        }

        [Fact]
        public void ProyeccionIndividuo_WhenTeacherHasConflict_ShouldReportNotAvailable()
        {
            // Arrange
            var contexto = CrearContexto();
            var ind = new AgIndividuo();
            // Dos genes en distintas aulas el mismo día y hora
            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 101,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G1",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 2,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 102,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G2",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            ind.Genes.Add(gen1);
            ind.Genes.Add(gen2);

            var proyeccion = new AgProyeccionIndividuo(ind, contexto);

            // Act - Intentar asignar docente 101 al grupo 2 (que se solapa con grupo 1)
            bool disp101EnG2 = proyeccion.DisponibleDocente(101, gen2);

            // Assert
            disp101EnG2.Should().BeFalse();
        }

        [Fact]
        public void ProyeccionIndividuo_ActualizarDocenteGrupo_ShouldUpdateOccupancyImmediately()
        {
            // Arrange
            var contexto = CrearContexto();
            var ind = new AgIndividuo();
            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 102,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G1",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 2,
                IdAsignatura = 10,
                IdBloqueHorario = 1,
                IdTercero = 0,
                HorasGen = 4,
                TotalHorasAsignatura = 4,
                ClaveGrupo = "G2",
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) }
            };
            ind.Genes.Add(gen1);
            ind.Genes.Add(gen2);

            var proyeccion = new AgProyeccionIndividuo(ind, contexto);

            // Act
            proyeccion.DisponibleDocente(101, gen1).Should().BeTrue();
            proyeccion.ActualizarDocenteGrupo(gen1, 101);

            // Ahora 101 está en G1, no debe estar disponible para G2 (mismo horario)
            bool disp101EnG2 = proyeccion.DisponibleDocente(101, gen2);

            // Assert
            disp101EnG2.Should().BeFalse();
        }
    }
}
