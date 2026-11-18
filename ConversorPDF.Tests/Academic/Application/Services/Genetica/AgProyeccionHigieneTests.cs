using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgProyeccionHigieneTests
    {
        [Fact]
        public void AgBloqueHorario_WhenInitializedOrChanged_ShouldPrecomputeTimeSpansAndHorarioTexto()
        {
            // Arrange & Act
            var bloque = new AgBloqueHorario();

            // Assert defaults
            bloque.TimeSpanInicio.Should().Be(TimeSpan.FromHours(8));
            bloque.TimeSpanFin.Should().Be(TimeSpan.FromHours(12));
            bloque.HorarioTexto.Should().Be("08:00 A 12:00");
            bloque.ObtenerHorarioTexto().Should().Be("08:00 A 12:00");

            // Act - change hours
            bloque.HoraInicio = "14:00";
            bloque.HoraFin = "18:00";

            // Assert updated
            bloque.TimeSpanInicio.Should().Be(TimeSpan.FromHours(14));
            bloque.TimeSpanFin.Should().Be(TimeSpan.FromHours(18));
            bloque.HorarioTexto.Should().Be("14:00 A 18:00");
            bloque.ObtenerHorarioTexto().Should().Be("14:00 A 18:00");
        }

        [Fact]
        public void NormalizarCiudad_WhenCalledMultipleTimes_ShouldReturnConsistentCachedValue()
        {
            // Arrange
            string input = "Bogotá D.C. ";

            // Act
            string norm1 = CalculadoraViaticaSesion.NormalizarCiudad(input);
            string norm2 = CalculadoraViaticaSesion.NormalizarCiudad(input);

            // Assert
            norm1.Should().Be("BOGOTA");
            norm2.Should().Be(norm1);
        }

        [Fact]
        public void AgContexto_WhenConstruirIndicesCalled_ShouldIndexProgramasAndUnidadesPorAula()
        {
            // Arrange
            var contexto = new AgContexto();
            var prog = new Programa { Id = 10, Nombre = "Administración Pública", Modalidad = "Distancia" };
            contexto.Programas.Add(prog);

            var aula = new AgAula
            {
                Id = 1,
                IdPrograma = 10,
                Estado = AgEstadoAula.Disponible,
                Jornada = "SABADO"
            };
            contexto.Aulas.Add(aula);

            // Act
            contexto.ConstruirIndices();

            // Assert
            contexto.ProgramasPorId.Should().ContainKey(10);
            contexto.ProgramasPorId[10].Should().BeSameAs(prog);
            contexto.UnidadHorasPorAulaId.Should().ContainKey(1);
            contexto.UnidadHorasPorAulaId[1].Should().BeGreaterThan(0);
        }

        [Fact]
        public void DisponibleGrupo_WhenSesionesConservadasEmpty_ShouldEvaluateQuicklyWithoutError()
        {
            // Arrange
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
            contexto.RegistrarDocenteHabilitado(101, 0, 10, "");

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

            // Assert
            disp.Should().BeTrue();
        }
    }
}
