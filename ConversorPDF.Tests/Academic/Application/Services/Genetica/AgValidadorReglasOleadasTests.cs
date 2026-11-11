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
    /// <summary>
    /// Pruebas integradas de regresión para el escenario real de oleadas consecutivas (Fase 1573 / Caso #48).
    /// </summary>
    public class AgValidadorReglasOleadasTests
    {
        private readonly AgValidadorReglas _sut = new AgValidadorReglas();

        private AgContexto CrearContextoBase()
        {
            var ctx = new AgContexto();
            ctx.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;

            // Aulas del caso real #48
            ctx.AulasPorId[12] = new AgAula { Id = 12, Codigo = "AULA-12", Semestre = "X", Grupo = "A", IdPrograma = 1, Estado = AgEstadoAula.Disponible, DiasVentana = "VIE-SAB" };
            ctx.AulasPorId[30] = new AgAula { Id = 30, Codigo = "AULA-30", Semestre = "IX", Grupo = "A", IdPrograma = 1, Estado = AgEstadoAula.Disponible, DiasVentana = "JUE-SAB" };
            ctx.AulasPorId[31] = new AgAula { Id = 31, Codigo = "AULA-31", Semestre = "IX", Grupo = "A", IdPrograma = 1, Estado = AgEstadoAula.Disponible, DiasVentana = "JUE-SAB" };

            // Bloques de ejemplo
            ctx.BloquesPorId[24] = new AgBloqueHorario { Id = 24, DiaSemanaNumero = 4, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA" }; // Jueves
            ctx.BloquesPorId[29] = new AgBloqueHorario { Id = 29, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA" }; // Viernes
            ctx.BloquesPorId[5] = new AgBloqueHorario { Id = 5, DiaSemanaNumero = 6, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA" };  // Sábado mañana
            ctx.BloquesPorId[6] = new AgBloqueHorario { Id = 6, DiaSemanaNumero = 6, HoraInicio = "13:00", HoraFin = "17:00", Franja = "TARDE" };  // Sábado tarde

            // Docentes
            ctx.RegistrarDocenteHabilitado(10, 1, 148, "IX");
            ctx.RegistrarDocenteHabilitado(20, 1, 149, "IX");
            ctx.RegistrarDocenteHabilitado(30, 1, 150, "IX");
            ctx.RegistrarDocenteHabilitado(40, 1, 151, "IX");
            ctx.RegistrarDocenteHabilitado(50, 1, 152, "IX");
            ctx.RegistrarDocenteHabilitado(60, 1, 153, "IX");
            ctx.RegistrarDocenteHabilitado(70, 1, 154, "X");
            ctx.RegistrarDocenteHabilitado(80, 1, 156, "X");

            return ctx;
        }

        [Fact]
        public void Oleadas_XA_Viernes_MismaAula_SinViolaciones()
        {
            // Aula 12 (X-A), bloque 29 (viernes 08:00–12:00): gen 154 (02/02→05/04) y gen 156 (06/04→19/04)
            var ctx = CrearContextoBase();
            var ind = new AgIndividuo();

            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 70,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 5)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 80,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 4, 6),
                FechaFinEstimada = new DateTime(2027, 4, 19)
            });

            _sut.Evaluar(ind, ctx);

            ind.HardViolations.Should().Be(0, "violaciones detectadas: " + string.Join(" | ", ind.Violaciones.Select(v => v.Regla + " -> " + v.Mensaje)));
        }

        [Fact]
        public void Oleadas_IXA_Jueves_TresAsignaturas_SinViolaciones()
        {
            // Aula 30 (IX-A), bloque 24: oleadas 148 (→2027-03-15), 150 (2027-03-16→2027-05-03), 152 (2027-05-04→2027-06-20)
            var ctx = CrearContextoBase();
            var ind = new AgIndividuo();

            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 148,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 3, 15)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 30,
                IdAsignatura = 150,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                FechaInicioEstimada = new DateTime(2027, 3, 16),
                FechaFinEstimada = new DateTime(2027, 5, 3)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 3,
                IdTercero = 50,
                IdAsignatura = 152,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                FechaInicioEstimada = new DateTime(2027, 5, 4),
                FechaFinEstimada = new DateTime(2027, 6, 20)
            });

            _sut.Evaluar(ind, ctx);

            ind.HardViolations.Should().Be(0, "violaciones detectadas: " + string.Join(" | ", ind.Violaciones.Select(v => v.Regla + " -> " + v.Mensaje)));
        }

        [Fact]
        public void Oleadas_IXA_Sabado_DosBloques_SinViolaciones()
        {
            // Sábado bloques 5 (08:00–12:00) y 6 (13:00–17:00): 149 (→2027-04-05), 151 (04-06→04-19), 153 (04-20→06-07)
            var ctx = CrearContextoBase();
            var ind = new AgIndividuo();

            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 20,
                IdAsignatura = 149,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 5,
                DiaSemana = 6,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 5)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 40,
                IdAsignatura = 151,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 5,
                DiaSemana = 6,
                FechaInicioEstimada = new DateTime(2027, 4, 6),
                FechaFinEstimada = new DateTime(2027, 4, 19)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 3,
                IdTercero = 60,
                IdAsignatura = 153,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 30,
                IdBloqueHorario = 6,
                DiaSemana = 6,
                FechaInicioEstimada = new DateTime(2027, 4, 20),
                FechaFinEstimada = new DateTime(2027, 6, 7)
            });

            _sut.Evaluar(ind, ctx);

            ind.HardViolations.Should().Be(0, "violaciones detectadas: " + string.Join(" | ", ind.Violaciones.Select(v => v.Regla + " -> " + v.Mensaje)));
        }

        [Fact]
        public void VentanasCruzadas_MismaAula_EmiteR2YR14()
        {
            // Bloque 5 y 6 solapados en horas o mismo bloque cruzado
            var ctx = CrearContextoBase();
            // Creamos bloque solapado 99 con bloque 29
            ctx.BloquesPorId[99] = new AgBloqueHorario { Id = 99, DiaSemanaNumero = 5, HoraInicio = "10:00", HoraFin = "14:00", Franja = "MANANA" };

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 70,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 15)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 80,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 99, // Distinto bloque, pero horas cruzadas e intervalo de fecha cruzado
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 4, 1),
                FechaFinEstimada = new DateTime(2027, 5, 1)
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R14_BloqueSolapadoPorIntervalo);
        }

        [Fact]
        public void SinFechas_MismaAula_EmiteConservador()
        {
            var ctx = CrearContextoBase();
            var ind = new AgIndividuo();

            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 70,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = null,
                FechaFinEstimada = null
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 80,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = null,
                FechaFinEstimada = null
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
        }

        [Fact]
        public void AsignaturasFueraDeAlcance_NoEmiteR5Requerida()
        {
            var ctx = CrearContextoBase();
            ctx.Asignaturas = new List<Asignatura>
            {
                new Asignatura { Id = 154, Nombre = "SEMINARIO ELECTIVO II" },
                new Asignatura { Id = 999, Nombre = "MATEMATICA I (Semestre I - Aula 13)" }
            };
            ctx.AulaAsignaturas = new List<AgAulaAsignatura>
            {
                new AgAulaAsignatura { IdAula = 12, IdAsignatura = 154 },
                new AgAulaAsignatura { IdAula = 13, IdAsignatura = 999 } // Aula 13 no seleccionada / no cubierta
            };

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 70,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 5)
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R5_HabilitacionDocentePool
                && v.Mensaje.Contains("MATEMATICA I"));
        }
    }
}

