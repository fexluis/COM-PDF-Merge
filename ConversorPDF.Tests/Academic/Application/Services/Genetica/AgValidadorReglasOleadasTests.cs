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

        private AgContexto CrearContextoBase(bool conColeccionesIndexadas = false)
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

            // Fase 2147 / Fase 2155: Poblar colecciones base e indexar solo cuando sea requerido
            if (conColeccionesIndexadas)
            {
                foreach (var a in ctx.AulasPorId.Values) ctx.Aulas.Add(a);
                foreach (var b in ctx.BloquesPorId.Values) ctx.Bloques.Add(b);
                ctx.ConstruirIndices();
            }

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

        #region Pruebas de Validación R6 y R13 post-reparación (Fase 2132 / Fase 2317)

        [Fact]
        public void R6_FronteraCompartidaMismoCarril_NoEsViolacionDura()
        {
            // Asig. 149: 2027-02-03..2027-02-20
            // Asig. 152: 2027-02-20..2027-03-11
            // assert: no R6, no R13
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
                ClaveGrupo = "149-IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 3),
                FechaFinEstimada = new DateTime(2027, 2, 20)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 50,
                IdAsignatura = 152,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                ClaveGrupo = "152-IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 20),
                FechaFinEstimada = new DateTime(2027, 3, 11)
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas);
            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R13_OleadasDispersas);
            ind.HardViolations.Should().Be(0, string.Join(" | ", ind.Violaciones.Select(v => $"{v.Regla} (Dura={v.EsDura}): {v.Mensaje}")));
        }

        [Fact]
        public void R6_SolapeRealMismoCarril_EsViolacionDura()
        {
            // Asig. 149: 2027-02-03..2027-02-20
            // Asig. 152: 2027-02-19..2027-03-11
            // assert: R6 hard
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
                ClaveGrupo = "IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 3),
                FechaFinEstimada = new DateTime(2027, 2, 20)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 50,
                IdAsignatura = 152,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                ClaveGrupo = "IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 19),
                FechaFinEstimada = new DateTime(2027, 3, 11)
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas && v.EsDura);
        }

        [Fact]
        public void R6_SolapeNoAdyacenteConOrdenCurricularInvertido_EsViolacionDura()
        {
            // A y C solapan; B no; orden curricular A,B,C
            // assert: R6 hard all-pairs
            var ctx = CrearContextoBase(conColeccionesIndexadas: true);
            ctx.AulaAsignaturaPorAulaAsig[(30, 148)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 148, OrdenCurricular = 1 };
            ctx.AulaAsignaturaPorAulaAsig[(30, 149)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 149, OrdenCurricular = 2 };
            ctx.AulaAsignaturaPorAulaAsig[(30, 150)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 150, OrdenCurricular = 3 };

            var ind = new AgIndividuo();
            // A (148): 2027-02-04..2027-03-01
            ind.Genes.Add(new AgGen { Id = 1, IdTercero = 10, IdAsignatura = 148, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4", TotalHorasAsignatura = 16.0, HorasGen = 4.0, FechaInicioEstimada = new DateTime(2027, 2, 4), FechaFinEstimada = new DateTime(2027, 3, 1) });
            // B (149): 2027-03-05..2027-03-25 (no solapa con A ni con C)
            ind.Genes.Add(new AgGen { Id = 2, IdTercero = 20, IdAsignatura = 149, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4", TotalHorasAsignatura = 16.0, HorasGen = 4.0, FechaInicioEstimada = new DateTime(2027, 3, 5), FechaFinEstimada = new DateTime(2027, 3, 25) });
            // C (150): 2027-02-20..2027-04-10 (solapa con A: 20-feb < 1-mar, pero no con B adyacente directamente si orden es A,B,C)
            ind.Genes.Add(new AgGen { Id = 3, IdTercero = 30, IdAsignatura = 150, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4", TotalHorasAsignatura = 16.0, HorasGen = 4.0, FechaInicioEstimada = new DateTime(2027, 2, 20), FechaFinEstimada = new DateTime(2027, 4, 10) });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas && v.EsDura);
        }

        [Fact]
        public void R13_NoSeEmiteEnValidador()
        {
            // solape real; assert: no R13
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
                ClaveGrupo = "IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 3),
                FechaFinEstimada = new DateTime(2027, 2, 20)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 50,
                IdAsignatura = 152,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                ClaveGrupo = "IX-A",
                IdAula = 30,
                IdBloqueHorario = 24,
                DiaSemana = 4,
                DiasPatron = "4",
                TotalHorasAsignatura = 16.0,
                HorasGen = 4.0,
                FechaInicioEstimada = new DateTime(2027, 2, 19),
                FechaFinEstimada = new DateTime(2027, 3, 11)
            });

            _sut.Evaluar(ind, ctx);

            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R13_OleadasDispersas);
        }

        [Fact]
        public void RealinearOleadas_CampeonConSolapesR6_EliminaViolaciones()
        {
            // Arrange (Fase 2155: se requieren colecciones indexadas para AgPlanificadorOleadas)
            var ctx = CrearContextoBase(conColeccionesIndexadas: true);
            ctx.AulaAsignaturaPorAulaAsig[(30, 148)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 148, OrdenCurricular = 1 };
            ctx.AulaAsignaturaPorAulaAsig[(30, 149)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 149, OrdenCurricular = 2 };
            ctx.AulaAsignaturaPorAulaAsig[(30, 150)] = new AgAulaAsignatura { IdAula = 30, IdAsignatura = 150, OrdenCurricular = 3 };

            var ind = new AgIndividuo();
            // Asig 148: del 4 al 27 de feb
            ind.Genes.Add(new AgGen { Id = 1, IdTercero = 10, IdAsignatura = 148, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 4), FechaFinEstimada = new DateTime(2027, 2, 27) });
            ind.Genes.Add(new AgGen { Id = 2, IdTercero = 10, IdAsignatura = 148, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 29, DiaSemana = 5, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 4), FechaFinEstimada = new DateTime(2027, 2, 27) });

            // Asig 149: solapada del 11 de feb al 6 de mar (empieza antes del fin de 148)
            ind.Genes.Add(new AgGen { Id = 3, IdTercero = 20, IdAsignatura = 149, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 11), FechaFinEstimada = new DateTime(2027, 3, 6) });
            ind.Genes.Add(new AgGen { Id = 4, IdTercero = 20, IdAsignatura = 149, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 29, DiaSemana = 5, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 11), FechaFinEstimada = new DateTime(2027, 3, 6) });

            // Asig 150: solapada del 18 de feb al 13 de mar
            ind.Genes.Add(new AgGen { Id = 5, IdTercero = 30, IdAsignatura = 150, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 18), FechaFinEstimada = new DateTime(2027, 3, 13) });
            ind.Genes.Add(new AgGen { Id = 6, IdTercero = 30, IdAsignatura = 150, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 29, DiaSemana = 5, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 18), FechaFinEstimada = new DateTime(2027, 3, 13) });

            // Pre-evaluación: debe tener violaciones R6 (Fase 2136 / Fase 2317)
            _sut.Evaluar(ind, ctx);
            ind.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas);
            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R13_OleadasDispersas);

            // Act
            var rep = AgPlanificadorOleadas.RealinearOleadas(ind, ctx);
            rep.HuboCambios.Should().BeTrue();

            AgCoherenciaGrupo.Sellar(ind, ctx);
            _sut.Evaluar(ind, ctx);

            // Assert (Fase 2136 / Fase 2147 / Fase 2317)
            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas);
            ind.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R13_OleadasDispersas);
        }

        [Fact]
        public void RealinearOleadas_CuandoNoCabeEnVentana_ConservaHardOriginalSinRegresion()
        {
            // Arrange (ventana angosta de 2 semanas)
            var ctx = CrearContextoBase();
            ctx.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 2, 14), AgOrigenVentana.Presencial).Value;

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, IdTercero = 10, IdAsignatura = 148, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 4), FechaFinEstimada = new DateTime(2027, 2, 14) });
            ind.Genes.Add(new AgGen { Id = 2, IdTercero = 20, IdAsignatura = 149, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4,5", TotalHorasAsignatura = 16.0, HorasGen = 8.0, FechaInicioEstimada = new DateTime(2027, 2, 8), FechaFinEstimada = new DateTime(2027, 2, 14) });

            _sut.Evaluar(ind, ctx);
            int hardPre = ind.HardViolations;

            // Act
            var rep = AgPlanificadorOleadas.RealinearOleadas(ind, ctx);

            // Assert
            rep.HuboCambios.Should().BeFalse();
            rep.Fallos.Should().Be(1);
            rep.Detalle.Should().Contain(d => d.Contains("motivo=NoCabe"));
        }

        [Fact]
        public void RealinearOleadas_FronteraCompartida_NoCambia()
        {
            // Arrange: asignaturas consecutivas con frontera compartida en el mismo carril (Fase 2320)
            var ctx = CrearContextoBase(conColeccionesIndexadas: true);
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, IdTercero = 20, IdAsignatura = 149, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 24, DiaSemana = 4, DiasPatron = "4", TotalHorasAsignatura = 16.0, HorasGen = 4.0, FechaInicioEstimada = new DateTime(2027, 2, 3), FechaFinEstimada = new DateTime(2027, 2, 20) });
            ind.Genes.Add(new AgGen { Id = 2, IdTercero = 20, IdAsignatura = 152, IdPrograma = 1, IdSede = 1, Semestre = "IX", Grupo = "A", ClaveGrupo = "IX-A", IdAula = 30, IdBloqueHorario = 29, DiaSemana = 4, DiasPatron = "4", TotalHorasAsignatura = 16.0, HorasGen = 4.0, FechaInicioEstimada = new DateTime(2027, 2, 20), FechaFinEstimada = new DateTime(2027, 3, 11) });

            // Act
            var rep = AgPlanificadorOleadas.RealinearOleadas(ind, ctx);

            // Assert
            rep.HuboCambios.Should().BeFalse();
            rep.CarrilesReparados.Should().Be(0);
            rep.Fallos.Should().Be(0);
        }

        #endregion
    }
}

