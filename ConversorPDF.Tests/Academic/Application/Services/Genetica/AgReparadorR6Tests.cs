using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgReparadorR6Tests
    {
        private readonly IAgEvaluadorAptitud _evaluador;

        public AgReparadorR6Tests()
        {
            _evaluador = new AgEvaluadorAptitud(new AgValidadorReglas());
        }

        private AgContexto CrearContexto()
        {
            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2027, 2, 1),
                    FechaFin = new DateTime(2027, 6, 30),
                    PeriodoNombre = "2027-1",
                    Origen = AgOrigenVentana.Default
                },
                Aulas = new List<AgAula>
                {
                    new AgAula { Id = 12, Nombre = "Aula 12", DiasVentana = "LUN,MAR,MIE,JUE", Jornada = "MANANA", HorasDia = 4 }
                },
                Bloques = new List<AgBloqueHorario>
                {
                    new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 4, DiaSemanaNumero = 4, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" }
                },
                Asignaturas = new List<Asignatura>
                {
                    new Asignatura { Id = 151, Horas = 8, Semestre = "IX", Nombre = "Asig 151" },
                    new Asignatura { Id = 152, Horas = 64, Semestre = "IX", Nombre = "Asig 152" },
                    new Asignatura { Id = 153, Horas = 8, Semestre = "IX", Nombre = "Asig 153" }
                },
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 151, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 152, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 153, OrigenPatron = "LEGADO", Requerida = 1 }
                },
                Festivos = new List<(DateTime Inicio, DateTime Fin)>
                {
                    (new DateTime(2027, 3, 25), new DateTime(2027, 3, 25)),
                    (new DateTime(2027, 3, 26), new DateTime(2027, 3, 26))
                }
            };
            ctx.RegistrarDocenteHabilitado(151, 0, 151, "IX");
            ctx.RegistrarDocenteHabilitado(152, 0, 152, "IX");
            ctx.RegistrarDocenteHabilitado(153, 0, 153, "IX");
            ctx.ConstruirIndices();
            return ctx;
        }

        private AgIndividuo CrearIndividuo(bool fijarCortas = false)
        {
            var ind = new AgIndividuo { Violaciones = new List<AgViolacion>() };

            var fechas152 = new List<DateTime>();
            for (var d = new DateTime(2027, 2, 2); d <= new DateTime(2027, 5, 18); d = d.AddDays(7))
            {
                fechas152.Add(d);
            }

            ind.Genes.Add(new AgGen
            {
                Id = 1, IdAsignatura = 152, IdAula = 12, IdTercero = 152, IdBloqueHorario = 2, DiaSemana = 2,
                DiasPatron = "1,2,3,4", Semestre = "IX", Grupo = "A", TotalHorasAsignatura = 64.0, HorasGen = 64.0,
                FechasGen = fechas152,
                FechaInicioEstimada = fechas152.Min(), FechaFinEstimada = fechas152.Max()
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2, IdAsignatura = 151, IdAula = 12, IdTercero = 151, IdBloqueHorario = 1, DiaSemana = 1,
                DiasPatron = "1,2,3,4", Semestre = "IX", Grupo = "A", TotalHorasAsignatura = 8.0, HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 5, 3), new DateTime(2027, 5, 10) },
                FechaInicioEstimada = new DateTime(2027, 5, 3), FechaFinEstimada = new DateTime(2027, 5, 10),
                EsFijo = fijarCortas ? 1 : 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 3, IdAsignatura = 153, IdAula = 12, IdTercero = 153, IdBloqueHorario = 1, DiaSemana = 1,
                DiasPatron = "1,2,3,4", Semestre = "IX", Grupo = "A", TotalHorasAsignatura = 8.0, HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 4, 12), new DateTime(2027, 4, 19) },
                FechaInicioEstimada = new DateTime(2027, 4, 12), FechaFinEstimada = new DateTime(2027, 4, 19),
                EsFijo = fijarCortas ? 1 : 0
            });

            return ind;
        }

        [Fact]
        public void Reparar_SolapeCarril_MueveCortasDespuesDelBloqueLargo()
        {
            var ctx = CrearContexto();
            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);
            ind.HardViolations.Should().Be(2, string.Join(" | ", ind.Violaciones.Select(v => v.Regla + ":" + v.Mensaje)));

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeTrue();
            res.HardPre.Should().Be(2);
            res.HardPost.Should().Be(0);
            res.Intentos.Should().BeGreaterThan(0);
            res.Individuo.HardViolations.Should().Be(0);

            var g153 = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            var g151 = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 151);
            g153.Should().NotBeNull();
            g151.Should().NotBeNull();
            g153.FechaInicioEstimada.Should().Be(new DateTime(2027, 5, 24));
            g151.FechaInicioEstimada.Should().Be(new DateTime(2027, 6, 7));
        }

        [Fact]
        public void Reparar_CortasFijas_NoMueveYHaceRollback()
        {
            var ctx = CrearContexto();
            var ind = CrearIndividuo(fijarCortas: true);
            new AgValidadorReglas().Evaluar(ind, ctx);
            ind.HardViolations.Should().BeGreaterThan(0);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeFalse();
            res.HardPost.Should().Be(ind.HardViolations);
            res.Individuo.Should().BeSameAs(ind);
        }
    }
}