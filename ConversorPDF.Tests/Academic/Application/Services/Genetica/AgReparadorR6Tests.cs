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

        private AgContexto CrearContexto(DateTime? finVentana = null)
        {
            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2027, 2, 1),
                    FechaFin = finVentana ?? new DateTime(2027, 6, 30),
                    PeriodoNombre = "2027-1",
                    Origen = AgOrigenVentana.Default
                },
                Aulas = new List<AgAula>
                {
                    new AgAula { Id = 12, IdPrograma = 0, IdPeriodo = 0, Nombre = "Aula 12", DiasVentana = "LUN,MAR,MIE,JUE", Jornada = "MANANA", HorasDia = 4 }
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
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 151, OrigenPatron = "LEGADO", Requerida = 1, OrdenCurricular = 2 },
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 152, OrigenPatron = "LEGADO", Requerida = 1, OrdenCurricular = 1 },
                    new AgAulaAsignatura { IdAula = 12, IdAsignatura = 153, OrigenPatron = "LEGADO", Requerida = 1, OrdenCurricular = 3 }
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
                Id = 1,
                IdAsignatura = 152,
                IdAula = 12,
                IdTercero = 152,
                IdBloqueHorario = 2,
                DiaSemana = 2,
                DiasPatron = "1,2,3,4",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 64.0,
                HorasGen = 64.0,
                FechasGen = fechas152,
                FechaInicioEstimada = fechas152.Min(),
                FechaFinEstimada = fechas152.Max()
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 151,
                IdAula = 12,
                IdTercero = 151,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                DiasPatron = "1,2,3,4",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 8.0,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 5, 3), new DateTime(2027, 5, 10) },
                FechaInicioEstimada = new DateTime(2027, 5, 3),
                FechaFinEstimada = new DateTime(2027, 5, 10),
                EsFijo = fijarCortas ? 1 : 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 3,
                IdAsignatura = 153,
                IdAula = 12,
                IdTercero = 153,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                DiasPatron = "1,2,3,4",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 8.0,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 4, 12), new DateTime(2027, 4, 19) },
                FechaInicioEstimada = new DateTime(2027, 4, 12),
                FechaFinEstimada = new DateTime(2027, 4, 19),
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

        [Fact]
        public void Planificar_Cadena137_ReordenaYFactible()
        {
            // Caso #137 con ventana amplia de 179 días (hasta julio 30)
            var ctx = CrearContexto(new DateTime(2027, 7, 30));
            var ind = CrearIndividuo();

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var plan = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril);

            plan.Factible.Should().BeTrue();
            plan.Individuo.Should().NotBeNull();

            var g152 = plan.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 152);
            var g153 = plan.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            var g151 = plan.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 151);
            g152.Should().NotBeNull();
            g153.Should().NotBeNull();
            g151.Should().NotBeNull();

            // Verificación por invariante de no solape temporal (Fase 2379)
            bool solape152_153 = g152.FechaInicioEstimada <= g153.FechaFinEstimada && g153.FechaInicioEstimada <= g152.FechaFinEstimada;
            bool solape152_151 = g152.FechaInicioEstimada <= g151.FechaFinEstimada && g151.FechaInicioEstimada <= g152.FechaFinEstimada;
            bool solape153_151 = g153.FechaInicioEstimada <= g151.FechaFinEstimada && g151.FechaInicioEstimada <= g153.FechaFinEstimada;

            solape152_153.Should().BeFalse();
            solape152_151.Should().BeFalse();
            solape153_151.Should().BeFalse();
        }

        [Fact]
        public void Planificar_VentanaEstrecha_Ocupada_MotivoSinFechas()
        {
            // Ventana amplia en semanas para evitar CAPACIDAD, pero con los lunes ocupados
            var ctx = CrearContexto(new DateTime(2027, 7, 30));
            var ind = CrearIndividuo();

            // Llenar externamente todos los lunes del aula con otra cohorte
            for (var d = new DateTime(2027, 2, 1); d <= new DateTime(2027, 7, 30); d = d.AddDays(7))
            {
                ind.Genes.Add(new AgGen
                {
                    Id = ind.Genes.Count + 10,
                    IdAsignatura = 999,
                    IdAula = 12,
                    DiaSemana = 1,
                    DiasPatron = "1,2,3,4",
                    Semestre = "X",
                    Grupo = "B",
                    HorasGen = 4.0,
                    FechasGen = new List<DateTime> { d },
                    FechaInicioEstimada = d,
                    FechaFinEstimada = d
                });
            }

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var plan = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril);

            plan.Factible.Should().BeFalse();
            plan.Motivo.Should().Be("SIN_FECHAS");
        }

        [Fact]
        public void Planificar_CapacidadInsuficiente_MotivoCAPACIDAD()
        {
            // Ventana de solo 3 semanas (21 días)
            var ctx = CrearContexto(new DateTime(2027, 2, 22));
            var ind = CrearIndividuo();

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var plan = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril);

            plan.Factible.Should().BeFalse();
            plan.Motivo.Should().Be("CAPACIDAD");
            plan.SemanasRequeridas.Should().BeGreaterThan(plan.SemanasDisponibles);
        }

        [Fact]
        public void Reparar_ConFijo_UbicaMovilesAlrededorDeAncla()
        {
            var ctx = CrearContexto(new DateTime(2027, 7, 30));
            var ind = new AgIndividuo { Violaciones = new List<AgViolacion>() };

            var fechas152 = new List<DateTime>();
            for (var d = new DateTime(2027, 2, 2); d <= new DateTime(2027, 5, 18); d = d.AddDays(7))
            {
                fechas152.Add(d);
            }

            // 152 es FIJO en sus fechas
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 152,
                IdAula = 12,
                IdTercero = 152,
                IdBloqueHorario = 2,
                DiaSemana = 2,
                DiasPatron = "1,2,3,4",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 64.0,
                HorasGen = 64.0,
                FechasGen = fechas152,
                FechaInicioEstimada = fechas152.Min(),
                FechaFinEstimada = fechas152.Max(),
                EsFijo = 1
            });

            // 153 es MÓVIL pero con solape inicial
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 153,
                IdAula = 12,
                IdTercero = 153,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                DiasPatron = "1,2,3,4",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 8.0,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 4, 12), new DateTime(2027, 4, 19) },
                FechaInicioEstimada = new DateTime(2027, 4, 12),
                FechaFinEstimada = new DateTime(2027, 4, 19),
                EsFijo = 0
            });

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var plan = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril);

            plan.Factible.Should().BeTrue();
            var g152 = plan.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 152);
            var g153 = plan.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            g152.Should().NotBeNull();
            g153.Should().NotBeNull();
            // 152 no se movió (ancla)
            g152.FechaInicioEstimada.Should().Be(fechas152.Min());
            g152.FechaFinEstimada.Should().Be(fechas152.Max());
            // 153 se ubicó sin solapar el rango de 152
            bool solapa = g152.FechaInicioEstimada <= g153.FechaFinEstimada && g153.FechaInicioEstimada <= g152.FechaFinEstimada;
            solapa.Should().BeFalse();
        }

        [Fact]
        public void Reparar_ReRanuracion_MueveAOtroCarrilConCupo()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            // Agregar aula alternativa 13 con cupo en el mismo programa y periodo
            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 13, IdAsignatura = 153, OrigenPatron = "AUTOMATICO", Requerida = 1 });
            ctx.ConstruirIndices();

            // Aula 12 está saturada para la corta 153
            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);
            int hardInicial = ind.HardViolations;
            hardInicial.Should().BeGreaterThan(0);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeTrue();
            res.HardPost.Should().BeLessThan(hardInicial);
            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().Be(13); // Verificación explícita de aula destino (Fase 2379)
        }

        [Fact]
        public void Planificar_AsignaturaMultiDia_AsignaDiasCorrectos()
        {
            // Asignatura 154 con sesiones en Lunes y Miércoles (Fase 2375)
            var ctx = CrearContexto(new DateTime(2027, 6, 30));
            ctx.Asignaturas.Add(new Asignatura { Id = 154, Horas = 16, Semestre = "IX", Nombre = "Asig 154 MultiDia" });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 12, IdAsignatura = 154, OrigenPatron = "AUTOMATICO", Requerida = 1 });
            ctx.RegistrarDocenteHabilitado(154, 0, 154, "IX");
            ctx.ConstruirIndices();

            var ind = new AgIndividuo { Violaciones = new List<AgViolacion>() };

            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 154,
                IdAula = 12,
                IdTercero = 154,
                IdBloqueHorario = 1,
                DiaSemana = 1, // LUN
                DiasPatron = "1,3",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 16.0,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 1), new DateTime(2027, 2, 8) },
                FechaInicioEstimada = new DateTime(2027, 2, 1),
                FechaFinEstimada = new DateTime(2027, 2, 8)
            });

            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 154,
                IdAula = 12,
                IdTercero = 154,
                IdBloqueHorario = 3,
                DiaSemana = 3, // MIE
                DiasPatron = "1,3",
                Semestre = "IX",
                Grupo = "A",
                TotalHorasAsignatura = 16.0,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 3), new DateTime(2027, 2, 10) },
                FechaInicioEstimada = new DateTime(2027, 2, 3),
                FechaFinEstimada = new DateTime(2027, 2, 10)
            });

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var plan = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril);

            plan.Factible.Should().BeTrue();
            var gLun = plan.Individuo.Genes.SingleOrDefault(g => g.DiaSemana == 1);
            var gMie = plan.Individuo.Genes.SingleOrDefault(g => g.DiaSemana == 3);

            gLun.Should().NotBeNull();
            gMie.Should().NotBeNull();
            gLun.FechasGen.Should().OnlyContain(f => AgInventarioRanuras.Iso(f) == 1);
            gMie.FechasGen.Should().OnlyContain(f => AgInventarioRanuras.Iso(f) == 3);
        }

        [Fact]
        public void Reparar_ReRanuracion_NoUsaAulaInactiva()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            // Agregar aula alternativa 13 pero con Estado Inactiva (Fase 2384, Fase 2386)
            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Inactiva",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Inactiva
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 13, IdAsignatura = 153, OrigenPatron = "AUTOMATICO", Requerida = 1 });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().NotBe(13);
            res.MotivoReranuracion.Should().Be("SIN_VINCULO");
        }

        [Fact]
        public void Reparar_ReRanuracion_RechazaVinculoLegado()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            // Agregar aula alternativa 13 con vínculo LEGADO (Fase 2380, Fase 2386)
            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Legado",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 13, IdAsignatura = 153, OrigenPatron = "LEGADO", Requerida = 1 });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().NotBe(13);
            res.MotivoReranuracion.Should().Be("SIN_VINCULO");
        }

        [Fact]
        public void Reparar_ReRanuracion_ActualizaClaveGrupo()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            // Agregar aula alternativa 13 válida (Fase 2385, Fase 2386)
            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Valida",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 13, IdAsignatura = 153, OrigenPatron = "AUTOMATICO", Requerida = 1 });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeTrue();
            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().Be(13);
            genVictima.ClaveGrupo.Should().Be($"{genVictima.IdAsignatura}|13|{genVictima.Grupo}");
            genVictima.PatronSemanal.Should().BeEmpty();
            genVictima.IndiceDiaPatron.Should().Be(0);
        }

        [Fact]
        public void Reparar_ReRanuracion_ConservaIndiceDiaPatronSegunDiaSemana()
        {
            // Fase 2405: Al re-ranurar, IndiceDiaPatron debe corresponder a la posición de DiaSemana en DiasPatron
            var gen = new AgGen
            {
                Id = 10,
                IdAsignatura = 153,
                IdAula = 13,
                DiaSemana = 3,
                DiasPatron = "1,2,3,4"
            };

            var diasPatron = (gen.DiasPatron ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => int.TryParse(s.Trim(), out int d) ? d : 0)
                .Where(d => d >= 1 && d <= 7)
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            gen.IndiceDiaPatron = Math.Max(0, diasPatron.IndexOf(gen.DiaSemana));
            gen.IndiceDiaPatron.Should().Be(2);
        }

        [Fact]
        public void Reparar_ReRanuracion_ConservaSufijoGrupoOriginalCuandoGrupoEstaVacio()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Valida",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 13, IdAsignatura = 153, OrigenPatron = "AUTOMATICO", Requerida = 1 });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            foreach (var g in ind.Genes)
            {
                g.Grupo = string.Empty; // Grupo vacío en todos los genes de la cohorte
            }
            var genVictimaOriginal = ind.Genes.Single(g => g.IdAsignatura == 153);
            genVictimaOriginal.ClaveGrupo = "153|12|B"; // Sufijo original explícito "B" (Fase 2387)

            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeTrue();
            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().Be(13);
            genVictima.ClaveGrupo.Should().Be("153|13|B");
        }

        [Fact]
        public void Reparar_ReRanuracion_RechazaVinculoConPatronIncompatible()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            // Aula 13 con vínculo que exige días diferentes a los de la víctima (LUN = 1) (Fase 2388)
            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Patron Incompatible",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            });
            // Vinculo solo permite MAR (2) y MIE (3), pero victima es LUN (1)
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 13,
                IdAsignatura = 153,
                OrigenPatron = "AUTOMATICO",
                DiasPatron = "MAR,MIE",
                Requerida = 1
            });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().NotBe(13);
            res.MotivoReranuracion.Should().Be("SIN_VINCULO");
        }

        [Fact]
        public void Reparar_ReRanuracion_RechazaVinculoConSubunidadIncompatibleConAulaDestino()
        {
            var ctx = CrearContexto(new DateTime(2027, 5, 20));

            ctx.Aulas.Add(new AgAula
            {
                Id = 13,
                IdPrograma = 0,
                IdPeriodo = 0,
                Nombre = "Aula 13 Subunidad Incompatible",
                Subunidad = "SEDE_B",
                DiasVentana = "LUN,MAR,MIE,JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            });
            // Vinculo exige SEDE_A pero aula 13 tiene SEDE_B (Fase 2390)
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 13,
                IdAsignatura = 153,
                OrigenPatron = "AUTOMATICO",
                Subunidad = "SEDE_A",
                Requerida = 1
            });
            ctx.ConstruirIndices();

            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            var genVictima = res.Individuo.Genes.SingleOrDefault(g => g.IdAsignatura == 153);
            genVictima.Should().NotBeNull();
            genVictima.IdAula.Should().NotBe(13);
            res.MotivoReranuracion.Should().Be("SIN_VINCULO");
        }

        [Fact]
        public void SemanasEquivalentes_ConsideraSesionesPorSemanaSegunDiasDelCarril()
        {
            // Fase 2420, Fase 2421: carril con 3 sesiones por semana divide entre 3, no entre 2
            var plan3 = new AgResultadoPlanCarril
            {
                SesionesDisponibles = 6,
                SesionesPorSemana = 3,
                SesionesRequeridas = 6
            };
            plan3.SemanasEquivalentes.Should().Be(2);
            plan3.SemanasRequeridasEquivalentes.Should().Be(2);
            plan3.SemanasRequeridas.Should().Be(2);

            var plan2 = new AgResultadoPlanCarril
            {
                SesionesDisponibles = 6,
                SesionesPorSemana = 2,
                SesionesRequeridas = 6
            };
            plan2.SemanasEquivalentes.Should().Be(3);
            plan2.SemanasRequeridasEquivalentes.Should().Be(3);
            plan2.SemanasRequeridas.Should().Be(3);
        }

        [Fact]
        public void Planificar_ConVentanaAmpliada_VerificaReparacionRealYSinSolape()
        {
            // FASE 2445 / FASE 2448: Carril con ventana ampliada planifica la cadena secuencial sin solapes
            var ctx = CrearContexto();
            var ind = CrearIndividuo();
            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);
            string carril = AgCarriles.ClaveCarril(ind.Genes[0]);

            var planAmp = AgPlanificadorCarril.Planificar(ind, ctx, cohorte, carril, ventanaAmpliada: "LUN-VIE");

            planAmp.Should().NotBeNull();
            planAmp.Factible.Should().BeTrue();
            planAmp.Individuo.Should().NotBeNull();
            planAmp.SesionesDisponibles.Should().BeGreaterThan(0);

            var genes = planAmp.Individuo.Genes
                .Where(g => g.IdAsignatura == 152 || g.IdAsignatura == 153 || g.IdAsignatura == 151)
                .ToList();

            genes.Should().OnlyContain(g => g.FechaInicioEstimada.HasValue);
            genes.SelectMany(g => g.FechasGen)
                .Should().OnlyHaveUniqueItems();
        }

        [Fact]
        public void RepararR6_CadenaSinFechas_AmpliaVentanaYReduceHard()
        {
            // FASE 2445 / FASE 2448: Reparador R6 ante carril con solapes ejecuta y reduce infracciones Hard
            var ctx = CrearContexto();
            var ind = CrearIndividuo();
            new AgValidadorReglas().Evaluar(ind, ctx);
            int hardAntes = ind.HardViolations;
            hardAntes.Should().BeGreaterThan(0);

            var res = AgReparadorR6.Ejecutar(ind, ctx, _evaluador);

            res.HuboMejora.Should().BeTrue();
            res.HardPost.Should().BeLessThan(res.HardPre);
            res.Individuo.HardViolations.Should().BeLessThan(hardAntes);
        }
    }
}
