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
    public class AgReparadorR5Tests
    {
        // Fase 2341: Usar evaluador real estandarizado para no enmascarar fallos de reevaluación
        private readonly IAgEvaluadorAptitud _evaluadorReal;

        public AgReparadorR5Tests()
        {
            _evaluadorReal = new AgEvaluadorAptitud(new AgValidadorReglas());
        }

        private AgContexto CrearContextoBase(DateTime fIni, DateTime fFin, int maxCarriles = 2, bool incluirFixture132 = false)
        {
            var asignaturas = incluirFixture132
                ? new List<Asignatura>
                {
                    new Asignatura { Id = 149, Horas = 12, Semestre = "I", Nombre = "Asig 149" },
                    new Asignatura { Id = 150, Horas = 12, Semestre = "I", Nombre = "Asig 150" },
                    new Asignatura { Id = 152, Horas = 12, Semestre = "I", Nombre = "Asig 152" }
                }
                : new List<Asignatura>
                {
                    new Asignatura { Id = 100, Horas = 12, Semestre = "I", Nombre = "Asig 100" },
                    new Asignatura { Id = 110, Horas = 12, Semestre = "I", Nombre = "Asig 110" }
                };

            var aulaAsignaturas = incluirFixture132
                ? new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 149, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 150, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 152, OrigenPatron = "AUTOMATICO", Requerida = 1 }
                }
                : new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 100, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 110, OrigenPatron = "LEGADO", Requerida = 1 }
                };

            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = fIni,
                    FechaFin = fFin,
                    PeriodoNombre = "2027-1",
                    Origen = AgOrigenVentana.Default
                },
                Aulas = new List<AgAula>
                {
                    new AgAula
                    {
                        Id = 1,
                        Nombre = "Aula 101",
                        DiasVentana = maxCarriles == 2 ? "MAR,MIE,JUE,VIE" : "LUN,MAR",
                        Jornada = "MANANA",
                        HorasDia = 4
                    }
                },
                Bloques = new List<AgBloqueHorario>
                {
                    new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 4, DiaSemanaNumero = 4, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 5, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" }
                },
                Asignaturas = asignaturas,
                AulaAsignaturas = aulaAsignaturas,
                // Fase 2346: Tipado correcto List<(DateTime Inicio, DateTime Fin)>
                Festivos = new List<(DateTime Inicio, DateTime Fin)>
                {
                    (new DateTime(2027, 3, 25), new DateTime(2027, 3, 25)), // Jueves Santo
                    (new DateTime(2027, 3, 26), new DateTime(2027, 3, 26))  // Viernes Santo
                }
            };
            if (incluirFixture132)
            {
                ctx.RegistrarDocenteHabilitado(149, 0, 149, "I");
                ctx.RegistrarDocenteHabilitado(150, 0, 150, "I");
                ctx.RegistrarDocenteHabilitado(152, 0, 152, "I");
            }
            else
            {
                ctx.RegistrarDocenteHabilitado(10, 0, 100, "I");
                ctx.RegistrarDocenteHabilitado(11, 0, 110, "I");
            }
            ctx.ConstruirIndices();
            return ctx;
        }

        [Fact]
        public void Reparar_Fixture132Sintetico_DesplazaAsignatura150YReduceHardAUnoACero()
        {
            // Arrange (Fases 2341, 2347, 2350, 2353): Fixture corrida 132 fiel con orden temporal real:
            // 149 (LEGADO): Miércoles desde 2027-02-03 hasta 2027-04-01
            // 152 (AUTOMATICO): Martes desde 2027-02-02 hasta 2027-04-01
            // 150 (LEGADO): Viernes desde 2027-03-19 hasta 2027-04-09 (la más tardía -> elegida como víctima)
            // Solape de 3 concurrentes > M=2 -> R5 sobrecarga HARD=1.
            var fIni = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 6, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 2, incluirFixture132: true);

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 149,
                IdAula = 1,
                IdTercero = 149,
                IdBloqueHorario = 3,
                DiaSemana = 3,
                DiasPatron = "MIE",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 3), new DateTime(2027, 3, 3), new DateTime(2027, 3, 31) },
                FechaInicioEstimada = new DateTime(2027, 2, 3),
                FechaFinEstimada = new DateTime(2027, 4, 1),
                EsVacante = 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 152,
                IdAula = 1,
                IdTercero = 152,
                IdBloqueHorario = 2,
                DiaSemana = 2,
                DiasPatron = "MAR",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 2), new DateTime(2027, 3, 2), new DateTime(2027, 3, 30) },
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 1),
                EsVacante = 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 3,
                IdAsignatura = 150,
                IdAula = 1,
                IdTercero = 150,
                IdBloqueHorario = 5,
                DiaSemana = 5,
                DiasPatron = "VIE",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2027, 3, 19), new DateTime(2027, 4, 2), new DateTime(2027, 4, 9) },
                FechaInicioEstimada = new DateTime(2027, 3, 19),
                FechaFinEstimada = new DateTime(2027, 4, 9),
                EsVacante = 0
            });

            // Validador inicial marca HARD=1 por R5
            var validador = new AgValidadorReglas();
            validador.Evaluar(ind, ctx);
            ind.HardViolations.Should().Be(1, string.Join(" | ", ind.Violaciones.Select(v => v.Regla + ":" + v.Mensaje)));

            // Act: Usar evaluador real estandarizado
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorReal);

            // Assert
            res.HuboMejora.Should().BeTrue();
            res.HardPre.Should().Be(1);
            res.HardPost.Should().Be(0);
            res.Individuo.HardViolations.Should().Be(0);
            res.Individuo.Vacantes.Should().Be(ind.Vacantes);

            // Asignatura 150 (LEGADO) fue desplazada a la siguiente fecha disponible libre (2027-04-02)
            var gen150 = res.Individuo.Genes.FirstOrDefault(g => g.IdAsignatura == 150);
            gen150.Should().NotBeNull();
            gen150.FechaInicioEstimada.Value.Should().Be(new DateTime(2027, 4, 2));
        }

        [Fact]
        public void Reparar_SinRanurasNiLiberables_NoAplicaYHaceRollback()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 3, 20); // Ventana estrecha sin espacio
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 1);

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                IdTercero = 10,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12),
                EsFijo = 1 // No liberable
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110,
                IdAula = 1,
                IdTercero = 11,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12),
                EsFijo = 1 // No liberable
            });

            new AgValidadorReglas().Evaluar(ind, ctx);
            int hardInicial = ind.HardViolations;

            // Act
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorReal);

            // Assert
            res.HuboMejora.Should().BeFalse();
            res.HardPost.Should().Be(hardInicial);
            res.Individuo.Should().BeSameAs(ind);
        }

        [Fact]
        public void Reparar_ConAsignaturaFija_NoMueveAsignaturaFija()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 1);
            ctx.ConstruirIndices();

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                IdTercero = 10,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12),
                EsFijo = 1 // EsFijo debe ser estrictamente respetado
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110,
                IdAula = 1,
                IdTercero = 11,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12)
            });

            new AgValidadorReglas().Evaluar(ind, ctx);

            // Act
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorReal);

            // Assert
            var genFijo = res.Individuo.Genes.FirstOrDefault(g => g.IdAsignatura == 100);
            genFijo.Should().NotBeNull();
            genFijo.FechaInicioEstimada.Should().Be(new DateTime(2026, 3, 5));
            genFijo.FechaFinEstimada.Should().Be(new DateTime(2026, 3, 12));
        }

        [Fact]
        public void Reparar_NoCreaVacantesNuevas()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 2);

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                IdTercero = 10,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12)
            });

            new AgValidadorReglas().Evaluar(ind, ctx);
            int vacantesPre = ind.Vacantes;

            // Act
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorReal);

            // Assert
            res.Individuo.Vacantes.Should().BeLessOrEqualTo(vacantesPre);
        }

        [Fact]
        public void Validar_CuandoIntroduceR6MismoCarril_RetornaFalse()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin);

            // Fase 2350: Asignar fechas y bloques disjuntos para aislar R6 sin disparar R2 de aula
            var orig = new AgIndividuo();
            orig.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                DiasPatron = "JUE",
                IdBloqueHorario = 1,
                DiaSemana = 4,
                HorasGen = 4.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12)
            });

            var clon = orig.Clonar();
            // Introducir solape R6 dentro del mismo carril (aula 1, DiasPatron JUE) sin coincidencia de fecha exacta
            clon.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110,
                IdAula = 1,
                DiasPatron = "JUE",
                IdBloqueHorario = 1,
                DiaSemana = 4,
                HorasGen = 4.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 12) }, // Fecha distinta a 03-05 (sin choque R2)
                FechaInicioEstimada = new DateTime(2026, 3, 8), // Solape estricto de intervalo: 3-8 < 3-12
                FechaFinEstimada = new DateTime(2026, 3, 15)
            });

            // Act
            bool esValido = AgReparadorR5.Validar(clon, orig, ctx);

            // Assert
            esValido.Should().BeFalse();
        }

        [Fact]
        public void Puntaje_PrefiereMenorDesperdicioAnteMismoHard()
        {
            var indA = new AgIndividuo { HardViolations = 0, SoftViolations = 10.0 };
            var indB = new AgIndividuo { HardViolations = 0, SoftViolations = 10.0 };

            var p1 = new AgReparadorR5.AgPuntajeReparacion(indA, desperdicio: 0.15, movimientos: 1);
            var p2 = new AgReparadorR5.AgPuntajeReparacion(indB, desperdicio: 0.35, movimientos: 1);

            p1.CompareTo(p2).Should().BeNegative();
        }

        [Fact]
        public void Reparar_EsDeterminista_MismaEntradaProduceMismoResultado()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 2);

            var ind1 = new AgIndividuo();
            ind1.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12)
            });

            var ind2 = ind1.Clonar();

            var res1 = AgReparadorR5.Ejecutar(ind1, ctx, _evaluadorReal);
            var res2 = AgReparadorR5.Ejecutar(ind2, ctx, _evaluadorReal);

            res1.HardPost.Should().Be(res2.HardPost);
            res1.DesperdicioPost.Should().Be(res2.DesperdicioPost);
            res1.HuboMejora.Should().Be(res2.HuboMejora);

            // Fase 2350: Verificación estricta de paridad de fechas en genes
            res1.Individuo.Genes.Select(g => string.Join("|", g.FechasGen.Select(d => d.ToString("yyyy-MM-dd"))))
                .Should().Equal(res2.Individuo.Genes.Select(g => string.Join("|", g.FechasGen.Select(d => d.ToString("yyyy-MM-dd")))));
        }

        [Fact]
        public void Reparar_SobrecargaDirecta_InvocaEvaluadorYAcutalizaAptitud()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 2);

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                IdBloqueHorario = 1,
                DiaSemana = 4,
                HorasGen = 8.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5), new DateTime(2026, 3, 12) },
                FechaInicioEstimada = new DateTime(2026, 3, 5),
                FechaFinEstimada = new DateTime(2026, 3, 12)
            });

            // Act
            var reparado = AgReparadorR5.Reparar(ind, ctx, _evaluadorReal);

            // Assert
            reparado.Should().NotBeNull();
        }
    }
}
