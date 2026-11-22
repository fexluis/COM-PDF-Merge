using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgReparadorR5Tests
    {
        private readonly Mock<IAgEvaluadorAptitud> _evaluadorMock;

        public AgReparadorR5Tests()
        {
            _evaluadorMock = new Mock<IAgEvaluadorAptitud>();
            _evaluadorMock
                .Setup(e => e.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                {
                    // Reevaluar hard violations usando el validador de reglas real
                    var validador = new AgValidadorReglas();
                    validador.Evaluar(ind, ctx);
                });
        }

        private AgContexto CrearContextoBase(DateTime fIni, DateTime fFin, int maxCarriles = 2)
        {
            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = fIni,
                    FechaFin = fFin,
                    PeriodoNombre = "2026-1",
                    Origen = AgOrigenVentana.Default
                },
                Aulas = new List<AgAula>
                {
                    new AgAula
                    {
                        Id = 1,
                        Nombre = "Aula 101",
                        DiasVentana = maxCarriles == 2 ? "LUN,MAR,MIE,JUE" : "LUN,MAR",
                        Jornada = "MANANA",
                        HorasDia = 4
                    }
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
                    new Asignatura { Id = 100, Horas = 12, Semestre = "I", Nombre = "Asig 100" },
                    new Asignatura { Id = 110, Horas = 12, Semestre = "I", Nombre = "Asig 110" },
                    new Asignatura { Id = 150, Horas = 12, Semestre = "I", Nombre = "Asig 150" }
                },
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 100, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 110, OrigenPatron = "LEGADO", Requerida = 1 },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 150, OrigenPatron = "AUTOMATICO", Requerida = 1 }
                }
            };
            ctx.RegistrarDocenteHabilitado(10, 0, 100, "I");
            ctx.RegistrarDocenteHabilitado(11, 0, 110, "I");
            ctx.RegistrarDocenteHabilitado(12, 0, 150, "I");
            ctx.ConstruirIndices();
            return ctx;
        }

        [Fact]
        public void Reparar_Fixture132Sintetico_DesplazaAsignatura150YReduceHardAUnoACero()
        {
            // Arrange: 3 asignaturas en días distintos (LUN, MAR, MIE) para evitar conflicto de aula.
            // Asig 100 (Lunes): 2026-03-02 a 2026-03-16 (3 semanas: 03-02, 03-09, 03-16)
            // Asig 110 (Martes): 2026-03-10 a 2026-03-24 (3 semanas: 03-10, 03-17, 03-24)
            // Asig 150 (Miércoles): 2026-03-11 a 2026-03-25 (3 semanas: 03-11, 03-18, 03-25)
            // Entre 2026-03-11 y 2026-03-16 coinciden las 3 asignaturas -> sobrecarga R5 (3 > M=2).
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
                DiaSemana = 1,
                DiasPatron = "P1",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 2), new DateTime(2026, 3, 9), new DateTime(2026, 3, 16) },
                FechaInicioEstimada = new DateTime(2026, 3, 2),
                FechaFinEstimada = new DateTime(2026, 3, 16),
                EsVacante = 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110,
                IdAula = 1,
                IdTercero = 11,
                IdBloqueHorario = 2,
                DiaSemana = 2,
                DiasPatron = "P2",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 10), new DateTime(2026, 3, 17), new DateTime(2026, 3, 24) },
                FechaInicioEstimada = new DateTime(2026, 3, 10),
                FechaFinEstimada = new DateTime(2026, 3, 24),
                EsVacante = 0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 3,
                IdAsignatura = 150,
                IdAula = 1,
                IdTercero = 12,
                IdBloqueHorario = 3,
                DiaSemana = 3,
                DiasPatron = "P3",
                Semestre = "I",
                Grupo = "A",
                TotalHorasAsignatura = 12.0,
                HorasGen = 12.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 11), new DateTime(2026, 3, 18), new DateTime(2026, 3, 25) },
                FechaInicioEstimada = new DateTime(2026, 3, 11),
                FechaFinEstimada = new DateTime(2026, 3, 25),
                EsVacante = 0
            });

            // Validador inicial marca HARD=1
            var validador = new AgValidadorReglas();
            validador.Evaluar(ind, ctx);
            ind.HardViolations.Should().Be(1, string.Join(" | ", ind.Violaciones.Select(v => v.Regla + ":" + v.Mensaje)));

            // Act
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorMock.Object);

            // Assert
            res.HuboMejora.Should().BeTrue();
            res.HardPre.Should().Be(1);
            res.HardPost.Should().Be(0);
            res.Individuo.HardViolations.Should().Be(0);

            // Asignatura 150 fue desplazada posteriormente a la finalización de asig 100
            var gen150 = res.Individuo.Genes.FirstOrDefault(g => g.IdAsignatura == 150);
            gen150.Should().NotBeNull();
            gen150.FechaInicioEstimada.Value.Should().BeOnOrAfter(new DateTime(2026, 3, 17));
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
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorMock.Object);

            // Assert
            res.HuboMejora.Should().BeFalse();
            res.HardPost.Should().Be(hardInicial);
            res.Individuo.Should().BeSameAs(ind);
        }

        [Fact]
        public void Reparar_ConAsignaturaLegado_NoMueveAsignaturaLegado()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin, maxCarriles: 1);
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.ConstruirIndices();

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100, // LEGADO
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
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110, // AUTOMATICO (puede moverse)
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
            int hardInicial = ind.HardViolations;

            // Act
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorMock.Object);

            // Assert
            var genLegado = res.Individuo.Genes.FirstOrDefault(g => g.IdAsignatura == 100);
            genLegado.Should().NotBeNull();
            genLegado.FechaInicioEstimada.Should().Be(new DateTime(2026, 3, 5));
            genLegado.FechaFinEstimada.Should().Be(new DateTime(2026, 3, 12));
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
            var res = AgReparadorR5.Ejecutar(ind, ctx, _evaluadorMock.Object);

            // Assert
            res.Individuo.Vacantes.Should().BeLessOrEqualTo(vacantesPre);
        }

        [Fact]
        public void Validar_CuandoIntroduceR6MismoCarril_RetornaFalse()
        {
            var fIni = new DateTime(2026, 3, 1);
            var fFin = new DateTime(2026, 5, 30);
            var ctx = CrearContextoBase(fIni, fFin);

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
            // Introducir solape R6 dentro del mismo carril
            clon.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 110,
                IdAula = 1,
                DiasPatron = "JUE",
                IdBloqueHorario = 1,
                DiaSemana = 4,
                HorasGen = 4.0,
                FechasGen = new List<DateTime> { new DateTime(2026, 3, 5) },
                FechaInicioEstimada = new DateTime(2026, 3, 8), // Solape estricto 3-8 < 3-12
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

            var res1 = AgReparadorR5.Ejecutar(ind1, ctx, _evaluadorMock.Object);
            var res2 = AgReparadorR5.Ejecutar(ind2, ctx, _evaluadorMock.Object);

            res1.HardPost.Should().Be(res2.HardPost);
            res1.DesperdicioPost.Should().Be(res2.DesperdicioPost);
            res1.HuboMejora.Should().Be(res2.HuboMejora);
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
            var reparado = AgReparadorR5.Reparar(ind, ctx, _evaluadorMock.Object);

            // Assert
            reparado.Should().NotBeNull();
        }
    }
}
