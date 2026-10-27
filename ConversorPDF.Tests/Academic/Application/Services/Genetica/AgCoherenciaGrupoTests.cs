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
    public class AgCoherenciaGrupoTests
    {
        [Fact]
        public void Sellar_WhenGrupoHasExcessHours_ShouldRedistributeCanonicallyAndSatisfyI1()
        {
            // Arrange: grupo con 3 genes de 24h cada uno (total = 72h) pero TotalHorasAsignatura = 48h
            var ind = new AgIndividuo();
            var g1 = new AgGen { Id = 1, ClaveGrupo = "1|10|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 0 };
            var g2 = new AgGen { Id = 2, ClaveGrupo = "1|10|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 1 };
            var g3 = new AgGen { Id = 3, ClaveGrupo = "1|10|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 2 };

            ind.Genes.Add(g1);
            ind.Genes.Add(g2);
            ind.Genes.Add(g3);

            // Act: Fase 562 & 563
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert: 48h = 12 sesiones repartidas entre 3 genes (4 sesiones c/u = 16h)
            double suma = g1.HorasGen + g2.HorasGen + g3.HorasGen;
            suma.Should().Be(48.0);
            g1.HorasGen.Should().Be(16.0);
            g2.HorasGen.Should().Be(16.0);
            g3.HorasGen.Should().Be(16.0);
        }

        [Fact]
        public void Sellar_WhenGrupoHasDeficitHours_ShouldRedistributeCanonicallyAndSatisfyI1()
        {
            // Arrange: grupo con 2 genes de 16h cada uno (total = 32h) pero TotalHorasAsignatura = 48h
            var ind = new AgIndividuo();
            var g1 = new AgGen { Id = 1, ClaveGrupo = "2|10|A", TotalHorasAsignatura = 48.0, HorasGen = 16.0, IndiceDiaPatron = 0 };
            var g2 = new AgGen { Id = 2, ClaveGrupo = "2|10|A", TotalHorasAsignatura = 48.0, HorasGen = 16.0, IndiceDiaPatron = 1 };

            ind.Genes.Add(g1);
            ind.Genes.Add(g2);

            // Act
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert: 48h = 12 sesiones repartidas entre 2 genes (6 sesiones c/u = 24h)
            double suma = g1.HorasGen + g2.HorasGen;
            suma.Should().Be(48.0);
            g1.HorasGen.Should().Be(24.0);
            g2.HorasGen.Should().Be(24.0);
        }

        [Fact]
        public void Sellar_WhenGrupoHasFixedGene_ShouldPreserveFixedGeneAndAdjustOnlyMutableGenes()
        {
            // Arrange: 1 gen fijado de 16h, total esperado 48h -> 32h restantes para los otros 2 genes
            var ind = new AgIndividuo();
            var gFijo = new AgGen { Id = 1, ClaveGrupo = "3|10|A", TotalHorasAsignatura = 48.0, HorasGen = 16.0, Estado = AgEstadoGen.Fijado, IndiceDiaPatron = 0 };
            var gMut1 = new AgGen { Id = 2, ClaveGrupo = "3|10|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, Estado = AgEstadoGen.Valido, IndiceDiaPatron = 1 };
            var gMut2 = new AgGen { Id = 3, ClaveGrupo = "3|10|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, Estado = AgEstadoGen.Valido, IndiceDiaPatron = 2 };

            ind.Genes.Add(gFijo);
            ind.Genes.Add(gMut1);
            ind.Genes.Add(gMut2);

            // Act
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert
            gFijo.HorasGen.Should().Be(16.0);
            (gMut1.HorasGen + gMut2.HorasGen).Should().Be(32.0);
            (gFijo.HorasGen + gMut1.HorasGen + gMut2.HorasGen).Should().Be(48.0);
        }

        [Fact]
        public void Sellar_WhenInheritedGeneHasZeroHoursOnCompleteGroup_ShouldRedistribute()
        {
            // Arrange: 2 genes heredados a 0 h + 1 gen con 48 h en el mismo ClaveGrupo (Fase 566 / 570)
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, ClaveGrupo = "9|10|A", TotalHorasAsignatura = 48.0, HorasGen = 0.0, IndiceDiaPatron = 0 });
            ind.Genes.Add(new AgGen { Id = 2, ClaveGrupo = "9|10|A", TotalHorasAsignatura = 48.0, HorasGen = 48.0, IndiceDiaPatron = 1 });

            // Act
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert
            ind.Genes.Sum(g => g.HorasGen).Should().Be(48.0);
            ind.Genes.Should().OnlyContain(g => g.HorasGen > 0.0);
        }

        [Fact]
        public void Sellar_WhenFewerSessionsThanGenes_ShouldConsolidateAndBeIdempotent()
        {
            // Arrange: 4 h objetivo y 3 slots => 1 sesión en 1 gen, 2 genes a 0 h (Fase 572 & 580)
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, ClaveGrupo = "20|1|A", TotalHorasAsignatura = 4.0, HorasGen = 4.0, IndiceDiaPatron = 0 });
            ind.Genes.Add(new AgGen { Id = 2, ClaveGrupo = "20|1|A", TotalHorasAsignatura = 4.0, HorasGen = 4.0, IndiceDiaPatron = 1 });
            ind.Genes.Add(new AgGen { Id = 3, ClaveGrupo = "20|1|A", TotalHorasAsignatura = 4.0, HorasGen = 4.0, IndiceDiaPatron = 2 });

            // Act
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert
            ind.Genes.Sum(g => g.HorasGen).Should().Be(4.0);

            var antes = ind.Genes.Select(g => g.HorasGen).ToList();
            AgCoherenciaGrupo.Sellar(ind, null); // idempotencia
            ind.Genes.Select(g => g.HorasGen).Should().Equal(antes);
        }

        [Fact]
        public void Sellar_WhenDeficitExceedsSingleGeneWindow_ShouldDistributeAcrossAllGenesWithSlack()
        {
            // Arrange (Fase 585): 2 genes mismo ClaveGrupo, TotalHorasAsignatura=48, HorasGen=0
            var ind = new AgIndividuo();
            var fIni = new DateTime(2026, 2, 1);
            var fFin = new DateTime(2026, 6, 30);
            ind.Genes.Add(new AgGen { Id = 1, ClaveGrupo = "50|1|A", TotalHorasAsignatura = 48.0, HorasGen = 0.0, IndiceDiaPatron = 0, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });
            ind.Genes.Add(new AgGen { Id = 2, ClaveGrupo = "50|1|A", TotalHorasAsignatura = 48.0, HorasGen = 0.0, IndiceDiaPatron = 1, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });

            var ctx = new AgContexto
            {
                IdCorrida = 101,
                Ventana = new VentanaSemestre { FechaInicio = fIni, FechaFin = fFin }
            };

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            ind.Genes.Sum(g => g.HorasGen).Should().Be(48.0);
            ind.Genes.Should().OnlyContain(g => g.HorasGen >= 0.0);

            // Idempotencia
            var snapshot = ind.Genes.Select(g => g.HorasGen).ToList();
            AgCoherenciaGrupo.Sellar(ind, ctx);
            ind.Genes.Select(g => g.HorasGen).Should().Equal(snapshot);
        }

        [Fact]
        public void Sellar_WhenMultipleGenes_ShouldLimitEachGeneCapacityToItsOwnDayOccurrences()
        {
            // Arrange (Fase 590): 3 genes (martes=2, miercoles=3, jueves=4) en ventana de 3 semanas (3 ocurrencias c/u)
            var ind = new AgIndividuo();
            var fIni = new DateTime(2026, 2, 2);
            var fFin = new DateTime(2026, 2, 20); // 3 semanas completas: 3 martes, 3 miércoles, 3 jueves
            ind.Genes.Add(new AgGen { Id = 1, ClaveGrupo = "60|1|A", TotalHorasAsignatura = 36.0, HorasGen = 0.0, DiaSemana = 2, IndiceDiaPatron = 0, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });
            ind.Genes.Add(new AgGen { Id = 2, ClaveGrupo = "60|1|A", TotalHorasAsignatura = 36.0, HorasGen = 0.0, DiaSemana = 3, IndiceDiaPatron = 1, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });
            ind.Genes.Add(new AgGen { Id = 3, ClaveGrupo = "60|1|A", TotalHorasAsignatura = 36.0, HorasGen = 0.0, DiaSemana = 4, IndiceDiaPatron = 2, FechaInicioEstimada = fIni, FechaFinEstimada = fFin });

            var ctx = new AgContexto
            {
                IdCorrida = 202,
                Ventana = new VentanaSemestre { FechaInicio = fIni, FechaFin = fFin }
            };

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert: se reparten 36h (9 sesiones) respetando el tope de 3 sesiones (12h) por gen
            ind.Genes.Sum(g => g.HorasGen).Should().Be(36.0);
            ind.Genes.Should().OnlyContain(g => (int)Math.Round(g.HorasGen / 4.0) <= 3);
        }

        [Fact]
        public void Sellar_WhenAllGenesFixedAndSumExceedsExpected_ShouldRescaleToSatisfyI1()
        {
            // Arrange: Caso H5 (Grupo 51|285|A con fijados=96h y esperado=48h)
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                ClaveGrupo = "51|285|A",
                TotalHorasAsignatura = 48.0,
                HorasGen = 96.0,
                Estado = AgEstadoGen.Fijado,
                FechasGen = Enumerable.Range(1, 24).Select(i => new DateTime(2026, 3, 1).AddDays(i * 7)).ToList()
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, null);

            // Assert: se reescala a 48h y se limpian las fechas residuales
            ind.Genes.Sum(g => g.HorasGen).Should().Be(48.0);
            ind.Genes[0].HorasGen.Should().Be(48.0);
            ind.Genes[0].FechasGen.Should().BeEmpty();
        }

        [Fact]
        public void Sellar_WhenAllGenesFixedAndAula8Hours_ShouldRescaleIn8HourMultiples()
        {
            // Arrange (Fase 901): aula con HorasDia = 8, 2 genes fijados con 32h cada uno (total 64h), esperado 48h
            var ctx = new AgContexto();
            ctx.AulasPorId[10] = new AgAula { Id = 10, HorasDia = 8, Codigo = "AULA-8H" };

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, IdAula = 10, ClaveGrupo = "70|10|A", TotalHorasAsignatura = 48.0, HorasGen = 32.0, Estado = AgEstadoGen.Fijado });
            ind.Genes.Add(new AgGen { Id = 2, IdAula = 10, ClaveGrupo = "70|10|A", TotalHorasAsignatura = 48.0, HorasGen = 32.0, Estado = AgEstadoGen.Fijado });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert: se reescala a múltiplos de 8h (24h c/u = 3 sesiones de 8h) sumando exactamente 48h
            ind.Genes.Sum(g => g.HorasGen).Should().Be(48.0);
            ind.Genes.Should().OnlyContain(g => g.HorasGen % 8.0 == 0);
        }

        [Fact]
        public void Sellar_GrupoMixtoConSinPlanOleada_PreservaMotivoYVacante()
        {
            // Arrange (Fase 1091): Un grupo con 1 gen con docente y 1 gen con SIN_PLAN_OLEADA (e.g. cruce mixto)
            // Sellar debe preservar la vacante y el motivo SIN_PLAN_OLEADA en todo el grupo sin sobreescribirlo a real.
            var ctx = new AgContexto();
            ctx.IdTerceroVacante = 999;

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                ClaveGrupo = "1|10|A",
                IdTercero = 50,
                EsVacante = 0,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                ClaveGrupo = "1|10|A",
                IdTercero = 0,
                EsVacante = 1,
                MotivoVacante = ConversorPDF.XLL.Academic.Domain.Enums.AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            ind.Genes.Should().OnlyContain(g => g.EsVacante == 1);
            ind.Genes.Should().OnlyContain(g => g.MotivoVacante == ConversorPDF.XLL.Academic.Domain.Enums.AgMotivoVacante.SIN_PLAN_OLEADA.Codigo());
            ind.Genes.Should().OnlyContain(g => g.IdTercero == 999);
        }

        [Fact]
        public void Sellar_GrupoMixtoConSinBloqueNivel_PreservaMotivoYVacante()
        {
            // Arrange (Fase 1104 / Fase 1107): Un grupo con 1 gen con docente y 1 gen con SIN_BLOQUE_NIVEL (e.g. cruce mixto)
            // Sellar debe preservar la vacante y el motivo SIN_BLOQUE_NIVEL en todo el grupo sin sobreescribirlo a real.
            var ctx = new AgContexto();
            ctx.IdTerceroVacante = 999;

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                ClaveGrupo = "1|10|A",
                IdTercero = 50,
                EsVacante = 0,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });
            ind.Genes.Add(new AgGen
            {
                Id = 2,
                ClaveGrupo = "1|10|A",
                IdTercero = 0,
                EsVacante = 1,
                MotivoVacante = ConversorPDF.XLL.Academic.Domain.Enums.AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo(),
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            ind.Genes.Should().OnlyContain(g => g.EsVacante == 1);
            ind.Genes.Should().OnlyContain(g => g.MotivoVacante == ConversorPDF.XLL.Academic.Domain.Enums.AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo());
            ind.Genes.Should().OnlyContain(g => g.IdTercero == 999);
        }

        [Fact]
        public void Sellar_PregradoConBloque8h_NormalizaA4hYReasignaBloqueDelMismoDia()
        {
            // Arrange (Fase 1114.a / Fase 1114.d): Gen de aula de PREGRADO (IdPrograma 1, Presencial) con HorasSemanales = 8h y bloque 8h.
            // Sellar debe normalizar HorasSemanales a 4h y reasignar el bloque de 4h del mismo día.
            var ctx = new AgContexto();
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            ctx.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, HorasDia = 4 };
            ctx.Bloques.Add(new AgBloqueHorario { Id = 80, Activo = 1, DiaSemanaNumero = 2, DuracionHoras = 8.0 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 40, Activo = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0 });

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAula = 10,
                ClaveGrupo = "1|10|A",
                DiaSemana = 2,
                HorasSemanales = 8.0,
                IdBloqueHorario = 80,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(4.0);
            g.IdBloqueHorario.Should().Be(40);
        }

        [Fact]
        public void Sellar_PregradoSinBloque4h_MarcaSinBloqueNivelYNoAsigna8h()
        {
            // Arrange (Fase 1114.a / Fase 1114.d): Gen de PREGRADO con HorasSemanales = 8h en un día sin bloque de 4h.
            // Sellar debe ponerlo en cuarentena (IdBloqueHorario = 0, EsVacante = 1, SIN_BLOQUE_NIVEL).
            var ctx = new AgContexto();
            ctx.IdTerceroVacante = 999;
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            ctx.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, HorasDia = 4 };
            // Sólo existe bloque de 8h, no hay de 4h para día 2
            ctx.Bloques.Add(new AgBloqueHorario { Id = 80, Activo = 1, DiaSemanaNumero = 2, DuracionHoras = 8.0 });

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAula = 10,
                ClaveGrupo = "1|10|A",
                DiaSemana = 2,
                HorasSemanales = 8.0,
                IdBloqueHorario = 80,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(4.0);
            g.IdBloqueHorario.Should().Be(0);
            g.EsVacante.Should().Be(1);
            g.MotivoVacante.Should().Be(ConversorPDF.XLL.Academic.Domain.Enums.AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo());
        }

        [Fact]
        public void Sellar_PosgradoConBloque8h_NoDegradaDuracion()
        {
            // Arrange (Fase 1118 / Fase 1119): Gen de aula de POSGRADO con HorasSemanales = 8h y bloque 8h.
            // Sellar debe respetar la unidad del nivel de POSGRADO (8h) y no degradar a 4h.
            var ctx = new AgContexto();
            ctx.Programas.Add(new Programa { Id = 2, Modalidad = "POSGRADO ESPECIALIZACION" });
            ctx.AulasPorId[20] = new AgAula { Id = 20, IdPrograma = 2, HorasDia = 8 };
            ctx.Bloques.Add(new AgBloqueHorario { Id = 80, Activo = 1, DiaSemanaNumero = 5, DuracionHoras = 8.0 });

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAula = 20,
                ClaveGrupo = "2|20|A",
                DiaSemana = 5,
                HorasSemanales = 8.0,
                IdBloqueHorario = 80,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0,
                Estado = AgEstadoGen.Valido
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(8.0);
            g.IdBloqueHorario.Should().Be(80);
            g.EsVacante.Should().Be(0);
        }

        [Fact]
        public void Sellar_GenSinHorasSemanales_NoModificaFechasNiEstado()
        {
            // Arrange (Fase 1117 / Fase 1119): Gen con HorasSemanales = 0.0 (dato desconocido/no asignado).
            // Sellar no debe tratar 0 como sobre-duración ni borrar fechas precalculadas ni enviar a vacante.
            var ctx = new AgContexto();
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            ctx.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, HorasDia = 4 };

            var fIni = new DateTime(2026, 2, 1);
            var fFin = new DateTime(2026, 6, 30);
            var ind = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAula = 10,
                ClaveGrupo = "1|10|A",
                DiaSemana = 2,
                HorasSemanales = 0.0,
                IdBloqueHorario = 40,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0,
                FechaInicioEstimada = fIni,
                FechaFinEstimada = fFin,
                Estado = AgEstadoGen.Valido
            };
            gen.FechasGen.Add(new DateTime(2026, 2, 3));
            ind.Genes.Add(gen);

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(0.0);
            g.FechasGen.Should().NotBeEmpty();
            g.IdBloqueHorario.Should().Be(40);
            g.EsVacante.Should().Be(0);
        }

        [Fact]
        public void Sellar_AulaNoIndexada_NoAlteraDuracion()
        {
            // Arrange (Fase 1118 / Fase 1119): Aula no indexada en ctx.AulasPorId (IdAula = 999).
            // UnidadNivel retorna -1.0; Sellar no debe asumir PREGRADO (4h) ni alterar la duración del gen.
            var ctx = new AgContexto();
            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAula = 999,
                ClaveGrupo = "1|999|A",
                DiaSemana = 2,
                HorasSemanales = 8.0,
                IdBloqueHorario = 80,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0,
                Estado = AgEstadoGen.Valido
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(8.0);
            g.IdBloqueHorario.Should().Be(80);
            g.EsVacante.Should().Be(0);
        }

        [Fact]
        public void Sellar_PregradoTarde8h_ReasignaBloqueDeFranjaTarde()
        {
            // Arrange (Fase 1127 / Fase 1130): Aula de PREGRADO en jornada TARDE.
            // Existen dos bloques de 4h para el día 2: MANANA (Id 10) y TARDE (Id 20).
            // Un gen llega con 8h; Sellar debe normalizar a 4h y seleccionar el bloque de la franja TARDE (Id 20).
            var ctx = new AgContexto();
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            ctx.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, HorasDia = 4, Jornada = "TARDE" };
            ctx.Bloques.Add(new AgBloqueHorario { Id = 10, Activo = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Franja = "MANANA" });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 20, Activo = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Franja = "TARDE" });

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAula = 10,
                ClaveGrupo = "1|10|A",
                DiaSemana = 2,
                HorasSemanales = 8.0,
                IdBloqueHorario = 99,
                TotalHorasAsignatura = 32.0,
                HorasGen = 16.0,
                Estado = AgEstadoGen.Valido
            });

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert: se reasigna a 4h y franja TARDE (Id 20)
            var g = ind.Genes[0];
            g.HorasSemanales.Should().Be(4.0);
            g.IdBloqueHorario.Should().Be(20);
        }

        [Fact]
        public void Contexto_TrasInvalidarIndiceBloques_UsaCatalogoVigente()
        {
            // Arrange (Fase 1128 / Fase 1130): Catálogo inicial con bloque 10 (MANANA).
            var ctx = new AgContexto();
            ctx.Bloques.Add(new AgBloqueHorario { Id = 10, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "MANANA" });

            // Primera consulta inicializa caché
            ctx.TryGetBloque(1, 4.0, out var b1).Should().BeTrue();
            b1.Id.Should().Be(10);

            // Modificación de catálogo: se añade bloque TARDE
            ctx.Bloques.Add(new AgBloqueHorario { Id = 20, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "TARDE" });

            // Sin invalidar, la búsqueda por franja "TARDE" aún no conoce el bloque 20
            // Act: invalidar índice de bloques
            ctx.InvalidarIndiceBloques();

            // Assert: ahora debe resolver el nuevo bloque
            ctx.TryGetBloque(1, 4.0, "TARDE", out var b2).Should().BeTrue();
            b2.Id.Should().Be(20);
        }

        [Fact]
        public void Sellar_SabadoUnGenFijadoYUnLibre_MismaFranja_DesdoblaFranja()
        {
            // Arrange (Fase 1307): Grupo con 1 gen fijado y 1 gen libre en sábado con la misma franja
            var ind = new AgIndividuo();
            var fijado = new AgGen { Id = 1, DiaSemana = 6, Estado = AgEstadoGen.Fijado, IdAula = 1, HorasSemanales = 4, HorasGen = 24, TotalHorasAsignatura = 48, IdBloqueHorario = 3, ClaveGrupo = "G1" };
            var libre = new AgGen { Id = 2, DiaSemana = 6, Estado = AgEstadoGen.Valido, IdAula = 1, HorasSemanales = 4, HorasGen = 24, TotalHorasAsignatura = 48, IdBloqueHorario = 3, ClaveGrupo = "G1", IndiceDiaPatron = 1 };
            ind.Genes.Add(fijado);
            ind.Genes.Add(libre);

            var ctx = new AgContexto();
            ctx.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Jornada = "DIURNA", HorasDia = 4 });   // Fase 1313
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, DuracionHoras = 4, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA", Activo = 1 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 6, DuracionHoras = 4, HoraInicio = "13:00", HoraFin = "17:00", Franja = "TARDE", Activo = 1 });
            ctx.ConstruirIndices();

            // Act
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert: el gen libre se desdobla a TARDE (4) y el fijado se preserva en MANANA (3)
            libre.IdBloqueHorario.Should().Be(4);
            fijado.IdBloqueHorario.Should().Be(3);
        }
    }
}
