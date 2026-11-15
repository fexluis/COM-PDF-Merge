using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using CSharpFunctionalExtensions;
using Dapper;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Pruebas puras de dominio y aplicación para la redistribución de ranuras,
    /// evaluación de ventanas, cálculo de déficit, aislamiento de autorización y coherencia temporal.
    /// Ejecución rápida en memoria sin acceso a base de datos.
    /// </summary>
    public class AgRedistribucionRanurasAceptacionTests
    {
        // =========================================================================
        // Caso 1: LUN-JUE / MAR-VIE -> Proponer LUN-VIE cuando mejora cobertura
        // =========================================================================
        [Theory]
        [InlineData("LUN-JUE")]
        [InlineData("MAR-VIE")]
        public void Evaluar_WhenLunJueOMarVieConSobredemanda_ShouldProponerLunVie(string ventanaInicial)
        {
            // Arrange: 17 semanas. 4 días * 1 ranura = 68 ranuras. Demanda = 5 materias * 16 sesiones = 80 sesiones.
            var ctx = CrearContextoSintetico(17, ventanaInicial, numAsignaturas: 5, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(ctx);

            // Assert
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be(ventanaInicial);
            d.Hacia.Should().Be("LUN-VIE");
            d.CapacidadAntes.Should().Be(68);
            d.CapacidadDespues.Should().Be(85);
            d.DemandaSesiones.Should().Be(80);
            d.Aplica.Should().BeTrue();
        }

        // =========================================================================
        // Caso 2: VIE-DOM -> JUE-DOM -> MIE-DOM; máximo dos pasos
        // =========================================================================
        [Fact]
        public void Evaluar_WhenVieDomDemanda6W_ShouldEscalarEnMaximoDosPasos()
        {
            // Arrange: 17 semanas, VIE-DOM (4 ranuras/sem = 68 ranuras). Demanda = 6 materias * 16 ses = 96 sesiones.
            var ctx = CrearContextoSintetico(17, "VIE-DOM", numAsignaturas: 6, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Primer paso
            var decPaso1 = sut.Evaluar(ctx);

            // Assert 1: Propone exclusivamente JUE-DOM
            decPaso1.Should().ContainSingle();
            var d1 = decPaso1[0];
            d1.Desde.Should().Be("VIE-DOM");
            d1.Hacia.Should().Be("JUE-DOM");
            d1.CapacidadAntes.Should().Be(68);
            d1.CapacidadDespues.Should().Be(85);
            d1.DemandaSesiones.Should().Be(96);
            d1.DemandaPendiente.Should().Be(11);

            // Act 2: Aceptar paso 1 y re-evaluar
            ctx.Aulas[0].DiasVentana = d1.Hacia;
            var decPaso2 = sut.Evaluar(ctx);

            // Assert 2: Segundo paso propone MIE-DOM (alcanza 102 ranuras >= 96 sesiones)
            decPaso2.Should().ContainSingle();
            var d2 = decPaso2[0];
            d2.Desde.Should().Be("JUE-DOM");
            d2.Hacia.Should().Be("MIE-DOM");
            d2.CapacidadAntes.Should().Be(85);
            d2.CapacidadDespues.Should().Be(102);
            d2.DemandaPendiente.Should().Be(0);

            // Act 3: Aceptar paso 2 y re-evaluar
            ctx.Aulas[0].DiasVentana = d2.Hacia;
            var decPaso3 = sut.Evaluar(ctx);

            // Assert 3: Sin más propuestas (estado factible alcanzado en máximo 2 pasos)
            decPaso3.Should().BeEmpty();
        }

        // =========================================================================
        // Caso 3: MIE-DOM -> Terminal
        // =========================================================================
        [Fact]
        public void Evaluar_WhenMieDomTerminal_ShouldNoProponerMasEscalones()
        {
            // Arrange: Aula en MIE-DOM con sobre-demanda extrema (120 sesiones > 102 ranuras)
            var ctx = CrearContextoSintetico(17, "MIE-DOM", numAsignaturas: 8, sesionesPorAsignatura: 15);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(ctx);

            // Assert: MIE-DOM es tope terminal de fin de semana; no inventa escalones adicionales
            decisiones.Should().BeEmpty();
        }

        // =========================================================================
        // Caso 5: Domingo obligatorio: 144 h frente a 64 h disponibles -> Informar déficit de 80 h
        // =========================================================================
        [Fact]
        public void Evaluar_WhenDomingoObligatorio144hFrenteA64h_ShouldInformarDeficitDe80h()
        {
            // Arrange: 16 semanas hasta mayo 2027. Domingo tiene 16 sesiones * 4h = 64h disponibles.
            // 3 asignaturas obligatorias de domingo con 48h cada una = 144h requeridas (36 domingos requeridos).
            var fIni = new DateTime(2027, 2, 7); // Domingo
            var ctx = CrearContextoSintetico(16, "VIE-DOM", numAsignaturas: 3, sesionesPorAsignatura: 12, fechaInicio: fIni);
            foreach (var aa in ctx.AulaAsignaturas)
            {
                aa.DiasPatron = "7";
                aa.OrigenPatron = "FIJO";
            }
            foreach (var kv in ctx.AulaAsignaturaPorAulaAsig)
            {
                kv.Value.DiasPatron = "7";
                kv.Value.OrigenPatron = "FIJO";
            }

            // Act
            var res = AgEvaluadorDemandaVentana.Evaluar(
                "VIE-DOM", "MANANA", 4.0, ctx, ctx.Aulas[0], ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin);

            // Assert
            res.DemandaPendiente.Should().Be(20, "144h demandadas (36 ses) vs 64h de domingo (16 ses) = 20 sesiones pendientes");
            res.DeficitHoras.Should().Be(80.0, "20 sesiones pendientes * 4h = déficit de 80 h");
            res.DeficitSesiones.Should().Be(20);
            res.Motivo.Should().Contain("80 h", "el informe diagnóstico debe expresar explícitamente el déficit de 80 h");
        }

        // =========================================================================
        // Caso 8: Plan fechado -> Mismas fechas y horas en staging
        // =========================================================================
        [Fact]
        public async Task MaterializarEnStaging_WhenPlanFechado_ShouldConservarMismasFechasYHorasEnStaging()
        {
            // Arrange: Dividir el fixture en genes por ranura coherente
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 2, sesionesPorAsignatura: 8);

            var fechasLunes = new List<DateTime>
            {
                new DateTime(2027, 2, 8),  // Lunes
                new DateTime(2027, 2, 15)  // Lunes
            };
            var fechasMiercoles = new List<DateTime>
            {
                new DateTime(2027, 2, 10), // Miércoles
                new DateTime(2027, 2, 17)  // Miércoles
            };

            var individuo = new AgIndividuo { Id = 100 };

            // Gen 1: Ranura Lunes (Bloque 1)
            var gen1 = new AgGen
            {
                Id = 1,
                IdIndividuo = 100,
                IdAula = 1,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                IdTercero = 10,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasSemanales = 4,
                HorasGen = 8,
                TotalHorasAsignatura = 16,
                EsFijoTemporal = true,
                FechasGen = new List<DateTime>(fechasLunes),
                ClaveGrupo = "1|1|A",
                Grupo = "A"
            };

            // Gen 2: Ranura Miércoles (Bloque 3)
            var gen2 = new AgGen
            {
                Id = 2,
                IdIndividuo = 100,
                IdAula = 1,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                IdTercero = 10,
                IdBloqueHorario = 3,
                DiaSemana = 3,
                HorasSemanales = 4,
                HorasGen = 8,
                TotalHorasAsignatura = 16,
                EsFijoTemporal = true,
                FechasGen = new List<DateTime>(fechasMiercoles),
                ClaveGrupo = "1|1|A",
                Grupo = "A"
            };

            individuo.Genes.Add(gen1);
            individuo.Genes.Add(gen2);

            List<AgVinculacion> guardadasVinc = null;
            List<AgSesion> guardadasSes = null;

            var mockRepo = new Mock<IAgScheduleRepository>();
            mockRepo.Setup(r => r.LimpiarStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .Returns(Task.FromResult(Result.Success()));
            mockRepo.Setup(r => r.GuardarStagingAsync(It.IsAny<long>(), It.IsAny<long>(),
                    It.IsAny<IEnumerable<AgVinculacion>>(), It.IsAny<IEnumerable<AgSesion>>(), It.IsAny<CancellationToken>()))
                .Callback<long, long, IEnumerable<AgVinculacion>, IEnumerable<AgSesion>, CancellationToken>(
                    (corr, ind, vincs, ses, token) =>
                    {
                        guardadasVinc = vincs != null ? vincs.ToList() : null;
                        guardadasSes = ses != null ? ses.ToList() : null;
                    })
                .Returns(Task.FromResult(Result.Success()));

            var expansor = new AgExpansorSesiones();
            var materializador = new AgMaterializadorSesiones(mockRepo.Object, expansor);

            // Act
            var res = await materializador.MaterializarEnStagingAsync(1, individuo, ctx, CancellationToken.None);

            // Assert
            res.IsSuccess.Should().BeTrue();
            guardadasSes.Should().NotBeNull();
            guardadasSes.Should().HaveCount(4);

            var todasFechasEsperadas = fechasLunes.Concat(fechasMiercoles).OrderBy(d => d).ToList();
            var fechasStaging = guardadasSes.Select(s => s.FechaSesion.Value.Date).OrderBy(d => d).ToList();
            fechasStaging.Should().Equal(todasFechasEsperadas, "las fechas en staging deben coincidir exactamente con el plan fechado");

            guardadasSes.Sum(s => s.Horas).Should().Be(16.0);
            guardadasSes.Should().OnlyContain(s => s.HorarioSesion == "08:00 A 12:00");
        }

        // =========================================================================
        // Validación de Fechas antes de Materializar (Paso 6)
        // =========================================================================
        [Fact]
        public void Expandir_WhenFechaIncompatibleConDiaBloque_ShouldRetornarFallo()
        {
            // Arrange: Gen asignado a Lunes (Bloque 1) pero FechasGen contiene un Miércoles
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 4);
            var expansor = new AgExpansorSesiones();

            var gen = new AgGen
            {
                Id = 10,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1, // Lunes
                DiaSemana = 1,
                HorasGen = 4,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 10) } // Miércoles
            };

            // Act
            var res = expansor.Expandir(gen, 1, ctx);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("incompatible con el día del bloque");
        }

        [Fact]
        public void Expandir_WhenFechaEsFestiva_ShouldRetornarFallo()
        {
            // Arrange: Fecha festiva configurada en el contexto
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 4);
            var fechaFestiva = new DateTime(2027, 2, 8); // Lunes
            ctx.Festivos.Add((fechaFestiva, fechaFestiva));
            ctx.ConstruirIndices();

            var expansor = new AgExpansorSesiones();
            var gen = new AgGen
            {
                Id = 11,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasGen = 4,
                FechasGen = new List<DateTime> { fechaFestiva }
            };

            // Act
            var res = expansor.Expandir(gen, 1, ctx);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("festivas");
        }

        [Fact]
        public void Expandir_WhenFechasContienenDuplicados_ShouldRetornarFallo()
        {
            // Arrange: FechasGen contiene la misma fecha dos veces
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 4);
            var expansor = new AgExpansorSesiones();

            var gen = new AgGen
            {
                Id = 12,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasGen = 8,
                FechasGen = new List<DateTime>
                {
                    new DateTime(2027, 2, 8),
                    new DateTime(2027, 2, 8)
                }
            };

            // Act
            var res = expansor.Expandir(gen, 1, ctx);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("duplicadas");
        }

        [Fact]
        public void Expandir_WhenCardinalidadDeFechasNoCoincideConSesionesRequeridas_ShouldRetornarFallo()
        {
            // Arrange: HorasGen = 8h (2 sesiones de 4h) pero solo 1 fecha en FechasGen
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 4);
            var expansor = new AgExpansorSesiones();

            var gen = new AgGen
            {
                Id = 13,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasGen = 8,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 8) }
            };

            // Act
            var res = expansor.Expandir(gen, 1, ctx);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("cardinalidad de fechas");
        }

        // =========================================================================
        // Reclasificación accesible en Ventanas Terminales (Paso 2)
        // =========================================================================
        [Fact]
        public void Evaluar_WhenVentanaTerminalConPatronLegado_ShouldIdentificarCandidatosSinExigirEscalarDias()
        {
            // Arrange: Aula en ventana terminal MIE-DOM con patrón LEGADO y déficit real en domingo (20 sesiones vs 17 domingos disponibles)
            var ctx = CrearContextoSintetico(17, "MIE-DOM", numAsignaturas: 2, sesionesPorAsignatura: 20);
            ctx.AulaAsignaturas[0].DiasPatron = "7";
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "7";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(ctx);

            // Assert: En ventana terminal MIE-DOM se identifica la decisión con candidatos de reclasificación
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("MIE-DOM");
            d.Hacia.Should().Be("MIE-DOM", "en ventana terminal no se escala a días superiores");
            d.ReclasificacionesPropuestas.Should().NotBeEmpty();
            d.ReclasificacionesPropuestas.Should().Contain(p => p.IdAsignatura == 1 && p.OrigenAnterior == "LEGADO");
            d.ReclasificacionesPropuestas.First().Autorizado.Should().BeFalse("inicialmente debe estar desmarcado");
        }

        // =========================================================================
        // Aislamiento de Autorización por Aula y Asignatura (Paso 3)
        // =========================================================================
        [Fact]
        public void Evaluar_WhenMismaAsignaturaEnDistintasAulas_ShouldAislarAutorizacionPorAula()
        {
            // Arrange: CALI III-A (Aula 36) y CALI III-B (Aula 37) comparten la asignatura 13 con saturación en domingo (20 sesiones vs 17 domingos)
            var ctx = CrearContextoSintetico(17, "VIE-DOM", numAsignaturas: 1, sesionesPorAsignatura: 20, aulaRef: "CALI-36");
            ctx.Aulas[0].Id = 36;
            ctx.Aulas[0].Subunidad = "VALLE";
            ctx.Asignaturas[0].Id = 13;
            ctx.Asignaturas[0].Horas = 80;
            ctx.Asignaturas[0].IntHoraria = 80;
            ctx.AulaAsignaturas[0].IdAula = 36;
            ctx.AulaAsignaturas[0].IdAsignatura = 13;
            ctx.AulaAsignaturas[0].Subunidad = "VALLE";
            ctx.AulaAsignaturas[0].HorasObjetivo = 80;
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturas[0].DiasPatron = "7";

            var aula37 = new AgAula
            {
                Id = 37,
                AulaRef = "CALI-37",
                Subunidad = "VALLE",
                DiasVentana = "VIE-DOM",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                IdPrograma = 1
            };
            ctx.Aulas.Add(aula37);

            var aa37 = new AgAulaAsignatura
            {
                IdAula = 37,
                IdAsignatura = 13,
                Subunidad = "VALLE",
                HorasObjetivo = 80,
                OrigenPatron = "LEGADO",
                DiasPatron = "7"
            };
            ctx.AulaAsignaturas.Add(aa37);
            ctx.AulaAsignaturaPorAulaAsig.Clear();
            ctx.AulaAsignaturaPorAulaAsig.Add((36, 13), ctx.AulaAsignaturas[0]);
            ctx.AulaAsignaturaPorAulaAsig.Add((37, 13), aa37);
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(ctx);

            // Assert: Cada aula debe tener su propia lista acotada de patrones afectados
            decisiones.Should().HaveCount(2);
            var dec36 = decisiones.First(d => d.AulaRef == "CALI-36");
            var dec37 = decisiones.First(d => d.AulaRef == "CALI-37");

            dec36.ReclasificacionesPropuestas.Should().OnlyContain(p => p.IdAula == 36 && p.AulaRef == "CALI-36");
            dec37.ReclasificacionesPropuestas.Should().OnlyContain(p => p.IdAula == 37 && p.AulaRef == "CALI-37");

            // Autorizar únicamente en el aula 36
            dec36.ReclasificacionesPropuestas.First().Autorizado = true;

            // Verificar que no contamina el aula 37
            dec37.ReclasificacionesPropuestas.First().Autorizado.Should().BeFalse();

            var simuladas = sut.Evaluar(ctx, 0, dec36.ReclasificacionesPropuestas);
            simuladas.Single(d => d.AulaRef == "CALI-36").ReclasificacionesAutorizadas.Should().ContainSingle();
            simuladas.Single(d => d.AulaRef == "CALI-37").ReclasificacionesAutorizadas.Should().BeEmpty();
            ctx.AulaAsignaturaPorAulaAsig[(37, 13)].OrigenPatron.Should().Be("LEGADO");
        }

        // =========================================================================
        // Simulación de Reclasificación antes de confirmar (Paso 4)
        // =========================================================================
        [Fact]
        public void Simular_WhenAplicarReclasificaciones_ShouldEvaluarMismaVentanaPrimero()
        {
            // Arrange: Contexto con demanda que se resuelve en la misma ventana al redistribuir
            // Asignatura 1 con 18 sesiones (72h) en domingo excede los 17 domingos disponibles (1 sesión / 4h pendiente).
            // Con 60 sesiones totales (18 + 3*14), cabe en las 68 ranuras de VIE-DOM al reclasificar a AUTOMATICO.
            var ctx = CrearContextoSintetico(17, "VIE-DOM", numAsignaturas: 4, sesionesPorAsignatura: 14);
            ctx.Asignaturas[0].Horas = 72;
            ctx.Asignaturas[0].IntHoraria = 72;
            ctx.AulaAsignaturas[0].HorasObjetivo = 72;
            ctx.AulaAsignaturas[0].DiasPatron = "7";
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].HorasObjetivo = 72;
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "7";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(ctx);

            var reclasificaciones = decisiones[0].ReclasificacionesPropuestas;
            reclasificaciones.First().Autorizado = true;

            // Act: Simular con la sobrecarga existente de Evaluar pasando reclasificaciones autorizadas
            var decisionesSimuladas = sut.Evaluar(ctx, 0, reclasificaciones);

            // Assert: El contexto original no se muta
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron.Should().Be("LEGADO");
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron.Should().Be("7");

            // La propuesta simulada sella la huella de decisión
            decisionesSimuladas.Should().NotBeEmpty();
            decisionesSimuladas[0].HuellaDecision.Should().NotBeNullOrWhiteSpace();
        }

        // =========================================================================
        // Coherencia Horas y Fechas (Paso 5)
        // =========================================================================
        [Fact]
        public void Sellar_WhenFechasGenPresente_ShouldMantenerInvarianteCantidadFechasPorUnidadIgualHorasGen()
        {
            // Arrange: Un individuo con gen de 16h y fechas precalculadas
            var fIni = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 5, 30);
            var ctx = CrearContextoSintetico(17, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 4, fechaInicio: fIni);

            var ind = new AgIndividuo();
            var fechas = new List<DateTime>
            {
                new DateTime(2027, 2, 8),
                new DateTime(2027, 2, 15),
                new DateTime(2027, 2, 22),
                new DateTime(2027, 3, 1)
            };

            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 1,
                IdBloqueHorario = 1,
                DiaSemana = 1,
                HorasGen = 16,
                TotalHorasAsignatura = 16,
                FechasGen = new List<DateTime>(fechas),
                ClaveGrupo = "1|1|A",
                Grupo = "A",
                Estado = AgEstadoGen.Valido
            };
            ind.Genes.Add(gen);

            // Act: Ejecutar Sellar en AgCoherenciaGrupo
            AgCoherenciaGrupo.Sellar(ind, ctx);

            // Assert: Se preserva cantidadFechas * unidad == HorasGen
            gen.FechasGen.Should().NotBeNull();
            (gen.FechasGen.Count * 4.0).Should().Be(gen.HorasGen);
            gen.HorasGen.Should().Be(16.0);
        }

        // =========================================================================
        // Fase 1796 (T9): Reclasificación Exclusiva de Materias con Sesiones Pendientes
        // =========================================================================

        [Fact]
        public void Evaluar_WhenTodasLasMateriasCabenEnPlazo_ShouldRetornarCeroCandidatas()
        {
            // Arrange: 17 semanas, LUN-VIE (85 ranuras disponibles).
            // 2 materias LEGADO con 8 sesiones cada una (total 16 sesiones = 64h), caben holgadamente.
            var ctx = CrearContextoSintetico(17, "LUN-VIE", numAsignaturas: 2, sesionesPorAsignatura: 8);
            var aula = ctx.Aulas[0];

            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturas[0].DiasPatron = "1"; // Lunes (17 lunes disponibles > 8 sesiones)
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "1";

            ctx.AulaAsignaturas[1].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturas[1].DiasPatron = "3"; // Miércoles (17 miércoles disponibles > 8 sesiones)
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].DiasPatron = "3";
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var evaluacion = AgEvaluadorDemandaVentana.Evaluar(
                aula.DiasVentana, aula.Jornada, 4.0, ctx, aula, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin);
            var candidatosAula = sut.IdentificarPatronesRedistribuibles(ctx, aula);
            var candidatosGlobal = sut.IdentificarPatronesRedistribuibles(ctx);
            var decisiones = sut.Evaluar(ctx);

            // Assert
            evaluacion.CabeEnPlazo.Should().BeTrue();
            evaluacion.DemandaPendiente.Should().Be(0);
            candidatosAula.Should().BeEmpty();
            candidatosGlobal.Should().BeEmpty();
            decisiones.Should().BeEmpty("un plan donde todas las materias caben en plazo no genera propuestas ni candidatas");
        }

        [Fact]
        public void Evaluar_WhenAulaMixtaConMateriaCompletaYMateriaPendiente_ShouldOfrecerSoloMateriaPendiente()
        {
            // Arrange: Ventana de 4 semanas en LUN-VIE (20 ranuras).
            // Asignatura 1 (Lunes, 3 sesiones = 12h): cabe completa en los 4 lunes disponibles (déficit 0).
            // Asignatura 2 (Viernes, 6 sesiones = 24h): solo hay 4 viernes disponibles -> déficit de 2 sesiones (8h pendientes).
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 2, sesionesPorAsignatura: 4);

            ctx.Asignaturas[0].Horas = 12;
            ctx.Asignaturas[0].IntHoraria = 12;
            ctx.AulaAsignaturas[0].HorasObjetivo = 12;
            ctx.AulaAsignaturas[0].DiasPatron = "1";
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].HorasObjetivo = 12;
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "1";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";

            ctx.Asignaturas[1].Horas = 24;
            ctx.Asignaturas[1].IntHoraria = 24;
            ctx.AulaAsignaturas[1].HorasObjetivo = 24;
            ctx.AulaAsignaturas[1].DiasPatron = "5";
            ctx.AulaAsignaturas[1].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].HorasObjetivo = 24;
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].DiasPatron = "5";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron = "LEGADO";
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var candidatos = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]);

            // Assert
            candidatos.Should().ContainSingle();
            var candidato = candidatos[0];
            candidato.IdAsignatura.Should().Be(2, "únicamente la Asignatura 2 tiene déficit en la ventana");
            candidato.SesionesPendientes.Should().Be(2);
            candidato.HorasPendientes.Should().Be(8.0);
            candidatos.Should().NotContain(c => c.IdAsignatura == 1, "la Asignatura 1 cabe completa y no debe ser ofrecida para reclasificación");
        }

        [Fact]
        public void Evaluar_WhenLegadoPendienteVsFijoPendiente_ShouldSoloOfrecerLegado()
        {
            // Arrange: Ventana de 4 semanas en LUN-VIE.
            // Asignatura 1: LEGADO, Viernes (DiasPatron "5"), 6 sesiones (24h) -> 2 sesiones pendientes frente a 4 viernes.
            // Asignatura 2: FIJO, Lunes (DiasPatron "1"), 6 sesiones (24h) -> 2 sesiones pendientes frente a 4 lunes.
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 2, sesionesPorAsignatura: 4);

            ctx.Asignaturas[0].Horas = 24;
            ctx.Asignaturas[0].IntHoraria = 24;
            ctx.AulaAsignaturas[0].HorasObjetivo = 24;
            ctx.AulaAsignaturas[0].DiasPatron = "5";
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].HorasObjetivo = 24;
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "5";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";

            ctx.Asignaturas[1].Horas = 24;
            ctx.Asignaturas[1].IntHoraria = 24;
            ctx.AulaAsignaturas[1].HorasObjetivo = 24;
            ctx.AulaAsignaturas[1].DiasPatron = "1";
            ctx.AulaAsignaturas[1].OrigenPatron = "FIJO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].HorasObjetivo = 24;
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].DiasPatron = "1";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron = "FIJO";
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var candidatos = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]);

            // Assert
            candidatos.Should().ContainSingle();
            var candidato = candidatos[0];
            candidato.IdAsignatura.Should().Be(1, "solo la asignatura con origen LEGADO debe ser ofrecida");
            candidato.OrigenAnterior.Should().Be("LEGADO");
            candidato.OrigenPropuesto.Should().Be("AUTOMATICO");
            candidatos.Should().NotContain(c => c.IdAsignatura == 2, "las asignaturas FIJO son inmutables y nunca deben ser ofrecidas como candidatas");
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron.Should().Be("FIJO", "el patrón FIJO debe permanecer inalterado en el contexto");
        }

        [Fact]
        public void Evaluar_WhenReclasificacionResuelveEnMismaVentana_ShouldNoEscalarDias()
        {
            // Arrange: Ventana de 4 semanas en LUN-VIE (20 ranuras disponibles).
            // Asignatura 1: LEGADO en Viernes (DiasPatron "5"), 6 sesiones (24h).
            // En 4 semanas solo hay 4 viernes (2 sesiones pendientes).
            // Sin embargo, en la ventana completa LUN-VIE hay 20 ranuras disponibles (sobran 14 ranuras para 6 sesiones).
            var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 6);
            ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturas[0].DiasPatron = "5";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "5";
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Evaluación previa sin autorizar: identifica candidata con déficit
            var decisionesIniciales = sut.Evaluar(ctx);
            decisionesIniciales.Should().ContainSingle();
            var dInicial = decisionesIniciales[0];
            dInicial.ReclasificacionesPropuestas.Should().ContainSingle();
            dInicial.ReclasificacionesPropuestas[0].SesionesPendientes.Should().Be(2);

            // Autorizar reclasificación de la materia pendiente
            var autorizada = dInicial.ReclasificacionesPropuestas[0].Clone();
            autorizada.Autorizado = true;

            // Act: Evaluar con la reclasificación autorizada
            var decisiones = sut.Evaluar(ctx, 0, new[] { autorizada });

            // Assert: Resuelve en la misma ventana sin ampliar días
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("LUN-VIE");
            d.Hacia.Should().Be("LUN-VIE", "al reclasificar la materia se resuelve el déficit en la misma ventana sin necesidad de escalar días");
            d.DemandaPendiente.Should().Be(0, "todas las 6 sesiones caben holgadamente en las 20 ranuras de la ventana");
            d.DeficitHoras.Should().Be(0.0);
            d.Aplica.Should().BeTrue();
            d.ReclasificacionesAutorizadas.Should().ContainSingle();
            d.ReclasificacionesAutorizadas[0].IdAsignatura.Should().Be(1);
            d.ReclasificacionesAutorizadas[0].OrigenPropuesto.Should().Be("AUTOMATICO");
        }

        [Fact]
        public void Evaluar_WhenAutomaticoPendiente_ShouldProponerEscaleraSinReclasificaciones()
        {
            // Arrange: 17 semanas en LUN-JUE (68 ranuras disponibles). Demanda = 5 materias AUTOMATICO * 16 sesiones = 80 sesiones (12 pendientes)
            var ctx = CrearContextoSintetico(17, "LUN-JUE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var candidatos = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]);
            var decisiones = sut.Evaluar(ctx);

            // Assert: Al ser AUTOMATICO no se ofrecen candidatos de reclasificación pero sí avanza la escalera
            candidatos.Should().BeEmpty("las materias AUTOMATICO no requieren ni admiten reclasificación");
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("LUN-JUE");
            d.Hacia.Should().Be("LUN-VIE");
            d.ReclasificacionesPropuestas.Should().BeEmpty();
            d.ReclasificacionesAutorizadas.Should().BeEmpty();
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public async Task Aplicar_WhenSegundaAplicacion_ShouldSerIdempotenteYSinCambios()
        {
            // Arrange: Base SQLite temporal con aula y asignatura LEGADO
            string dbPath = Path.Combine(Path.GetTempPath(), "AgAceptacion_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbPath);
            string connectionString = $"Data Source={dbPath};Version=3;";
            Func<IDbConnection> connectionFactory = () => new SQLiteConnection(connectionString);

            try
            {
                using (var conn = connectionFactory())
                {
                    conn.Open();
                    conn.Execute(@"
CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL,
    jornada TEXT NOT NULL,
    horas_dia INTEGER NOT NULL,
    dias_ventana TEXT,
    subunidad TEXT DEFAULT '',
    id_programa INTEGER DEFAULT 1,
    capacidad INTEGER DEFAULT 35,
    estado INTEGER DEFAULT 1
);
CREATE TABLE ag_aula_asignatura (
    id_aula INTEGER NOT NULL,
    id_asignatura INTEGER NOT NULL,
    horas_objetivo REAL NOT NULL,
    orden_curricular INTEGER NOT NULL,
    dias_patron TEXT,
    origen_patron TEXT,
    PRIMARY KEY (id_aula, id_asignatura)
);
INSERT INTO ag_aula (id, aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa, estado)
VALUES (1, 'A101', 'MANANA', 4, 'LUN-VIE', '', 1, 1);
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (1, 1, 24, 1, '5', 'LEGADO');");
                }

                // Contexto idéntico al estado persistido
                var ctx = CrearContextoSintetico(4, "LUN-VIE", numAsignaturas: 1, sesionesPorAsignatura: 6, aulaRef: "A101");
                ctx.AulaAsignaturas[0].OrigenPatron = "LEGADO";
                ctx.AulaAsignaturas[0].DiasPatron = "5";
                ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";
                ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "5";
                ctx.ConstruirIndices();

                var sut = new AgEscaladorRanuraAula(connectionFactory);

                // Identificar candidata y autorizar la reclasificación
                var candidatos = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]);
                candidatos.Should().ContainSingle();
                candidatos[0].Autorizado = true;

                var decisiones = sut.Evaluar(ctx, 0, candidatos);
                decisiones.Should().ContainSingle();
                decisiones[0].Aplica.Should().BeTrue();

                // Act 1: Primera aplicación -> Persiste los cambios en ag_aula_asignatura
                var res1 = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);

                // Assert 1: Primera aplicación confirma y modifica 1 registro
                res1.Exito.Should().BeTrue(res1.Error);
                res1.Confirmadas.Should().Be(1);
                res1.Modificadas.Should().Be(1);

                // Act 2: Segunda aplicación con el mismo lote de decisiones
                var res2 = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);

                // Assert 2: Idempotencia estricta: confirmada sin nuevas escrituras
                res2.Exito.Should().BeTrue(res2.Error);
                res2.Confirmadas.Should().Be(1);
                res2.Modificadas.Should().Be(0, "la segunda aplicación debe ser idempotente y no ejecutar escrituras redundantes");

                // Act 3: Re-evaluación en el contexto actualizado ya no tiene déficit ni candidatas -> SinCambios
                ctx.AulaAsignaturas[0].OrigenPatron = "AUTOMATICO";
                ctx.AulaAsignaturas[0].DiasPatron = string.Empty;
                ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "AUTOMATICO";
                ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = string.Empty;
                ctx.ConstruirIndices();

                var decisionesSinCambios = sut.Evaluar(ctx, 0);
                var res3 = await sut.AplicarDetalladoAsync(decisionesSinCambios, 0, CancellationToken.None);

                // Assert 3: Sin propuestas aplicables resulta en SinCambios (0 confirmadas, 0 modificadas)
                decisionesSinCambios.Should().BeEmpty();
                res3.Exito.Should().BeTrue();
                res3.Confirmadas.Should().Be(0);
                res3.Modificadas.Should().Be(0);
            }
            finally
            {
                SQLiteConnection.ClearAllPools();
                try
                {
                    if (File.Exists(dbPath))
                    {
                        File.Delete(dbPath);
                    }
                }
                catch
                {
                    // Limpieza defensiva en temp
                }
            }
        }

        private static AgContexto CrearContextoSintetico(
            int w,
            string diasVentana,
            int numAsignaturas,
            int sesionesPorAsignatura,
            int horasDia = 4,
            DateTime? fechaInicio = null,
            string aulaRef = "A101")
        {
            var fInicio = fechaInicio ?? new DateTime(2027, 2, 1);
            var fFin = fInicio.AddDays(7 * w - 1);

            var contexto = new AgContexto();
            contexto.IdCorrida = 1;
            contexto.IdSede = 1;
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Aceptación",
                Modalidad = "PRESENCIAL"
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = aulaRef,
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = horasDia,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 35,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I"
            };
            contexto.Aulas.Add(aula);

            for (int d = 1; d <= 7; d++)
            {
                if (d == 6)
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 61,
                        DiaSemanaNumero = 6,
                        Franja = "MANANA",
                        HoraInicio = "08:00",
                        HoraFin = "12:00",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 62,
                        DiaSemanaNumero = 6,
                        Franja = "TARDE",
                        HoraInicio = "13:00",
                        HoraFin = "17:00",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                }
                else
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = d,
                        DiaSemanaNumero = d,
                        Franja = "MANANA",
                        HoraInicio = "08:00",
                        HoraFin = "12:00",
                        DuracionHoras = horasDia,
                        Activo = 1
                    });
                }
            }

            for (int i = 1; i <= numAsignaturas; i++)
            {
                int horas = sesionesPorAsignatura * horasDia;
                var asig = new Asignatura
                {
                    Id = i,
                    Nombre = $"Materia {i}",
                    IdPrograma = 1,
                    Semestre = "I",
                    Horas = horas,
                    IntHoraria = horas
                };
                contexto.Asignaturas.Add(asig);

                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = horas,
                    OrdenCurricular = i,
                    DiasPatron = string.Empty,
                    OrigenPatron = "AUTOMATICO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig.Add((1, i), aa);
            }

            contexto.Docentes.Add(new Tercero { Id = 10, Nombre = "Docente 10", Residencia = "CALI" });
            contexto.Sedes.Add(new Sede { Id = 1, Ciudad = "CALI" });
            contexto.Presenciales.Add((fInicio, fFin));

            contexto.ConstruirIndices();
            return contexto;
        }
    }
}
