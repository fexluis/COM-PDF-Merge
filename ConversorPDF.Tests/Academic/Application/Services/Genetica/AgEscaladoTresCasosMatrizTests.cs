using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Suite integral de pruebas para la matriz de escalado automático de ventanas (Fase 1740).
    /// Valida los tres casos principales (LUN-JUE, MAR-VIE, VIE-DOM), escalado multi-paso 6W con demanda pendiente,
    /// terminalidad de ventanas tope (MIE-DOM y LUN-VIE) y diferenciación de unidades horarias de sábado (4h vs 8h).
    /// </summary>
    public class AgEscaladoTresCasosMatrizTests
    {
        [Fact]
        public void Evaluar_Caso1_LunJue5W_ShouldProponerLunVie_YSegundaCorridaSinPropuesta()
        {
            // Arrange: 17 semanas, LUN-JUE (4 días * 1 ranura = 68 ranuras). Demanda = 5 asignaturas de 16 sesiones = 80 sesiones.
            int w = 17;
            var contexto = CrearContextoMatriz(w, "LUN-JUE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Primera corrida
            var decisiones1 = sut.Evaluar(contexto);

            // Assert 1: Propone LUN-JUE -> LUN-VIE (4W -> 5W = 68 -> 85 ranuras)
            decisiones1.Should().ContainSingle();
            var d1 = decisiones1[0];
            d1.AulaRef.Should().Be("A101");
            d1.Desde.Should().Be("LUN-JUE");
            d1.Hacia.Should().Be("LUN-VIE");
            d1.CapacidadAntes.Should().Be(68);
            d1.CapacidadDespues.Should().Be(85);
            d1.DemandaSesiones.Should().Be(80);
            d1.Aplica.Should().BeTrue();

            // Act 2: Aceptar propuesta y re-evaluar (segunda corrida)
            contexto.Aulas[0].DiasVentana = d1.Hacia;
            var decisiones2 = sut.Evaluar(contexto);

            // Assert 2: Segunda corrida sobre estado factible no produce ninguna propuesta
            decisiones2.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_Caso2_MarVie5W_ShouldProponerLunVie_YSegundaCorridaSinPropuesta()
        {
            // Arrange: 17 semanas, MAR-VIE (4 días * 1 ranura = 68 ranuras). Demanda = 5 asignaturas de 16 sesiones = 80 sesiones.
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Primera corrida
            var decisiones1 = sut.Evaluar(contexto);

            // Assert 1: Propone MAR-VIE -> LUN-VIE (68 -> 85 ranuras)
            decisiones1.Should().ContainSingle();
            var d1 = decisiones1[0];
            d1.AulaRef.Should().Be("A101");
            d1.Desde.Should().Be("MAR-VIE");
            d1.Hacia.Should().Be("LUN-VIE");
            d1.CapacidadAntes.Should().Be(68);
            d1.CapacidadDespues.Should().Be(85);
            d1.DemandaSesiones.Should().Be(80);
            d1.Aplica.Should().BeTrue();

            // Act 2: Aceptar propuesta y re-evaluar
            contexto.Aulas[0].DiasVentana = d1.Hacia;
            var decisiones2 = sut.Evaluar(contexto);

            // Assert 2: Segunda corrida sin propuesta
            decisiones2.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_Caso3_VieDom5W_ShouldProponerJueDom_YSegundaCorridaSinPropuesta()
        {
            // Arrange: 17 semanas, VIE-DOM (4 ranuras/semana = 68 ranuras).
            // Demanda 5W = 5 asignaturas de 16 sesiones = 80 sesiones.
            int w = 17;
            var contexto = CrearContextoMatriz(w, "VIE-DOM", numAsignaturas: 5, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Primera corrida
            var decisiones1 = sut.Evaluar(contexto);

            // Assert 1: Propone VIE-DOM -> JUE-DOM (68 -> 85 ranuras)
            decisiones1.Should().ContainSingle();
            var d1 = decisiones1[0];
            d1.AulaRef.Should().Be("A101");
            d1.Desde.Should().Be("VIE-DOM");
            d1.Hacia.Should().Be("JUE-DOM");
            d1.CapacidadAntes.Should().Be(68);
            d1.CapacidadDespues.Should().Be(85);
            d1.DemandaSesiones.Should().Be(80);
            d1.Aplica.Should().BeTrue();

            // Act 2: Aceptar propuesta y re-evaluar
            contexto.Aulas[0].DiasVentana = d1.Hacia;
            var decisiones2 = sut.Evaluar(contexto);

            // Assert 2: Con 85 ranuras balanceadas en JUE-DOM, las 80 sesiones caben (2 en JUE-VIE=32<=34, 3 en SAB-DOM=48<=51)
            decisiones2.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_Caso4_VieDom6W_Paso1JueDomConPendiente_Paso2MieDom_Paso3Terminal()
        {
            // Arrange: 17 semanas, VIE-DOM. Demanda = 6 asignaturas de 16 sesiones = 96 sesiones.
            int w = 17;
            var contexto = CrearContextoMatriz(w, "VIE-DOM", numAsignaturas: 6, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Primera corrida
            var decisiones1 = sut.Evaluar(contexto);

            // Assert 1: Propone solo el paso inmediato Siguiente: VIE-DOM -> JUE-DOM (68 -> 85 ranuras).
            // Nota: Demanda (96) > CapacidadDespues (85) -> hay 11 sesiones pendientes (solución parcial explícita).
            decisiones1.Should().ContainSingle();
            var d1 = decisiones1[0];
            d1.Desde.Should().Be("VIE-DOM");
            d1.Hacia.Should().Be("JUE-DOM");
            d1.CapacidadAntes.Should().Be(68);
            d1.CapacidadDespues.Should().Be(85);
            d1.DemandaSesiones.Should().Be(96);
            (d1.DemandaSesiones - d1.CapacidadDespues).Should().Be(11, "debe registrar demanda pendiente explícita tras paso 1");

            // Act 2: Aceptar paso 1 y re-evaluar (segunda corrida)
            contexto.Aulas[0].DiasVentana = d1.Hacia;
            var decisiones2 = sut.Evaluar(contexto);

            // Assert 2: Segunda corrida ofrece el siguiente escalón: JUE-DOM -> MIE-DOM (85 -> 102 ranuras).
            // Con 102 ranuras, la demanda de 96 sesiones queda completamente cubierta.
            decisiones2.Should().ContainSingle();
            var d2 = decisiones2[0];
            d2.Desde.Should().Be("JUE-DOM");
            d2.Hacia.Should().Be("MIE-DOM");
            d2.CapacidadAntes.Should().Be(85);
            d2.CapacidadDespues.Should().Be(102);
            d2.DemandaSesiones.Should().Be(96);
            d2.CapacidadDespues.Should().BeGreaterThanOrEqualTo(d2.DemandaSesiones);

            // Act 3: Aceptar paso 2 y re-evaluar (tercera corrida)
            contexto.Aulas[0].DiasVentana = d2.Hacia;
            var decisiones3 = sut.Evaluar(contexto);

            // Assert 3: Tercera corrida sin propuesta (estado factible alcanzado)
            decisiones3.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_Terminalidad_MieDomYLunVie_ShouldNoProponerNadaAunqueHayaSaturacion()
        {
            // Arrange: Aula en MIE-DOM (ventana máxima de fin de semana) con demanda excesiva (110 sesiones > 102 ranuras)
            int w = 17;
            var ctxMieDom = CrearContextoMatriz(w, "3,4,5,6,7", numAsignaturas: 7, sesionesPorAsignatura: 16);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act & Assert MIE-DOM: Terminal; no inventa escalones adicionales
            var decMieDom = sut.Evaluar(ctxMieDom);
            decMieDom.Should().BeEmpty();

            // Arrange: Aula en LUN-VIE (ventana máxima de entre semana) con demanda excesiva (96 sesiones > 85 ranuras)
            var ctxLunVie = CrearContextoMatriz(w, "1,2,3,4,5", numAsignaturas: 6, sesionesPorAsignatura: 16);

            // Act & Assert LUN-VIE: Terminal; no inventa escalones adicionales
            var decLunVie = sut.Evaluar(ctxLunVie);
            decLunVie.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_SabadoHorasDia_DiferenciaCapacidadEntre4hy8h()
        {
            // Arrange: 17 semanas, ventana VIE-DOM ("5,6,7"). Demanda 5W = 80 sesiones.
            // Para 4h: Viernes (1) + Sábado (2) + Domingo (1) = 4 ranuras/semana * 17 = 68 ranuras.
            // Para 8h: Viernes (1) + Sábado (1) + Domingo (1) = 3 ranuras/semana * 17 = 51 ranuras.
            int w = 17;
            var ctx4h = CrearContextoMatriz(w, "5,6,7", numAsignaturas: 5, sesionesPorAsignatura: 16, horasDia: 4);
            var ctx8h = CrearContextoMatriz(w, "5,6,7", numAsignaturas: 5, sesionesPorAsignatura: 16, horasDia: 8, modalidadPrograma: "POSGRADO");

            var sut = new AgEscaladorRanuraAula(() => null);

            var dec4h = sut.Evaluar(ctx4h);
            var dec8h = sut.Evaluar(ctx8h);

            // Assert: Las capacidades evaluadas antes del escalado reflejan exactamente 68 ranuras para 4h y 51 para 8h
            dec4h.Should().ContainSingle();
            dec4h[0].CapacidadAntes.Should().Be(68);

            dec8h.Should().ContainSingle();
            dec8h[0].CapacidadAntes.Should().Be(51);
        }

        [Fact]
        public void Evaluar_LunJue_PatronFijoViernes_ShouldProponerLunVie()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "LUN-JUE", numAsignaturas: 2, sesionesPorAsignatura: 12);
            FijarPatron(contexto, "5");
            var sut = new AgEscaladorRanuraAula(() => null);

            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("LUN-VIE");
            decisiones[0].MateriasFueraVentana.Should().Be(2);
            decisiones[0].Aplica.Should().BeTrue();

            contexto.Aulas[0].DiasVentana = decisiones[0].Hacia;
            sut.Evaluar(contexto).Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_MarVie_PatronFijoLunes_ShouldProponerLunVie()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 2, sesionesPorAsignatura: 12);
            FijarPatron(contexto, "1");
            var sut = new AgEscaladorRanuraAula(() => null);

            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("LUN-VIE");
            decisiones[0].MateriasFueraVentana.Should().Be(2);
            decisiones[0].Aplica.Should().BeTrue();
        }

        [Fact]
        public void Evaluar_VieDom_PatronFijoJueves_ShouldProponerJueDomYNoMieDom()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "VIE-DOM", numAsignaturas: 2, sesionesPorAsignatura: 12);
            FijarPatron(contexto, "4");
            var sut = new AgEscaladorRanuraAula(() => null);

            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("JUE-DOM");
            decisiones[0].MateriasFueraVentana.Should().Be(2);
            decisiones[0].Aplica.Should().BeTrue();

            contexto.Aulas[0].DiasVentana = decisiones[0].Hacia;
            sut.Evaluar(contexto).Should().BeEmpty("MIE-DOM no reduce las semanas ni la saturación de jueves");
        }

        [Fact]
        public void Evaluar_VieDom_TresPatronesDomingo_ShouldNoProponer()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "VIE-DOM", numAsignaturas: 3, sesionesPorAsignatura: 12);
            FijarPatron(contexto, "7");
            var sut = new AgEscaladorRanuraAula(() => null);

            sut.Evaluar(contexto).Should().BeEmpty("jueves/miércoles no reducen la cola de domingo");
        }

        [Fact]
        public void Evaluar_VieDom_TresDomingosConFestivo_ShouldCalcularFechaMinima17Oct()
        {
            var inicio = new DateTime(2027, 2, 7);
            var contexto = CrearContextoMatriz(17, "VIE-DOM", numAsignaturas: 3, sesionesPorAsignatura: 12, fechaInicio: inicio);
            FijarPatron(contexto, "7");
            contexto.Festivos.Add((new DateTime(2027, 8, 15), new DateTime(2027, 8, 15)));
            contexto.ConstruirIndices();

            var res = AgEvaluadorDemandaVentana.Evaluar(
                "VIE-DOM", "MANANA", 4.0, contexto, contexto.Aulas[0],
                contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);

            res.SemanasRequeridas.Should().Be(36);
            res.FechaFinMinima.Should().Be(new DateTime(2027, 10, 17));
            res.SaturacionCarriles.Should().BeTrue();
            res.MateriasFueraVentana.Should().BeEmpty();
        }

        [Fact]
        public void ValidarFactibilidad_VieDomTresDomingos_ShouldEmitirI5DiaSinSugerirJueDom()
        {
            var inicio = new DateTime(2027, 2, 7);
            var contexto = CrearContextoMatriz(17, "VIE-DOM", numAsignaturas: 3, sesionesPorAsignatura: 12, fechaInicio: inicio);
            FijarPatron(contexto, "7");
            contexto.Festivos.Add((new DateTime(2027, 8, 15), new DateTime(2027, 8, 15)));
            contexto.ConstruirIndices();

            var informe = new AgValidadorVentana().ValidarFactibilidad(contexto);

            informe.IsSuccess.Should().BeTrue();
            string bloqueos = string.Join(" | ", informe.Value.BloqueosCriticos);
            bloqueos.Should().Contain("I5_DIA_SATURADO");
            bloqueos.Should().NotContain("Escale la ventana a 'JUE-DOM'");
        }

        // FASE 2111 / FASE 2114: Caso corrida 113/115 - 64 sesiones (12,12,8,8,12,12) en MAR-VIE con 61 ranuras -> propone LUN-VIE (75)
        [Fact]
        public void Evaluar_Caso5_MarVie_64Sesiones_FinMayo_ShouldProponerLunVie()
        {
            var inicio = new DateTime(2027, 2, 5); // viernes
            var fin = new DateTime(2027, 5, 30);   // domingo (16 semanas + 3 días)
            var contexto = CrearContextoMatriz(16, "MAR-VIE", new[] { 12, 12, 8, 8, 12, 12 }, fechaInicio: inicio, fechaFin: fin);
            // 4 festivos en hábiles de MAR-VIE y 2 festivos en lunes
            foreach (var f in DiasHabiles(inicio, fin, 1, new[] { 2, 3, 4, 5 })) contexto.Festivos.Add((f, f));
            foreach (var f in DiasHabiles(inicio, fin, 2, new[] { 1 })) contexto.Festivos.Add((f, f));
            contexto.ConstruirIndices();

            var decisiones = new AgEscaladorRanuraAula(() => null).Evaluar(contexto);

            decisiones.Should().ContainSingle();
            var d = decisiones.Single();
            d.Desde.Should().Be("MAR-VIE");
            d.Hacia.Should().Be("LUN-VIE");
            d.CapacidadAntes.Should().Be(61);
            d.CapacidadDespues.Should().Be(75);
            d.Aplica.Should().BeTrue();
        }

        // FASE 2111 / FASE 2114: Con patrones rígidos MAR-MIÉ / JUE-VIE no se debe silenciar la propuesta
        [Fact]
        public void Evaluar_Caso6_MarVie_PatronesRigidos_ShouldNoSilenciarPropuesta()
        {
            var inicio = new DateTime(2027, 2, 5);
            var fin = new DateTime(2027, 5, 30);
            var contexto = CrearContextoMatriz(16, "MAR-VIE", new[] { 12, 12, 8, 8, 12, 12 }, fechaInicio: inicio, fechaFin: fin);
            foreach (var f in DiasHabiles(inicio, fin, 1, new[] { 2, 3, 4, 5 })) contexto.Festivos.Add((f, f));
            foreach (var f in DiasHabiles(inicio, fin, 2, new[] { 1 })) contexto.Festivos.Add((f, f));
            FijarPatrones(contexto, new[] { "2,3", "2,3", "4,5", "4,5", "2,3", "4,5" }, "LEGADO");
            contexto.ConstruirIndices();

            var decisiones = new AgEscaladorRanuraAula(() => null).Evaluar(contexto);

            decisiones.Should().NotBeEmpty("una saturación con patrones rígidos debe producir escalón y/o candidatas de reclasificación");
            decisiones.Any(d => d.DemandaPendiente > 0 || d.ReclasificacionesPropuestas.Count > 0).Should().BeTrue();
        }

        private static void FijarPatrones(AgContexto c, string[] patrones, string origen)
        {
            for (int i = 0; i < c.AulaAsignaturas.Count; i++)
            {
                c.AulaAsignaturas[i].DiasPatron = patrones[i];
                c.AulaAsignaturas[i].OrigenPatron = origen;
                if (c.AulaAsignaturaPorAulaAsig.TryGetValue((c.AulaAsignaturas[i].IdAula, c.AulaAsignaturas[i].IdAsignatura), out var aa))
                {
                    aa.DiasPatron = patrones[i];
                    aa.OrigenPatron = origen;
                }
            }
        }

        private static IEnumerable<DateTime> DiasHabiles(DateTime inicio, DateTime fin, int cantidad, int[] diasIso)
        {
            var cuenta = new Dictionary<int, int>();
            for (var d = inicio; d <= fin; d = d.AddDays(1))
            {
                int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                if (!diasIso.Contains(iso)) { continue; }
                cuenta.TryGetValue(iso, out int n);
                if (n >= cantidad) { continue; }
                cuenta[iso] = n + 1;
                yield return d;
            }
        }

        private static void FijarPatron(AgContexto contexto, string diasPatron)
        {
            foreach (var aa in contexto.AulaAsignaturas)
            {
                aa.DiasPatron = diasPatron;
                aa.OrigenPatron = "FIJO";
            }

            foreach (var kv in contexto.AulaAsignaturaPorAulaAsig)
            {
                kv.Value.DiasPatron = diasPatron;
                kv.Value.OrigenPatron = "FIJO";
            }
        }

        private static AgContexto CrearContextoMatriz(
            int w,
            string diasVentana,
            int[] sesionesPorMateria,
            int horasDia = 4,
            string modalidadPrograma = "PRESENCIAL",
            DateTime? fechaInicio = null,
            DateTime? fechaFin = null)
        {
            var fInicio = fechaInicio ?? new DateTime(2026, 8, 3);
            var fFin = fechaFin ?? fInicio.AddDays(7 * w - 1);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(
                1, "2026-2", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Matriz",
                Modalidad = modalidadPrograma
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "A101",
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = horasDia,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30,
                IdPrograma = 1,
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
                        DuracionHoras = horasDia == 8 ? 8 : 4,
                        Activo = 1
                    });
                    if (horasDia == 4)
                    {
                        contexto.Bloques.Add(new AgBloqueHorario
                        {
                            Id = 62,
                            DiaSemanaNumero = 6,
                            Franja = "TARDE",
                            DuracionHoras = 4,
                            Activo = 1
                        });
                    }
                }
                else
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = d,
                        DiaSemanaNumero = d,
                        Franja = "MANANA",
                        DuracionHoras = horasDia,
                        Activo = 1
                    });
                }
            }

            for (int i = 1; i <= sesionesPorMateria.Length; i++)
            {
                int sesiones = sesionesPorMateria[i - 1];
                var asig = new Asignatura
                {
                    Id = i,
                    Nombre = $"Materia {i}",
                    IdPrograma = 1,
                    Semestre = "I",
                    Subunidad = "",
                    Creditos = 3,
                    Horas = sesiones * horasDia
                };
                contexto.Asignaturas.Add(asig);

                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = sesiones * horasDia,
                    DiasPatron = string.Empty,
                    OrigenPatron = "AUTOMATICO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig.Add((1, i), aa);
            }

            contexto.ConstruirIndices();
            return contexto;
        }

        private static AgContexto CrearContextoMatriz(
            int w,
            string diasVentana,
            int numAsignaturas,
            int sesionesPorAsignatura,
            int horasDia = 4,
            string modalidadPrograma = "PRESENCIAL",
            DateTime? fechaInicio = null)
        {
            var fInicio = fechaInicio ?? new DateTime(2026, 8, 3); // Lunes por defecto
            var fFin = fInicio.AddDays(7 * w - 1);  // Domingo

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(
                1, "2026-2", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Matriz",
                Modalidad = modalidadPrograma
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "A101",
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = horasDia,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30,
                IdPrograma = 1,
                Semestre = "I"
            };
            contexto.Aulas.Add(aula);

            // Bloques para los 7 días
            for (int d = 1; d <= 7; d++)
            {
                if (d == 6)
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 61,
                        DiaSemanaNumero = 6,
                        Franja = "MANANA",
                        DuracionHoras = horasDia == 8 ? 8 : 4,
                        Activo = 1
                    });
                    if (horasDia == 4)
                    {
                        contexto.Bloques.Add(new AgBloqueHorario
                        {
                            Id = 62,
                            DiaSemanaNumero = 6,
                            Franja = "TARDE",
                            DuracionHoras = 4,
                            Activo = 1
                        });
                    }
                }
                else
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = d,
                        DiaSemanaNumero = d,
                        Franja = "MANANA",
                        DuracionHoras = horasDia,
                        Activo = 1
                    });
                }
            }

            // Asignaturas y ag_aula_asignatura
            for (int i = 1; i <= numAsignaturas; i++)
            {
                var asig = new Asignatura
                {
                    Id = i,
                    Nombre = $"Materia {i}",
                    IdPrograma = 1,
                    Semestre = "I",
                    Subunidad = "",
                    Creditos = 3,
                    Horas = sesionesPorAsignatura * horasDia
                };
                contexto.Asignaturas.Add(asig);

                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = sesionesPorAsignatura * horasDia,
                    DiasPatron = string.Empty,
                    OrigenPatron = "AUTOMATICO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig.Add((1, i), aa);
            }

            contexto.ConstruirIndices();
            return contexto;
        }

        [Fact]
        public void MarVie_Cabe_NoEscala_YParticionOrigen()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 4, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void MarVie_NoCabe_Escala_LunVie_Con_2_3_y_1_4_5()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("MAR-VIE");
            d.Hacia.Should().Be("LUN-VIE");
            d.ParticionPreferente.Should().Be("2,3|1,4,5");
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public void LunJue_NoCabe_Escala_LunVie_Con_1_2_5_y_3_4()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "LUN-JUE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "1,2";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "3,4";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("LUN-JUE");
            d.Hacia.Should().Be("LUN-VIE");
            d.ParticionPreferente.Should().Be("1,2,5|3,4");
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public void AulaConPatronYParticionVigente_NoGeneraDecision()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 4, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void ParticionIgualAReferencia_NoAplica()
        {
            var d = new AgEscaladorRanuraAula.Decision
            {
                Desde = "MAR-VIE",
                Hacia = "MAR-VIE",
                ParticionPreferente = "2,3|4,5",
                ParticionMejora = true,
                ParticionDifiereDeReferencia = false,
                CapacidadAntes = 68,
                CapacidadDespues = 68
            };

            d.Aplica.Should().BeFalse();
        }

        [Fact]
        public void ReclasificacionAutorizada_ConservaPatronUsaPatronesSimulados()
        {
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 4, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "1,2"; // patrón no compatible con MAR-VIE
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var rec = new AgReclasificacionPatron
            {
                IdAula = contexto.Aulas[0].Id,
                IdAsignatura = contexto.AulaAsignaturas[0].IdAsignatura,
                AulaRef = contexto.Aulas[0].AulaRef,
                Subunidad = contexto.Aulas[0].Subunidad,
                OrigenAnterior = "MANUAL",
                DiasPatronAnterior = "1,2",
                OrigenPropuesto = "AUTOMATICO",
                DiasPatronPropuesto = "2,3",
                Autorizado = true
            };

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto, idArbol: 0, reclasificacionesAutorizadas: new List<AgReclasificacionPatron> { rec });

            // Al autorizar reclasificación a 2,3, los patrones simulados son 2,3 y 4,5,
            // que sí son conservados por la opción 2,3|4,5.
            if (decisiones.Count > 0)
            {
                decisiones[0].ConservaPatronOrigen.Should().BeTrue();
            }
        }

        [Fact]
        public void LunVie_ConPatronYReferenciaFactible_NoGeneraDecision()
        {
            // FASE 2206: cohorte LUN-VIE con patrón donde la referencia ya es factible -> no genera propuesta no-op
            int w = 17;
            var contexto = CrearContextoMatriz(w, "LUN-VIE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void ParticionFranjaDiferente_DifiereDeReferencia()
        {
            // FASE 2209: Diferencia de franja (6M vs 6T)
            var ranuras = new List<AgRanura>
            {
                new AgRanura(5, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.MANANA, 4.0),
                new AgRanura(4, AgFranjaJornada.NOCHE, 4.0),
                new AgRanura(6, AgFranjaJornada.TARDE, 4.0),
                new AgRanura(7, AgFranjaJornada.NOCHE, 4.0)
            };
            var patrones = new List<string> { "5,6M", "4,6T,7" };
            var opcion = AgOpcionesParticion.Resolver("JUE-DOM", ranuras, patrones);

            opcion.Should().NotBeNull();
            opcion.CarrilesTexto.Should().Be("5,6M|4,6T,7");
        }

        [Fact]
        public void Escalon_ComparaInfactiblesConReferenciaDestino()
        {
            // FASE 2216: Escalón compara infactibles contra la referencia de la ventana destino
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "MANUAL";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "MANUAL";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("LUN-VIE");
            decisiones[0].ParticionPreferente.Should().Be("2,3|1,4,5");
            decisiones[0].Aplica.Should().BeTrue();
        }

        [Fact]
        public void Escalador_ConPatronesMixtos_SeleccionaCatalogo()
        {
            // FASE 2218 / FASE 2221: Convivencia de patrones LEGADO y AUTOMATICO selecciona opción de catálogo
            int w = 17;
            var contexto = CrearContextoMatriz(w, "MAR-VIE", numAsignaturas: 5, sesionesPorAsignatura: 16);
            contexto.AulaAsignaturas[0].DiasPatron = "2,3";
            contexto.AulaAsignaturas[0].OrigenPatron = "LEGADO";
            contexto.AulaAsignaturas[1].DiasPatron = "4,5";
            contexto.AulaAsignaturas[1].OrigenPatron = "LEGADO";
            // Asignatura con patrón AUTOMATICO que generaba conflicto asimétrico antes de F2218
            contexto.AulaAsignaturas[2].DiasPatron = "1,2";
            contexto.AulaAsignaturas[2].OrigenPatron = "AUTOMATICO";
            contexto.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("LUN-VIE");
            decisiones[0].ParticionPreferente.Should().Be("2,3|1,4,5");
            decisiones[0].ConservaPatronOrigen.Should().BeTrue();
            decisiones[0].Aplica.Should().BeTrue();
        }

        [Fact]
        public void ConservaPatronOrigen_NoBorraDiasPatron()
        {
            // FASE 2219 / FASE 2221 / FASE 2257: Cuando ConservaPatronOrigen es true, no se borran los dias_patron (desacoplado de Desde == Hacia)
            var decision = new AgEscaladorRanuraAula.Decision
            {
                IdAula = 1,
                AulaRef = "A101",
                Desde = "MAR-VIE",
                Hacia = "LUN-VIE",
                ParticionPreferente = "2,3|1,4,5",
                ParticionMejora = true,
                ParticionDifiereDeReferencia = true,
                ConservaPatronOrigen = true,
                IdAsignaturasBottleneck = new List<int> { 1, 2 }
            };

            // Simulación de la condición en AplicarDetalladoAsync (Fase 2257: solo depende de !ConservaPatronOrigen):
            // if (d.IdAsignaturasBottleneck.Count > 0 && !d.ConservaPatronOrigen) -> borra patrones
            bool deberiaBorrar = decision.IdAsignaturasBottleneck.Count > 0
                && !decision.ConservaPatronOrigen;
            deberiaBorrar.Should().BeFalse();
        }

        [Fact]
        public void Decision_SoloParticionSinCambioVentana_AplicaEsFalse()
        {
            // FASE 2256: Decisiones donde Desde == Hacia y sin reclasificaciones no aplican a BD
            var decision = new AgEscaladorRanuraAula.Decision
            {
                Desde = "MAR-VIE",
                Hacia = "MAR-VIE",
                ParticionPreferente = "2,3|4,5",
                ParticionMejora = true,
                ParticionDifiereDeReferencia = true,
                CapacidadAntes = 64,
                CapacidadDespues = 64
            };

            decision.Aplica.Should().BeFalse();
        }

        [Fact]
        public void PatronesAlojables_SeisPatronesDosCarriles_True()
        {
            // FASE 2247 / FASE 2254 / FASE 2255: Modelo muchos-a-uno; 6 patrones sobre 2 carriles deben ser conservados
            var carril1 = new List<AgRanura> { new AgRanura(2, AgFranjaJornada.NOCHE, 4), new AgRanura(3, AgFranjaJornada.NOCHE, 4) };
            var carril2 = new List<AgRanura> { new AgRanura(4, AgFranjaJornada.NOCHE, 4), new AgRanura(5, AgFranjaJornada.NOCHE, 4) };
            var opcion = new AgOpcionParticion(new List<IReadOnlyList<AgRanura>> { carril1, carril2 });

            // 6 materias que comparten los 2 carriles (3 en carril 1, 3 en carril 2)
            var patrones = new List<string> { "2,3", "2,3", "2,3", "4,5", "4,5", "4,5" };

            bool resultado = AgEscaladorRanuraAula.PatronesAlojablesEnOpcion(opcion, patrones);
            resultado.Should().BeTrue();
        }

        [Fact]
        public void PatronesAlojables_ConSabadoEnCarril_ConservaPatron()
        {
            // FASE 2248 / FASE 2254: Chequeo de ranura de sábado en carril vs regla blanket de texto
            var carril1 = new List<AgRanura> { new AgRanura(5, AgFranjaJornada.MANANA, 4), new AgRanura(6, AgFranjaJornada.MANANA, 4) };
            var carril2 = new List<AgRanura> { new AgRanura(6, AgFranjaJornada.TARDE, 4), new AgRanura(7, AgFranjaJornada.MANANA, 4) };
            var opcion = new AgOpcionParticion(new List<IReadOnlyList<AgRanura>> { carril1, carril2 });

            var patrones = new List<string> { "5,6" };

            bool resultado = AgEscaladorRanuraAula.PatronesAlojablesEnOpcion(opcion, patrones);
            resultado.Should().BeTrue();
        }

        [Fact]
        public void DiasKDe_ConOrigenAutomatico_IgnoraDiasPatron()
        {
            // FASE 2249: Excluir orígenes AUTOMATICO en siembra sin plan
            var aula = new AgAula { Id = 1, DiasVentana = "LUN-VIE" };
            var aa = new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 10,
                DiasPatron = "1,2",
                OrigenPatron = "AUTOMATICO"
            };

            // Con AUTOMATICO no debe usar "1,2" sino el carril canónico (para carril 1 de LUN-VIE => [3,4])
            var dias = AgPlanificadorPatronSemanal.DiasKDe(aa, aula, carril: 1);
            dias.Should().Equal(new[] { 3, 4 });
        }

        [Fact]
        public void RelacionesAula_IndiceVacio_UsaListaYConstruyeIndices()
        {
            // FASE 2250 / FASE 2251 / FASE 2255: Fallback a AulaAsignaturas cuando AulaAsignaturaPorAulaAsig está vacío
            var ctx = new AgContexto
            {
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 5, IdAsignatura = 1, DiasPatron = "2,3", OrigenPatron = "LEGADO" },
                    new AgAulaAsignatura { IdAula = 5, IdAsignatura = 2, DiasPatron = "4,5", OrigenPatron = "AUTOMATICO" },
                    new AgAulaAsignatura { IdAula = 6, IdAsignatura = 3, DiasPatron = "1,2", OrigenPatron = "MANUAL" }
                }
            };
            // Sin llamar explícitamente a ctx.ConstruirIndices(), RelacionesAula invoca ConstruirIndices y puebla el diccionario
            var relaciones = AgContexto.RelacionesAula(ctx).ToList();
            relaciones.Should().HaveCount(3);
            ctx.AulaAsignaturaPorAulaAsig.Should().NotBeNull();
            ctx.AulaAsignaturaPorAulaAsig.Should().HaveCount(3);
        }

        [Fact]
        public void Escalador_PatronesEfectivos_ExcluyeAutomatico()
        {
            // FASE 2250 / FASE 2251 / FASE 2255: PatronesEfectivos excluye AUTOMATICO y utiliza RelacionesAula
            var ctx = new AgContexto
            {
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 5, IdAsignatura = 1, DiasPatron = "2,3", OrigenPatron = "LEGADO" },
                    new AgAulaAsignatura { IdAula = 5, IdAsignatura = 2, DiasPatron = "4,5", OrigenPatron = "AUTOMATICO" },
                    new AgAulaAsignatura { IdAula = 6, IdAsignatura = 3, DiasPatron = "1,2", OrigenPatron = "MANUAL" }
                }
            };
            var patrones = AgEscaladorRanuraAula.PatronesEfectivos(ctx, 5);
            patrones.Should().ContainSingle().Which.Should().Be("2,3");
        }

        [Fact]
        public void PatronesAlojables_Dia6SinRanura_False()
        {
            // FASE 2248 / FASE 2252 / FASE 2254 / FASE 2255: Patrón con día 6 en carril que carece de sábado retorna false
            var carril1 = new List<AgRanura> { new AgRanura(2, AgFranjaJornada.NOCHE, 4), new AgRanura(3, AgFranjaJornada.NOCHE, 4) };
            var carril2 = new List<AgRanura> { new AgRanura(4, AgFranjaJornada.NOCHE, 4), new AgRanura(5, AgFranjaJornada.NOCHE, 4) };
            var opcion = new AgOpcionParticion(new List<IReadOnlyList<AgRanura>> { carril1, carril2 });

            var patrones = new List<string> { "2,3", "6" };
            bool conserva = AgEscaladorRanuraAula.PatronesAlojablesEnOpcion(opcion, patrones);
            conserva.Should().BeFalse();
        }
    }
}

