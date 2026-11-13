using System;
using System.Collections.Generic;
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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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

            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
            var sut = new AgEscaladorVentanaAula(() => null);

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
    }
}

