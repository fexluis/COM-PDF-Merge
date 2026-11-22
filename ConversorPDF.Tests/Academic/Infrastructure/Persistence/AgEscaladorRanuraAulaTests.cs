using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgEscaladorRanuraAulaTests
    {
        [Fact]
        public void DecisionAplica_CuandoCapacidadCrece_DevuelveTrue()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "VIE-SAB",
                Hacia = "JUE-SAB",
                DemandaSesiones = 10,
                CapacidadAntes = 6,
                CapacidadDespues = 12,
                Motivo = "demanda > capacidad"
            };

            decision.Aplica.Should().BeTrue();
        }

        [Fact]
        public void DecisionAplica_CuandoCapacidadNoCrece_DevuelveFalse()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "VIE-SAB",
                Hacia = "VIE-SAB",
                DemandaSesiones = 10,
                CapacidadAntes = 6,
                CapacidadDespues = 6,
                Motivo = "sin cambio"
            };

            decision.Aplica.Should().BeFalse();
        }

        [Fact]
        public void DecisionAplica_CuandoHaciaEsVacio_DevuelveFalse()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "4,5,6,7",
                Hacia = "",
                DemandaSesiones = 20,
                CapacidadAntes = 16,
                CapacidadDespues = 16,
                Motivo = "tope de escalera"
            };

            decision.Aplica.Should().BeFalse();
        }

        [Fact]
        public void Evaluar_WhenAulaLunJueConDemanda5W_ShouldProponerEscalarALunVieConCapacidad4Wa5W()
        {
            // Arrange (W = 4 semanas completas: 16 ranuras antes, 20 después)
            int w = 4;
            var contexto = CrearContextoNominal(w, "LUN-JUE", 5 * w);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.AulaRef.Should().Be("A101");
            d.Desde.Should().Be("LUN-JUE");
            d.Hacia.Should().Be("LUN-VIE");
            d.CapacidadAntes.Should().Be(4 * w);
            d.CapacidadDespues.Should().Be(5 * w);
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public void Evaluar_WhenAulaMarVieConDemanda5W_ShouldProponerEscalarALunVieConCapacidad4Wa5W()
        {
            // Arrange
            int w = 4;
            var contexto = CrearContextoNominal(w, "MAR-VIE", 5 * w);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("MAR-VIE");
            d.Hacia.Should().Be("LUN-VIE");
            d.CapacidadAntes.Should().Be(4 * w);
            d.CapacidadDespues.Should().Be(5 * w);
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public void Evaluar_WhenAulaVieDomConDemanda5W_ShouldProponerEscalarAJueDomConCapacidad4Wa5W()
        {
            // Arrange
            int w = 4;
            var contexto = CrearContextoNominal(w, "VIE-DOM", 5 * w);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert
            decisiones.Should().ContainSingle();
            var d = decisiones[0];
            d.Desde.Should().Be("VIE-DOM");
            d.Hacia.Should().Be("JUE-DOM");
            d.CapacidadAntes.Should().Be(4 * w);
            d.CapacidadDespues.Should().Be(5 * w);
            d.Aplica.Should().BeTrue();
        }

        [Fact]
        public void Evaluar_WhenDiaAnadidoEsFestivo_ShouldNoProponerEscaladoPorFaltaDeGanancia()
        {
            // Arrange: W = 1 semana (2026-08-03 al 2026-08-09). Viernes 2026-08-07 festivo.
            int w = 1;
            var festivos = new[] { (new DateTime(2026, 8, 7), new DateTime(2026, 8, 7)) };
            var contexto = CrearContextoNominal(w, "LUN-JUE", 5 * w, festivos: festivos);
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert: LUN-VIE tiene 5 días pero el viernes festivo reduce la capacidad a 4 (sin ganancia)
            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_WhenBloqueDelDiaAnadidoAusente_ShouldNoProponerEscaladoPorFaltaDeGanancia()
        {
            // Arrange: W = 4 semanas, sin bloque para el día 5 (Viernes)
            int w = 4;
            var contexto = CrearContextoNominal(w, "LUN-JUE", 5 * w, diasBloquesExcluidos: new[] { 5 });
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert: Sin bloque para el viernes, LUN-VIE no aporta ranuras adicionales
            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_Caso1703_WhenPrimerPasoSinGananciaYSegundoConGanancia_ShouldNoProponerNingunaSugerencia()
        {
            // Arrange: VIE-SAB (escalones: JUE-SAB, 4,5,6,7).
            // Paso 1 (JUE-SAB) agrega día 4 (Jueves), pero día 4 no tiene bloque activo.
            // Paso 2 (4,5,6,7) agrega día 7 (Domingo), que sí tiene ranura.
            // Bajo el caso 1703, solo se evalúa Siguiente ("JUE-SAB"). Como no tiene ganancia, NO se salta a "4,5,6,7".
            int w = 4;
            var contexto = CrearContextoNominal(w, "VIE-SAB", 5 * w, diasBloquesExcluidos: new[] { 4 });
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var decisiones = sut.Evaluar(contexto);

            // Assert: Estrictamente vacío; no debe saltar al paso 2
            decisiones.Should().BeEmpty();
        }

        [Fact]
        public void CopiaSimulada_DisponibleAislaModelosYPatrones_YConservaSuResultado()
        {
            var c = CrearContextoNominal(4, "MAR-VIE", 20);
            var sut = new AgEscaladorRanuraAula(() => null);
            var antes = sut.Evaluar(c)[0];
            var copia = AgEscaladorRanuraAula.ClonarContextoParaSimulacion(c);
            copia.ConstruirIndices();
            copia.Aulas[0].DiasVentana = "LUN-VIE";
            copia.Ventana.FechaFin = copia.Ventana.FechaFin.AddDays(1);
            copia.Asignaturas[0].Horas = 4;
            copia.Bloques[0].Activo = 0;
            copia.AulaAsignaturas[0].DiasPatron = "1";
            c.Aulas[0].DiasVentana.Should().Be("MAR-VIE");
            c.AulaAsignaturas[0].DiasPatron.Should().BeEmpty();
            c.Asignaturas[0].Horas.Should().Be(80);
            c.Bloques[0].Activo.Should().Be(1);
            var despues = sut.Evaluar(c)[0];
            despues.Hacia.Should().Be(antes.Hacia);
            despues.CapacidadAntes.Should().Be(antes.CapacidadAntes);
            despues.CapacidadDespues.Should().Be(antes.CapacidadDespues);
            despues.DemandaPendiente.Should().Be(antes.DemandaPendiente);
        }

        [Fact]
        public void DecisionPersistible_ExcluyeAprobadasInclusoConSeleccionAutorizadaEnMemoria()
        {
            var c = CrearContextoNominal(1, "MAR-VIE", 5);
            c.Aulas[0].Estado = AgEstadoAula.Aprobado;
            var sut = new AgEscaladorRanuraAula(() => throw new InvalidOperationException("No debe abrir SQL"));
            sut.Evaluar(c).Should().BeEmpty();
            AgEscaladorRanuraAula.EsAulaElegible(c.Aulas[0]).Should().BeFalse();
            AgEscaladorRanuraAula.PuedeAplicarPropuesta(c.Aulas[0]).Should().BeFalse();
            var seleccion = sut.ValidarSelecciones(c, new[] { new AgReclasificacionPatron {
                IdAula = 1, AulaRef = "A101", Subunidad = "", IdAsignatura = 101,
                OrigenAnterior = "LEGADO", OrigenPropuesto = "AUTOMATICO", DiasPatronAnterior = "2,3", Autorizado = true } });
            seleccion.EsValido.Should().BeFalse();
            seleccion.Motivo.Should().Be("AULA_NO_ELEGIBLE");
        }

        private static AgContexto CrearContextoNominal(
            int w,
            string diasVentana,
            int demandaSesiones,
            IReadOnlyCollection<(DateTime Inicio, DateTime Fin)> festivos = null,
            IReadOnlyCollection<int> diasBloquesExcluidos = null)
        {
            var fInicio = new DateTime(2026, 8, 3); // Lunes
            var fFin = fInicio.AddDays(7 * w - 1);  // Domingo

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(
                1, "2026-2", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Test",
                Modalidad = "PRESENCIAL"
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "A101",
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30,
                IdPrograma = 1,
                Semestre = "I"
            };
            contexto.Aulas.Add(aula);

            var excluidos = new HashSet<int>(diasBloquesExcluidos ?? new int[0]);
            for (int d = 1; d <= 7; d++)
            {
                if (excluidos.Contains(d)) continue;

                if (d == 6)
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 61,
                        DiaSemanaNumero = 6,
                        Franja = "MANANA",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 62,
                        DiaSemanaNumero = 6,
                        Franja = "TARDE",
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
                        DuracionHoras = 4,
                        Activo = 1
                    });
                }
            }

            double horasTotales = demandaSesiones * 4.0;
            contexto.Asignaturas.Add(new Asignatura
            {
                Id = 101,
                Nombre = "Asig Demanda",
                IdPrograma = 1,
                Semestre = "I",
                Horas = (int)horasTotales
            });

            var aa = new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 101,
                HorasObjetivo = horasTotales,
                OrdenCurricular = 1
            };
            contexto.AulaAsignaturas.Add(aa);
            contexto.AulaAsignaturaPorAulaAsig.Add((1, 101), aa);

            if (festivos != null)
            {
                contexto.Festivos.AddRange(festivos);
            }
            contexto.ConstruirIndices();

            return contexto;
        }

        [Fact]
        public void DecisionSoloParticion_CuandoDifierePeroNoMejora_DevuelveTrue()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "VIE-SAB",
                Hacia = "VIE-SAB",
                ParticionMejora = false,
                ParticionDifiereDeReferencia = true,
                CapacidadAntes = 6,
                CapacidadDespues = 6
            };

            decision.SoloParticion.Should().BeTrue();
            decision.Aplica.Should().BeFalse();
        }

        [Fact]
        public void DecisionConservaPatronFalseConBottleneck_DevuelveAplicaTrue()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                Desde = "MAR-VIE",
                Hacia = "MAR-VIE",
                ConservaPatronOrigen = false,
                IdAsignaturasBottleneck = new List<int> { 101 },
                ParticionMejora = true,
                CapacidadAntes = 6,
                CapacidadDespues = 6
            };

            decision.Aplica.Should().BeTrue();
            decision.SoloParticion.Should().BeFalse();
        }

        [Fact]
        public void DecisionAplica_CuandoMejoraConMismaCapacidad_DevuelveAplicaTrue()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                Desde = "LUN-JUE",
                Hacia = "MAR-VIE",
                Mejora = true,
                CapacidadAntes = 16,
                CapacidadDespues = 16
            };

            decision.Aplica.Should().BeTrue();
            decision.SoloParticion.Should().BeFalse();
        }

        [Fact]
        public void HuellaAccion_DebeComenzarConPropuestaEscaladoV5()
        {
            var decision = new AgEscaladorRanuraAula.Decision
            {
                IdAula = 1,
                AulaRef = "A101",
                Subunidad = "CENTRAL",
                Desde = "LUN-JUE",
                Hacia = "LUN-VIE",
                ConservaPatronOrigen = true,
                CapacidadAntes = 16,
                CapacidadDespues = 20
            };

            AgEscaladorRanuraAula.HuellaAccionPartes(decision)[0]
                .Should().Be("propuesta-escalado-v5");

            string huella = AgEscaladorRanuraAula.HuellaAccion(decision);

            huella.Should().NotBeNullOrWhiteSpace();
            huella.Length.Should().Be(44);
        }
    }
}
