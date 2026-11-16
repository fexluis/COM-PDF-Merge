using System;
using System.Collections.Generic;
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
    /// Suite de aceptación end-to-end con 12 criterios exhaustivos (Corrida 92 - T7 / Fase 1804).
    /// Valida ocupación de huecos, paridad entre generación y evaluación, preservación de restricciones,
    /// independencia de año, fidelidad en fábrica/expansor y escalado sin regresiones.
    /// </summary>
    public class AgPlanificadorRanurasAceptacionTests
    {
        private static readonly Func<DateTime, bool> EsFestivoCanonico = d =>
            (d.Date >= new DateTime(2027, 3, 22) && d.Date <= new DateTime(2027, 3, 28))
            || d.Date == new DateTime(2027, 5, 10);

        /// <summary>
        /// Criterio 1: Cali III y IV completan horas exactas (256/272 h) antes del límite estipulado
        /// (13/05 para III, 19/05 para IV) y cero sesiones posteriores al 30/05/2027 en un plan nuevo.
        /// El completado con referencia explícita conserva las fechas exactas de las 5 materias restringidas.
        /// </summary>
        [Theory]
        [InlineData(false, 13)] // Cali III: 32, 48, 48, 48, 32, 48 h -> 256 h (64 sesiones) -> límite 13/05/2027
        [InlineData(true, 19)]  // Cali IV:  32, 48, 48, 48, 48, 48 h -> 272 h (68 sesiones) -> límite 19/05/2027
        public void Criterio1_CaliIII_IV_HorasExactasYCompletadasAntesDeLimite(bool cuarto, int diaFinEsperado)
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var fLimiteEsperado = new DateTime(2027, 5, diaFinEsperado);
            int[] horas = cuarto
                ? new int[] { 32, 48, 48, 48, 48, 48 }
                : new int[] { 32, 48, 48, 48, 32, 48 };

            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");
            var demandas = horas.Select((h, i) => (IdAsignatura: i + 1, Horas: (double)h, OrdenCurricular: i + 1)).ToList();

            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3,4" },
                { 3, "5" },
                { 4, "1,2" },
                { 5, "3,4" }
                // Asignatura 6: AUTOMATICO
            };

            var sut = new AgPlanificadorOleadas();

            // Plan de referencia previo con solo las 5 materias restringidas
            var demandasPrevias = demandas.Take(5).ToList();
            var resPrevio = sut.Construir(demandasPrevias, "LUN-VIE", fInicio, fFin, 4.0, EsFestivoCanonico, bloques, "TARDE", patrones);
            resPrevio.IsSuccess.Should().BeTrue();

            // Act
            var resMixto = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, EsFestivoCanonico, bloques, "TARDE", patrones);

            // Assert
            resMixto.IsSuccess.Should().BeTrue();
            resMixto.Value.Infactibles.Should().BeEmpty();

            var sesionesMixtas = resMixto.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            int horasTotales = horas.Sum();
            int sesionesEsperadas = horasTotales / 4;

            // 1. Horas exactas (256h o 272h) y conteo exacto de sesiones (64 o 68)
            sesionesMixtas.Count.Should().Be(sesionesEsperadas);
            sesionesMixtas.Sum(s => s.Ranura.Horas).Should().Be(horasTotales);

            // 2. Cero sesiones posteriores al 30/05/2027
            sesionesMixtas.Should().OnlyContain(s => s.Fecha <= fFin);

            // 3. Terminación antes del límite previsto (13/05 o 19/05)
            sesionesMixtas.Max(s => s.Fecha).Should().BeOnOrBefore(fLimiteEsperado);

            // 4. El plan nuevo respeta los días restringidos, sin heredar fechas de otra llamada.
            foreach (var patron in patrones)
            {
                sesionesMixtas.Where(s => s.IdAsignatura == patron.Key).Should()
                    .OnlyContain(s => AgCarriles.Ventana(patron.Value).Contains(s.Ranura.Dia));
            }

            // 5. La preservación exacta requiere suministrar la referencia protegida al completado.
            resPrevio.Value.Infactibles.Should().BeEmpty();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", bloques, "TARDE", 4.0);
            // La demanda ampliada declara pendientes solo las materias ausentes de la referencia previa.
            var previo = resPrevio.Value;
            var nuevasPendientes = demandas.Select(d => d.IdAsignatura)
                .Except(previo.Planes.Select(p => p.IdAsignatura).Union(previo.Infactibles));
            var referenciaAmpliada = new AgPlanOleadasResultado(previo.Planes,
                previo.Infactibles.Union(nuevasPendientes).ToList(), previo.DetalleInfactibles,
                previo.Particion, previo.VentanaSugerida, previo.PatronesRenormalizados,
                previo.OcupacionPorcentaje, previo.BusquedaAgotada);
            var completado = AgPlanificadorRanuras.CompletarReferencia(demandas, ranuras,
                fInicio, fFin, 4.0, EsFestivoCanonico, patrones, referenciaAmpliada, CancellationToken.None);
            completado.IsSuccess.Should().BeTrue(completado.IsFailure ? completado.Error : string.Empty);
            completado.Value.Infactibles.Should().BeEmpty();
            completado.Value.Planes.Sum(p => p.Sesiones.Count).Should().Be(sesionesEsperadas);
            completado.Value.Planes.SelectMany(p => p.Sesiones).Sum(s => s.Ranura.Horas).Should().Be(horasTotales);
            for (int asigId = 1; asigId <= 5; asigId++)
            {
                var fechasPrevias = resPrevio.Value.Planes.Single(p => p.IdAsignatura == asigId).Sesiones.Select(s => s.Fecha).OrderBy(d => d).ToList();
                var fechasCompletadas = completado.Value.Planes.Single(p => p.IdAsignatura == asigId).Sesiones.Select(s => s.Fecha).OrderBy(d => d).ToList();
                fechasCompletadas.Should().Equal(fechasPrevias);
            }
        }

        /// <summary>
        /// Criterio 2: Construir y ConstruirEvaluacion arrojan idénticas sesiones en fecha y ranura
        /// para mezclas LEGADO/AUTOMATICO sin ampliar el plazo ni introducir discrepancias.
        /// </summary>
        [Fact]
        public void Criterio2_ParidadEstricta_GeneracionYEvaluacion()
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 41, DiaSemanaNumero = 4, DuracionHoras = 4.0, Activo = 1, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00" },
                new AgBloqueHorario { Id = 51, DiaSemanaNumero = 5, DuracionHoras = 4.0, Activo = 1, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00" },
                new AgBloqueHorario { Id = 61, DiaSemanaNumero = 6, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" },
                new AgBloqueHorario { Id = 62, DiaSemanaNumero = 6, DuracionHoras = 4.0, Activo = 1, Franja = "TARDE", HoraInicio = "13:00", HoraFin = "17:00" },
                new AgBloqueHorario { Id = 71, DiaSemanaNumero = 7, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" }
            };
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 64.0, 3),
                (4, 112.0, 4) // AUTOMATICO
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "4,5" },
                { 2, "6" },
                { 3, "4,5" }
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resGeneracion = sut.Construir(demandas, "JUE-DOM", fInicio, fFin, 4.0, EsFestivoCanonico, bloques, "NOCHE", patrones);
            var resEvaluacion = sut.ConstruirEvaluacion(demandas, "JUE-DOM", fInicio, fFin, 4.0, EsFestivoCanonico, bloques, "NOCHE", patrones);

            // Assert
            resGeneracion.IsSuccess.Should().BeTrue();
            resEvaluacion.IsSuccess.Should().BeTrue();
            resGeneracion.Value.Infactibles.Should().BeEmpty();
            resEvaluacion.Value.Infactibles.Should().BeEmpty();

            var llavesGen = resGeneracion.Value.Planes.SelectMany(p => p.Sesiones)
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .OrderBy(x => x)
                .ToList();
            var llavesEval = resEvaluacion.Value.Planes.SelectMany(p => p.Sesiones)
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .OrderBy(x => x)
                .ToList();

            llavesGen.Should().Equal(llavesEval);
            llavesGen.Count.Should().Be(60);
        }

        /// <summary>
        /// Criterio 3: Una materia automática aprovecha múltiples días disponibles (ej. LUN, MIE, VIE)
        /// sin segregación por carril ni exclusividad artificial de banda.
        /// </summary>
        [Fact]
        public void Criterio3_MateriaAutomatica_AprovechaMultiplesDiasSinExclusividadArtificial()
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 3, 31);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            // Asig 1 fijada en Martes y Jueves ("2,4"), dejando libres Lunes (1), Miércoles (3) y Viernes (5)
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 48.0, 2) // AUTOMATICO (12 sesiones)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "2,4" }
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var result = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();

            var planAsig2 = result.Value.Planes.Single(p => p.IdAsignatura == 2);
            planAsig2.Sesiones.Count.Should().Be(12);

            // La materia automática debe usar al menos 3 días distintos aprovechando los huecos disponibles
            var diasUsados = planAsig2.Sesiones.Select(s => s.Ranura.Dia).Distinct().ToList();
            diasUsados.Count.Should().BeGreaterThanOrEqualTo(3);
            diasUsados.Should().Contain(new int[] { 1, 3, 5 });
        }

        /// <summary>
        /// Criterio 4: Dos materias automáticas comparten calendario ocupando huecos sin solapes,
        /// reutilizando los días liberados una vez que concluyen las materias precedentes.
        /// </summary>
        [Fact]
        public void Criterio4_ComparticionCalendario_SinColisionesYReusoHuecos()
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 4, 30);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            // Asig 1 restringida a días 1 y 2 (LUN, MAR, 8 sesiones).
            // Asig 2 y Asig 3 son automáticas (8 sesiones cada una).
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 32.0, 3)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" }
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var result = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();

            var todasSesiones = result.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            todasSesiones.Count.Should().Be(24);

            // 1. Cero colisiones en el calendario común
            todasSesiones.Select(s => s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .Should().OnlyHaveUniqueItems();

            // 2. Reuso de huecos: las automáticas toman días 3, 4, 5 y luego también días 1 y 2
            var sesionesAuto = todasSesiones.Where(s => s.IdAsignatura == 2 || s.IdAsignatura == 3).ToList();
            var diasAuto = sesionesAuto.Select(s => s.Ranura.Dia).Distinct().ToList();
            diasAuto.Should().Contain(new int[] { 1, 2, 3, 4, 5 });
        }

        /// <summary>
        /// Criterio 5: Las materias restringidas (LEGADO/FIJO) no autorizadas a reclasificación
        /// conservan estrictamente sus días asignados.
        /// </summary>
        [Fact]
        public void Criterio5_MateriasRestringidas_NoReclasificadas_ConservanDiasAsignados()
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 5, 30);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 32.0, 3),
                (4, 48.0, 4) // AUTOMATICO
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3" },
                { 3, "4,5" }
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var result = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();

            var sesiones = result.Value.Planes.SelectMany(p => p.Sesiones).ToList();

            // Invariantes estrictos de días permitidos
            sesiones.Where(s => s.IdAsignatura == 1).Should().OnlyContain(s => s.Ranura.Dia == 1 || s.Ranura.Dia == 2);
            sesiones.Where(s => s.IdAsignatura == 2).Should().OnlyContain(s => s.Ranura.Dia == 3);
            sesiones.Where(s => s.IdAsignatura == 3).Should().OnlyContain(s => s.Ranura.Dia == 4 || s.Ranura.Dia == 5);
        }

        /// <summary>
        /// Criterio 6: Un plan con terminación anterior por menos de 7 días es elegido sobre uno más tardío,
        /// comprobando que el antiguo umbral rígido de 7 días fue completamente removido.
        /// </summary>
        [Fact]
        public void Criterio6_SeleccionPlan_MasTempranoMenor7Dias()
        {
            // Arrange
            // Ventana de febrero 2027: Asig 1 en "1,2" (16h), Asig 2 en "3,4" (16h).
            // Con la 3ra materia en patrón "5" (12h, 3 sesiones), termina el viernes 19/02/2027.
            // Con la 3ra materia en AUTOMATICO, aprovecha los viernes iniciales y el lunes 15/02 liberado,
            // terminando el lunes 15/02/2027 (diferencia de 4 días, estrictamente menor a 7 días).
            var fInicio = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 2, 28);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1),
                (2, 16.0, 2),
                (3, 12.0, 3)
            };

            var patronesRigidos = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3,4" },
                { 3, "5" }
            };

            var patronesMixtos = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3,4" }
                // Asig 3: AUTOMATICO
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resRigido = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patronesRigidos);
            var resMixto = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patronesMixtos);

            // Assert
            resRigido.IsSuccess.Should().BeTrue();
            resRigido.Value.Infactibles.Should().BeEmpty();

            resMixto.IsSuccess.Should().BeTrue();
            resMixto.Value.Infactibles.Should().BeEmpty();

            DateTime finRigido = resRigido.Value.Planes.SelectMany(p => p.Sesiones).Max(s => s.Fecha);
            DateTime finMixto = resMixto.Value.Planes.SelectMany(p => p.Sesiones).Max(s => s.Fecha);

            // El plan mixto termina antes que el rígido
            finMixto.Should().BeBefore(finRigido);

            // La ganancia es estrictamente menor a 7 días (4 días de mejora: 19/02 a 15/02)
            double diasDiferencia = (finRigido - finMixto).TotalDays;
            diasDiferencia.Should().BeInRange(1.0, 6.0);
            finRigido.Should().Be(new DateTime(2027, 2, 19));
            finMixto.Should().Be(new DateTime(2027, 2, 15));
        }

        /// <summary>
        /// Criterio 7: En evaluaciones parciales donde la ventana no alcanza para toda la demanda,
        /// un candidato alternativo no degrada materias previamente completas en la referencia.
        /// </summary>
        [Fact]
        public void Criterio7_PreservacionMaterias_PreviamenteCompletasEnParciales()
        {
            // Arrange
            // Ventana reducida de 2 semanas (01/02 al 12/02 = 10 sesiones hábiles de 4h)
            var fInicio = new DateTime(2027, 2, 1);
            var fFin = new DateTime(2027, 2, 12);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            // Demanda total = 16h (4 sesiones) + 32h (8 sesiones) = 48h (12 sesiones requeridas > 10 disponibles)
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1),
                (2, 32.0, 2)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" } // Asig 1 completa cabe holgadamente en las 4 sesiones de días 1 y 2
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resEval = sut.ConstruirEvaluacion(
                demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);

            // Assert
            resEval.IsSuccess.Should().BeTrue();

            // Asignatura 1 conserva el 100% de sus sesiones requeridas (4 de 4)
            var planAsig1 = resEval.Value.Planes.Single(p => p.IdAsignatura == 1);
            planAsig1.Sesiones.Count.Should().Be(4);
            resEval.Value.Infactibles.Should().NotContain(1);

            // Asignatura 2 es la que absorbe el déficit y se declara en Infactibles
            resEval.Value.Infactibles.Should().Contain(2);
            var planAsig2 = resEval.Value.Planes.Single(p => p.IdAsignatura == 2);
            planAsig2.Sesiones.Count.Should().Be(6); // 10 sesiones totales - 4 de Asig 1 = 6 para Asig 2
        }

        /// <summary>
        /// Criterio 8: ValidarFechado rechaza estrictamente anomalías de calendario:
        /// plan nulo, colisiones, festivos, fechas fuera de ventana, patrones inválidos e inconsistencia de pendientes.
        /// </summary>
        [Fact]
        public void Criterio8_RechazoEstricto_AnomaliasCalendario()
        {
            // Arrange base
            var ini = new DateTime(2027, 2, 1);
            var fin = new DateTime(2027, 2, 28);
            var ranura = new AgRanura(1, "TARDE", 4.0);
            var ranuras = new List<AgRanura> { ranura };
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4.0, 1)
            };
            Func<DateTime, bool> festivo = d => d.Date == new DateTime(2027, 2, 8);

            // 1. Plan nulo
            var resNulo = AgPlanificadorRanuras.ValidarFechado(null, demandas, ranuras, ini, fin, 4.0, festivo, null, CancellationToken.None);
            resNulo.IsFailure.Should().BeTrue();
            resNulo.Error.Should().Be("I5_PLAN_NULO");

            // 2. Sesión en fecha festiva
            var sesionFestiva = new AgSesionPlanificada(1, new DateTime(2027, 2, 8), ranura);
            var planFestivo = CrearPlan(1, new List<AgSesionPlanificada> { sesionFestiva }, ranura, ini, fin, 1);
            var resFestivo = AgPlanificadorRanuras.ValidarFechado(new AgPlanOleadasResultado(new List<AgPlanOleada> { planFestivo }, new List<int>()), demandas, ranuras, ini, fin, 4.0, festivo, null, CancellationToken.None);
            resFestivo.IsFailure.Should().BeTrue();
            resFestivo.Error.Should().Be("I5_PLAN_FECHA_O_COLISION");

            // 3. Sesión fuera de rango
            var sesionFueraRango = new AgSesionPlanificada(1, new DateTime(2027, 3, 1), ranura);
            var planFuera = CrearPlan(1, new List<AgSesionPlanificada> { sesionFueraRango }, ranura, ini, fin, 1);
            var resFuera = AgPlanificadorRanuras.ValidarFechado(new AgPlanOleadasResultado(new List<AgPlanOleada> { planFuera }, new List<int>()), demandas, ranuras, ini, fin, 4.0, null, null, CancellationToken.None);
            resFuera.IsFailure.Should().BeTrue();
            resFuera.Error.Should().Be("I5_PLAN_SESION_INVALIDA");

            // 4. Colisión horaria (dos sesiones en la misma fecha y ranura)
            var demandasColision = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4.0, 1), (2, 4.0, 2)
            };
            var fechaColision = new DateTime(2027, 2, 1);
            var sesCol1 = new AgSesionPlanificada(1, fechaColision, ranura);
            var sesCol2 = new AgSesionPlanificada(2, fechaColision, ranura);
            var planCol1 = CrearPlan(1, new List<AgSesionPlanificada> { sesCol1 }, ranura, ini, fin, 1);
            var planCol2 = CrearPlan(2, new List<AgSesionPlanificada> { sesCol2 }, ranura, ini, fin, 1);
            var resColision = AgPlanificadorRanuras.ValidarFechado(new AgPlanOleadasResultado(new List<AgPlanOleada> { planCol1, planCol2 }, new List<int>()), demandasColision, ranuras, ini, fin, 4.0, null, null, CancellationToken.None);
            resColision.IsFailure.Should().BeTrue();
            resColision.Error.Should().Be("I5_PLAN_FECHA_O_COLISION");

            // 5. Inconsistencia de pendientes (materia con déficit no listada en Infactibles)
            var planVacio = CrearPlan(1, new List<AgSesionPlanificada>(), ranura, ini, fin, 1);
            var resPendientes = AgPlanificadorRanuras.ValidarFechado(new AgPlanOleadasResultado(new List<AgPlanOleada> { planVacio }, new List<int>()), demandas, ranuras, ini, fin, 4.0, null, null, CancellationToken.None);
            resPendientes.IsFailure.Should().BeTrue();
            resPendientes.Error.Should().Be("I5_PLAN_PENDIENTES_INCONSISTENTES");
        }

        /// <summary>
        /// Criterio 9: Planificador resuelve escenarios de unidad de 8h, distinguiendo franjas de sábado
        /// (mañana y tarde) y soportando semanas incompletas sin declarar falsas infactibilidades.
        /// </summary>
        [Fact]
        public void Criterio9_SoporteSemanasIncompletas_SabadoMT_Unidad8h()
        {
            // Arrange
            // Período que empieza un sábado (semana incompleta)
            var fInicio = new DateTime(2027, 2, 6); // Sábado
            var fFin = new DateTime(2027, 2, 28);   // Domingo
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 61, DiaSemanaNumero = 6, DuracionHoras = 8.0, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "16:00", Activo = 1 },
                new AgBloqueHorario { Id = 62, DiaSemanaNumero = 6, DuracionHoras = 8.0, Franja = "TARDE", HoraInicio = "13:00", HoraFin = "21:00", Activo = 1 },
                new AgBloqueHorario { Id = 71, DiaSemanaNumero = 7, DuracionHoras = 8.0, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "16:00", Activo = 1 }
            };

            // 2 asignaturas de posgrado (16h = 2 sesiones, 24h = 3 sesiones; total 5 sesiones de 8h)
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1),
                (2, 24.0, 2)
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var result = sut.Construir(demandas, "SAB-DOM", fInicio, fFin, 8.0, null, bloques, "MANANA", null);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();

            var sesiones = result.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            sesiones.Count.Should().Be(5);
            sesiones.Should().OnlyContain(s => Math.Abs(s.Ranura.Horas - 8.0) < 0.001);

            // Cero colisiones en sábado distinguiendo franjas de Mañana y Tarde
            sesiones.Where(s => s.Ranura.Dia == 6)
                .Select(s => s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Franja)
                .Should().OnlyHaveUniqueItems();
        }

        /// <summary>
        /// Criterio 10: La planificación es completamente independiente del año 2027 y de constantes rígidas
        /// (como el antiguo limiteMayo=30/05/2027), resolviendo períodos de 2028 que se extienden a junio.
        /// </summary>
        [Fact]
        public void Criterio10_IndependenciaAnio_PeriodoDistinto2027()
        {
            // Arrange
            // Período en 2028 (año bisiesto) que se extiende hasta el 30 de junio de 2028
            var fInicio = new DateTime(2028, 2, 1);
            var fFin = new DateTime(2028, 6, 30);
            var bloques = CrearBloquesLunesViernes(4.0, "TARDE");

            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 48.0, 2),
                (3, 48.0, 3)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3,4" }
                // Asig 3: AUTOMATICO
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resGen = sut.Construir(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);
            var resEval = sut.ConstruirEvaluacion(demandas, "LUN-VIE", fInicio, fFin, 4.0, null, bloques, "TARDE", patrones);

            // Assert
            resGen.IsSuccess.Should().BeTrue();
            resGen.Value.Infactibles.Should().BeEmpty();
            resEval.IsSuccess.Should().BeTrue();
            resEval.Value.Infactibles.Should().BeEmpty();

            var sesionesGen = resGen.Value.Planes.SelectMany(p => p.Sesiones).ToList();

            // Todas las sesiones pertenecen estrictamente al año 2028
            sesionesGen.Should().OnlyContain(s => s.Fecha.Year == 2028);

            // Cero colisiones horarias
            sesionesGen.Select(s => s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia)
                .Should().OnlyHaveUniqueItems();

            // Paridad exacta entre evaluación y generación en 2028
            var llavesGen = sesionesGen.Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia).OrderBy(x => x).ToList();
            var llavesEval = resEval.Value.Planes.SelectMany(p => p.Sesiones).Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia).OrderBy(x => x).ToList();
            llavesGen.Should().Equal(llavesEval);
        }

        /// <summary>
        /// Criterio 11: Cobertura integral y trazabilidad estricta de extremo a extremo:
        /// Plan Original -> Fábrica -> AgCoherenciaGrupo -> AgExpansorSesiones -> AgMaterializadorSesiones (Staging).
        /// Verifica paridad exacta en asignatura, fecha, bloque y horas hasta la persistencia en staging.
        /// </summary>
        [Fact]
        public async Task Criterio11_TrazabilidadYFidelidadDeFechas_PlanCoherenciaStaging()
        {
            // Arrange
            var ctx = CrearContexto(new int[] { 32, 48, 48 });
            ctx.AulaAsignaturas[0].DiasPatron = "1,2";
            ctx.AulaAsignaturas[1].DiasPatron = "3,4";
            ctx.AulaAsignaturas[2].DiasPatron = "5";
            foreach (var aa in ctx.AulaAsignaturas) { aa.OrigenPatron = "LEGADO"; }
            ctx.ConstruirIndices();
            var planificador = new AgPlanificadorOleadas();

            // 1. Obtener el Plan Original
            var aula = ctx.Aulas[0];
            var demandas = ctx.AulaAsignaturas.Where(a => a.IdAula == aula.Id)
                .OrderBy(a => a.OrdenCurricular)
                .Select(a => (IdAsignatura: a.IdAsignatura, Horas: a.HorasObjetivo, OrdenCurricular: a.OrdenCurricular))
                .ToList();
            var patrones = ctx.AulaAsignaturas.Where(a => a.IdAula == aula.Id)
                .ToDictionary(a => a.IdAsignatura, a => a.DiasPatron);

            var planOriginalRes = planificador.Construir(
                demandas, aula.DiasVentana, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin,
                4.0, ctx.EsFestivo, ctx.Bloques, aula.Jornada, patrones, diagnostico: "Criterio11");
            planOriginalRes.IsSuccess.Should().BeTrue();
            var planOriginal = planOriginalRes.Value;

            // 2. Generar individuo con la Fábrica
            var fabrica = new AgFabricaIndividuos(new Random(42), null, planificador);
            var individuo = fabrica.CrearAleatorio(ctx, 1, 0);
            individuo.Should().NotBeNull();

            // 3. Pasar por Coherencia de Grupo (Sellar)
            AgCoherenciaGrupo.Sellar(individuo, ctx);

            // 4. Materializar y persistir en Staging usando persistencia SQLite real
            string dbPath = Path.Combine(Path.GetTempPath(), "AgStagingCriterio11_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbPath);
            string connStr = $"Data Source={dbPath};Version=3;";

            try
            {
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE ag_vinculaciones (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            id_corrida              INTEGER NOT NULL DEFAULT 0,
                            id_individuo            INTEGER NOT NULL DEFAULT 0,
                            idtercero               INTEGER NOT NULL DEFAULT 0,
                            idprograma              INTEGER NOT NULL DEFAULT 0,
                            idsede                  INTEGER NOT NULL DEFAULT 0,
                            idasignatura            INTEGER NOT NULL DEFAULT 0,
                            idperiodo               INTEGER NOT NULL DEFAULT 0,
                            semestre                TEXT    NOT NULL DEFAULT '',
                            grupo                   TEXT    NOT NULL DEFAULT '',
                            dia_horas_semanales     TEXT    NOT NULL DEFAULT '',
                            modalidadVinculacion    TEXT    NOT NULL DEFAULT 'Presencial',
                            total_horas             INTEGER NOT NULL DEFAULT 0,
                            fecha_inicio            TEXT    NOT NULL DEFAULT '',
                            fecha_fin               TEXT    NOT NULL DEFAULT '',
                            estado                  INTEGER NOT NULL DEFAULT 0,
                            isTerminalAereo         INTEGER NOT NULL DEFAULT 0,
                            idTerminales            INTEGER NOT NULL DEFAULT 0,
                            gastoTerminalAerea      REAL    NOT NULL DEFAULT 0.0,
                            es_vacante              INTEGER NOT NULL DEFAULT 0,
                            motivo_vacante          TEXT    NOT NULL DEFAULT '',
                            candidatos_pool         INTEGER NOT NULL DEFAULT 0,
                            subunidad               TEXT    NOT NULL DEFAULT ''
                        );
                        CREATE TABLE ag_vinculacion_gen (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            idagvinculacion         INTEGER NOT NULL,
                            idgen                   INTEGER NOT NULL,
                            subunidad               TEXT    DEFAULT ''
                        );
                        CREATE TABLE ag_sesiones (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            idvinculacion           INTEGER NOT NULL,
                            idtercero               INTEGER NOT NULL,
                            id_aula                 INTEGER NOT NULL DEFAULT 0,
                            fechasesion             TEXT    NOT NULL,
                            horariosesion           TEXT    NOT NULL DEFAULT '',
                            horas                   REAL    NOT NULL DEFAULT 0.0,
                            viatica                 INTEGER NOT NULL DEFAULT 0,
                            subunidad               TEXT    NOT NULL DEFAULT '',
                            idindividuo             INTEGER NULL,
                            idgen                   INTEGER NULL,
                            origen_ag               INTEGER NOT NULL DEFAULT 1,
                            estado_ag               TEXT    NOT NULL DEFAULT 'BORRADOR',
                            es_vacante              INTEGER NOT NULL DEFAULT 0,
                            motivo_vacante          TEXT    NOT NULL DEFAULT '',
                            candidatos_pool         INTEGER NOT NULL DEFAULT 0,
                            fecha_creacion          TEXT    NOT NULL DEFAULT (datetime('now'))
                        );
                    ");
                }

                var repoReal = new AgScheduleRepository(() => new SQLiteConnection(connStr));
                var expansor = new AgExpansorSesiones();
                var materializador = new AgMaterializadorSesiones(repoReal, expansor);

                // Act
                var resMat = await materializador.MaterializarEnStagingAsync(1, individuo, ctx, CancellationToken.None);

                // Assert
                resMat.IsSuccess.Should().BeTrue(resMat.IsFailure ? resMat.Error : string.Empty);

                // 5. Consultar los datos persistidos en SQLite
                List<AgVinculacion> guardadasVinc;
                List<AgSesion> guardadasSes;
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    guardadasVinc = (await conn.QueryAsync<AgVinculacion>(
                        "SELECT id as Id, id_corrida as IdCorrida, id_individuo as IdIndividuo, idasignatura as IdAsignatura, total_horas as TotalHoras, grupo as Grupo, subunidad as Subunidad FROM ag_vinculaciones WHERE id_individuo = @IdIndividuo",
                        new { IdIndividuo = individuo.Id })).ToList();

                    guardadasSes = (await conn.QueryAsync<AgSesion>(
                        "SELECT id as Id, idvinculacion as IdVinculacion, id_aula as IdAula, fechasesion as FechaSesion, horariosesion as HorarioSesion, horas as Horas, idindividuo as IdIndividuo, idgen as IdGen FROM ag_sesiones WHERE idindividuo = @IdIndividuo",
                        new { IdIndividuo = individuo.Id })).ToList();
                }

                guardadasVinc.Should().NotBeNullOrEmpty();
                guardadasSes.Should().NotBeNullOrEmpty();

                int totalSesionesPlan = planOriginal.Planes.Sum(p => p.Sesiones.Count);
                double totalHorasPlan = planOriginal.Planes.Sum(p => p.Sesiones.Sum(s => s.Ranura.Horas));

                guardadasSes.Count.Should().Be(totalSesionesPlan);
                guardadasSes.Sum(s => s.Horas).Should().Be(totalHorasPlan);

                foreach (var planAsig in planOriginal.Planes)
                {
                    var vinculacion = guardadasVinc.FirstOrDefault(v => v.IdAsignatura == planAsig.IdAsignatura);
                    vinculacion.Should().NotBeNull();

                    var sesionesStagingAsig = guardadasSes
                        .Where(s => s.IdVinculacion == vinculacion.Id)
                        .OrderBy(s => s.FechaSesion.Value)
                        .ToList();

                    sesionesStagingAsig.Count.Should().Be(planAsig.Sesiones.Count);

                    var sesionesPlanOrdenadas = planAsig.Sesiones.OrderBy(s => s.Fecha).ToList();

                    for (int sIdx = 0; sIdx < sesionesPlanOrdenadas.Count; sIdx++)
                    {
                        var planSesion = sesionesPlanOrdenadas[sIdx];
                        var stagingSesion = sesionesStagingAsig[sIdx];

                        // Comparar asignatura, fecha, horas
                        stagingSesion.FechaSesion.Value.Date.Should().Be(planSesion.Fecha.Date);
                        stagingSesion.Horas.Should().Be(planSesion.Ranura.Horas);
                        stagingSesion.IdAula.Should().Be(aula.Id);

                        // Comparar bloque, franja e intervalo horario exactos
                        var bloqueEsperado = ctx.Bloques.First(b =>
                            b.DiaSemanaNumero == planSesion.Ranura.Dia
                            && string.Equals(b.Franja, planSesion.Ranura.Franja, StringComparison.OrdinalIgnoreCase)
                            && Math.Abs(b.DuracionHoras - planSesion.Ranura.Horas) < 0.001);

                        stagingSesion.HorarioSesion.Should().Be(bloqueEsperado.ObtenerHorarioTexto());

                        // Comparar con el gen del individuo (IdBloqueHorario exacto, Día, Franja, Horas)
                        var genCorrespondiente = individuo.Genes.FirstOrDefault(g =>
                            g.IdAsignatura == planAsig.IdAsignatura
                            && g.DiaSemana == planSesion.Ranura.Dia);
                        genCorrespondiente.Should().NotBeNull();
                        genCorrespondiente.FechasGen.Select(d => d.Date).Should().Contain(stagingSesion.FechaSesion.Value.Date);
                        genCorrespondiente.IdAsignatura.Should().Be(planAsig.IdAsignatura);
                        genCorrespondiente.IdBloqueHorario.Should().Be(bloqueEsperado.Id);
                        genCorrespondiente.DiaSemana.Should().Be(planSesion.Ranura.Dia);
                        genCorrespondiente.HorasSemanales.Should().Be(planSesion.Ranura.Horas);
                        bloqueEsperado.Franja.Should().Be(planSesion.Ranura.Franja);
                    }

                    // Paridad con los genes del individuo tras Coherencia
                    var genesAsig = individuo.Genes.Where(g => g.IdAsignatura == planAsig.IdAsignatura).ToList();
                    genesAsig.Should().NotBeEmpty();
                    var fechasGenes = genesAsig.SelectMany(g => g.FechasGen).OrderBy(d => d).Select(d => d.Date).ToList();
                    var fechasEsperadas = planAsig.Sesiones.Select(s => s.Fecha.Date).OrderBy(d => d).ToList();
                    fechasGenes.Should().Equal(fechasEsperadas);
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(dbPath))
                    {
                        File.Delete(dbPath);
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// Criterio 12: Las transiciones de escalado de ventanas 4->5 y 4->5->6 no presentan regresiones,
        /// ofreciendo pasos incrementales inmediatos e idempotencia al alcanzar déficit cero.
        /// </summary>
        [Fact]
        public void Criterio12_EscaladoVentanas_SinRegresiones()
        {
            // 1. Escalado 4 -> 5 días (LUN-JUE a LUN-VIE) con AgEvaluadorDemandaVentana
            var ctx5W = CrearContexto(new int[] { 16, 16, 16, 16, 16 });
            ctx5W.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx5W.Ventana.FechaFin = new DateTime(2027, 2, 28);

            var antes = AgEvaluadorDemandaVentana.Evaluar("LUN-JUE", "TARDE", 4, ctx5W, ctx5W.Aulas[0], ctx5W.Ventana.FechaInicio, ctx5W.Ventana.FechaFin);
            var despues = AgEvaluadorDemandaVentana.Evaluar("LUN-VIE", "TARDE", 4, ctx5W, ctx5W.Aulas[0], ctx5W.Ventana.FechaInicio, ctx5W.Ventana.FechaFin);

            antes.DemandaPendiente.Should().Be(4);
            despues.CabeEnPlazo.Should().BeTrue();
            despues.DemandaPendiente.Should().Be(0);
            despues.Sesiones.Count.Should().Be(20);

            // 2. Escalera 4 -> 5 -> 6 (VIE-DOM -> JUE-DOM -> MIE-DOM) con AgEscaladorRanuraAula
            var ctx6W = CrearContexto(new int[] { 16, 16, 16, 16, 16, 16 });
            ctx6W.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx6W.Ventana.FechaFin = new DateTime(2027, 2, 28);
            ctx6W.Aulas[0].DiasVentana = "VIE-DOM";

            var escalador = new AgEscaladorRanuraAula(() => null);

            var paso1 = escalador.Evaluar(ctx6W).Single();
            paso1.Hacia.Should().Be("JUE-DOM");
            paso1.DemandaPendiente.Should().Be(4);

            ctx6W.Aulas[0].DiasVentana = paso1.Hacia;
            var paso2 = escalador.Evaluar(ctx6W).Single();
            paso2.Hacia.Should().Be("MIE-DOM");
            paso2.DemandaPendiente.Should().Be(0);

            ctx6W.Aulas[0].DiasVentana = paso2.Hacia;
            var paso3 = escalador.Evaluar(ctx6W);
            paso3.Should().BeEmpty(); // Idempotencia y tope alcanzado
        }

        /// <summary>
        /// Criterio 13: Materia con OrigenPatron = "AUTOMATICO" y patrón vacío (DiasPatron = "")
        /// planificada con más de dos ranuras (> 2 ranuras por carril).
        /// La fábrica conserva íntegramente las sesiones planificadas fechadas (no las trunca a 2 ranuras
        /// ni asigna sesionesPlanificadas = null) y persiste la totalidad de las sesiones a SQLite staging.
        /// </summary>
        [Fact]
        public async Task Criterio13_AutomaticoSinPatron_ConMasDeDosRanuras_ConservaPlanFechadoHastaStaging()
        {
            // Arrange: contexto con 1 aula multi-carril (JUE-DOM tiene cinco ranuras de 4 h, incluyendo domingo: JUE, VIE, SAB_M, SAB_T, DOM_M)
            var ctx = new AgContexto();
            ctx.IdPeriodo = 1;
            ctx.IdPrograma = 1;
            ctx.IdSede = 1;
            ctx.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 4), new DateTime(2027, 2, 28), AgOrigenVentana.Presencial).Value;

            var aula = new AgAula
            {
                Id = 1,
                IdPrograma = 1,
                IdSede = 1,
                DiasVentana = "JUE-DOM",
                Jornada = "NOCHE",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            };
            ctx.Aulas.Add(aula);

            ctx.Bloques.Add(new AgBloqueHorario { Id = 40, DiaSemanaNumero = 4, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00", DuracionHoras = 4.0, Activo = 1 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 50, DiaSemanaNumero = 5, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00", DuracionHoras = 4.0, Activo = 1 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 61, DiaSemanaNumero = 6, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 62, DiaSemanaNumero = 6, Franja = "TARDE", HoraInicio = "14:00", HoraFin = "18:00", DuracionHoras = 4.0, Activo = 1 });
            ctx.Bloques.Add(new AgBloqueHorario { Id = 70, DiaSemanaNumero = 7, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Activo = 1 });

            // 1 Asignatura de 32 horas (8 sesiones) con patrón vacío y origen AUTOMATICO
            ctx.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "MateriaHuecos", Horas = 32, IntHoraria = 32 });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 1,
                HorasObjetivo = 32,
                OrdenCurricular = 1,
                DiasPatron = string.Empty, // Patrón vacío
                OrigenPatron = "AUTOMATICO"
            });
            ctx.Docentes.Add(new Tercero { Id = 10, Nombre = "Docente Huecos", Activo = 1 });
            ctx.PoolDocentes.Add(new TerceroAsignatura { IdTercero = 10, IdAsignatura = 1 });
            ctx.ConstruirIndices();

            var planificador = new AgPlanificadorOleadas();
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (IdAsignatura: 1, Horas: 32, OrdenCurricular: 1)
            };
            var patrones = new Dictionary<int, string>
            {
                { 1, string.Empty }
            };

            // 1. Plan original: distribuye en las ranuras de JUE-DOM
            var planOriginalRes = planificador.Construir(
                demandas, aula.DiasVentana, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin,
                4.0, ctx.EsFestivo, ctx.Bloques, aula.Jornada, patrones, diagnostico: "Criterio13");
            planOriginalRes.IsSuccess.Should().BeTrue();
            var planAsig = planOriginalRes.Value.Planes.First(p => p.IdAsignatura == 1);
            planAsig.Sesiones.Should().HaveCount(8);
            var ranurasPlan = planAsig.Sesiones.Select(s => s.Ranura).Distinct().ToList();
            ranurasPlan.Count.Should().BeGreaterThan(2);

            // 2. Generar con fábrica
            var fabrica = new AgFabricaIndividuos(new Random(42), null, planificador);
            var individuo = fabrica.CrearAleatorio(ctx, 1, 0);
            individuo.Should().NotBeNull();

            // Sellar coherencia
            AgCoherenciaGrupo.Sellar(individuo, ctx);

            // 3. Persistir en SQLite Staging
            string dbPath = Path.Combine(Path.GetTempPath(), "AgStagingCriterio13_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbPath);
            string connStr = $"Data Source={dbPath};Version=3;";

            try
            {
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE ag_vinculaciones (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            id_corrida              INTEGER NOT NULL DEFAULT 0,
                            id_individuo            INTEGER NOT NULL DEFAULT 0,
                            idtercero               INTEGER NOT NULL DEFAULT 0,
                            idprograma              INTEGER NOT NULL DEFAULT 0,
                            idsede                  INTEGER NOT NULL DEFAULT 0,
                            idasignatura            INTEGER NOT NULL DEFAULT 0,
                            idperiodo               INTEGER NOT NULL DEFAULT 0,
                            semestre                TEXT    NOT NULL DEFAULT '',
                            grupo                   TEXT    NOT NULL DEFAULT '',
                            dia_horas_semanales     TEXT    NOT NULL DEFAULT '',
                            modalidadVinculacion    TEXT    NOT NULL DEFAULT 'Presencial',
                            total_horas             INTEGER NOT NULL DEFAULT 0,
                            fecha_inicio            TEXT    NOT NULL DEFAULT '',
                            fecha_fin               TEXT    NOT NULL DEFAULT '',
                            estado                  INTEGER NOT NULL DEFAULT 0,
                            isTerminalAereo         INTEGER NOT NULL DEFAULT 0,
                            idTerminales            INTEGER NOT NULL DEFAULT 0,
                            gastoTerminalAerea      REAL    NOT NULL DEFAULT 0.0,
                            es_vacante              INTEGER NOT NULL DEFAULT 0,
                            motivo_vacante          TEXT    NOT NULL DEFAULT '',
                            candidatos_pool         INTEGER NOT NULL DEFAULT 0,
                            subunidad               TEXT    NOT NULL DEFAULT ''
                        );
                        CREATE TABLE ag_vinculacion_gen (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            idagvinculacion         INTEGER NOT NULL,
                            idgen                   INTEGER NOT NULL,
                            subunidad               TEXT    DEFAULT ''
                        );
                        CREATE TABLE ag_sesiones (
                            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                            idvinculacion           INTEGER NOT NULL,
                            idtercero               INTEGER NOT NULL,
                            id_aula                 INTEGER NOT NULL DEFAULT 0,
                            fechasesion             TEXT    NOT NULL,
                            horariosesion           TEXT    NOT NULL DEFAULT '',
                            horas                   REAL    NOT NULL DEFAULT 0.0,
                            viatica                 INTEGER NOT NULL DEFAULT 0,
                            subunidad               TEXT    NOT NULL DEFAULT '',
                            idindividuo             INTEGER NULL,
                            idgen                   INTEGER NULL,
                            origen_ag               INTEGER NOT NULL DEFAULT 1,
                            estado_ag               TEXT    NOT NULL DEFAULT 'BORRADOR',
                            es_vacante              INTEGER NOT NULL DEFAULT 0,
                            motivo_vacante          TEXT    NOT NULL DEFAULT '',
                            candidatos_pool         INTEGER NOT NULL DEFAULT 0,
                            fecha_creacion          TEXT    NOT NULL DEFAULT (datetime('now'))
                        );
                    ");
                }

                var repoReal = new AgScheduleRepository(() => new SQLiteConnection(connStr));
                var expansor = new AgExpansorSesiones();
                var materializador = new AgMaterializadorSesiones(repoReal, expansor);

                var resMat = await materializador.MaterializarEnStagingAsync(100, individuo, ctx, CancellationToken.None);
                resMat.IsSuccess.Should().BeTrue(resMat.IsFailure ? resMat.Error : string.Empty);

                List<AgSesion> guardadasSes;
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    guardadasSes = (await conn.QueryAsync<AgSesion>(
                        "SELECT id as Id, idvinculacion as IdVinculacion, id_aula as IdAula, fechasesion as FechaSesion, horariosesion as HorarioSesion, horas as Horas, idindividuo as IdIndividuo, idgen as IdGen FROM ag_sesiones WHERE idindividuo = @IdIndividuo",
                        new { IdIndividuo = individuo.Id })).ToList();
                }

                // Assert: Las 8 sesiones están presentes en SQLite sin haber sido descartadas por la fábrica
                guardadasSes.Should().HaveCount(8);
                guardadasSes.Select(s => s.FechaSesion.Value.Date).OrderBy(d => d).Should().Equal(
                    planAsig.Sesiones.Select(s => s.Fecha.Date).OrderBy(d => d));
            }
            finally
            {
                try { if (File.Exists(dbPath)) File.Delete(dbPath); } catch { }
            }
        }

        /// <summary>
        /// Criterio 14: Validación integral de hallazgos de Corrida 94 (Fase 1814):
        /// 1. Aula aprobada con déficit es bloqueada en ValidarAprobadas sin mutar estado.
        /// 2. Asignaturas con semanas intercaladas sin solape horario arrojan 0 HARD en AgValidadorReglas.
        /// 3. AgScheduleRepository persiste vacantes en ag_corrida y resuelve indicesGenes a id_gen.
        /// </summary>
        [Fact]
        public async Task Criterio14_Corrida94_CalendarioConflictos_Aprobadas_Y_ConsistenciaStaging()
        {
            // 1. ValidarAprobadas pre-vuelo
            var ctxAprobadas = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2027, 2, 1),
                    FechaFin = new DateTime(2027, 5, 30)
                }
            };
            ctxAprobadas.Programas.Add(new Programa { Id = 1, Nombre = "Prog Test", Modalidad = "PREGRADO" });
            var aulaDeficit = new AgAula
            {
                Id = 13,
                AulaRef = "CALI I-A",
                Estado = AgEstadoAula.Aprobado,
                DiasVentana = "5,6",
                Jornada = "NOCHE",
                HorasDia = 4,
                IdPrograma = 1
            };
            ctxAprobadas.Aulas.Add(aulaDeficit);

            for (int i = 1; i <= 4; i++)
            {
                ctxAprobadas.Asignaturas.Add(new Asignatura { Id = i, Nombre = "Materia " + i, IntHoraria = 64, IdPrograma = 1 });
                var aa = new AgAulaAsignatura
                {
                    IdAula = 13,
                    IdAsignatura = i,
                    HorasObjetivo = 64.0
                };
                ctxAprobadas.AulaAsignaturas.Add(aa);
                ctxAprobadas.AulaAsignaturaPorAulaAsig[(13, i)] = aa;
            }
            ctxAprobadas.ConstruirIndices();

            var resAprobadas = AgGeneradorHorarios.ValidarAprobadas(ctxAprobadas);
            resAprobadas.IsFailure.Should().BeTrue();
            resAprobadas.Error.Should().Contain("I5_AULA_APROBADA_CON_DEFICIT");

            // 2. AgCalendarioConflictos & AgValidadorReglas: Materias intercaladas sin cruce real
            var ctxReglas = new AgContexto
            {
                IdPrograma = 1,
                Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 4, 1), new DateTime(2027, 5, 30), AgOrigenVentana.Presencial).Value
            };
            ctxReglas.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            ctxReglas.AulasPorId[36] = new AgAula { Id = 36, IdPrograma = 1, Semestre = "III", Grupo = "A", Estado = AgEstadoAula.Disponible };
            ctxReglas.RegistrarDocenteHabilitado(10, 1, 201, "III");
            ctxReglas.RegistrarDocenteHabilitado(20, 1, 202, "III");

            var bloqueViernes = new AgBloqueHorario
            {
                Id = 50,
                DiaSemanaNumero = 5,
                Franja = "TARDE",
                HoraInicio = "14:00",
                HoraFin = "18:00",
                DuracionHoras = 4.0,
                Activo = 1
            };
            ctxReglas.Bloques.Add(bloqueViernes);
            ctxReglas.BloquesPorId[50] = bloqueViernes;
            ctxReglas.ConstruirIndices();

            var genEstadistica = new AgGen
            {
                Id = 1,
                IdAula = 36,
                IdAsignatura = 201,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "III",
                Grupo = "A",
                DiasPatron = "CARRIL_1",
                IdTercero = 10,
                IdBloqueHorario = 50,
                DiaSemana = 5,
                HorasGen = 8.0,
                HorasSemanales = 4.0,
                FechaInicioEstimada = new DateTime(2027, 4, 2),
                FechaFinEstimada = new DateTime(2027, 4, 16),
                EsVacante = 0,
                FechasGen = new List<DateTime>
                {
                    new DateTime(2027, 4, 2),  // Semana 1 (Viernes)
                    new DateTime(2027, 4, 16)  // Semana 3 (Viernes)
                }
            };

            var genTeorias = new AgGen
            {
                Id = 2,
                IdAula = 36,
                IdAsignatura = 202,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "III",
                Grupo = "A",
                DiasPatron = "CARRIL_2",
                IdTercero = 20,
                IdBloqueHorario = 50,
                DiaSemana = 5,
                HorasGen = 8.0,
                HorasSemanales = 4.0,
                FechaInicioEstimada = new DateTime(2027, 4, 9),
                FechaFinEstimada = new DateTime(2027, 4, 23),
                EsVacante = 0,
                FechasGen = new List<DateTime>
                {
                    new DateTime(2027, 4, 9),  // Semana 2 (Viernes)
                    new DateTime(2027, 4, 23)  // Semana 4 (Viernes)
                }
            };

            var ind = new AgIndividuo();
            ind.Genes.Add(genEstadistica);
            ind.Genes.Add(genTeorias);

            var validador = new AgValidadorReglas();
            validador.Evaluar(ind, ctxReglas);

            ind.HardViolations.Should().Be(0, "Estadística y Teorías usan semanas alternadas en el mismo bloque y aula sin cruce real");
            ind.Estado.Should().Be(AgEstadoIndividuo.Factible);

            // 3. Persistencia de vacantes y resolución de indicesGenes a id_gen
            string dbPath = Path.Combine(Path.GetTempPath(), "esap_criterio14_" + Guid.NewGuid().ToString("N") + ".db");
            string connStr = "Data Source=" + dbPath + ";Version=3;";

            try
            {
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE ag_corrida (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            idperiodo INTEGER NOT NULL DEFAULT 0,
                            idprograma INTEGER NOT NULL DEFAULT 0,
                            idaula INTEGER NOT NULL DEFAULT 0,
                            fecha_inicio_corrida TEXT NOT NULL DEFAULT '',
                            fecha_fin_corrida TEXT NOT NULL DEFAULT '',
                            estado INTEGER NOT NULL DEFAULT 0,
                            poblacion_tamano INTEGER NOT NULL DEFAULT 0,
                            generaciones_max INTEGER NOT NULL DEFAULT 0,
                            generaciones_ejecutadas INTEGER NOT NULL DEFAULT 0,
                            tasa_cruce REAL NOT NULL DEFAULT 0,
                            tasa_mutacion REAL NOT NULL DEFAULT 0,
                            elitismo INTEGER NOT NULL DEFAULT 0,
                            mejor_fitness REAL NOT NULL DEFAULT 0,
                            hard_violations INTEGER NOT NULL DEFAULT 0,
                            soft_violations REAL NOT NULL DEFAULT 0,
                            vacantes INTEGER NOT NULL DEFAULT 0,
                            id_individuo_optimo INTEGER NULL,
                            parametros_json TEXT NOT NULL DEFAULT '{}',
                            subunidad TEXT NOT NULL DEFAULT ''
                        );
                        CREATE TABLE ag_individuo (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            id_corrida INTEGER NOT NULL,
                            generacion INTEGER NOT NULL DEFAULT 0,
                            fitness REAL NOT NULL DEFAULT 0,
                            hard_violations INTEGER NOT NULL DEFAULT 0,
                            soft_violations REAL NOT NULL DEFAULT 0,
                            estado INTEGER NOT NULL DEFAULT 0,
                            subunidad TEXT NOT NULL DEFAULT ''
                        );
                        CREATE TABLE ag_gen (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            id_individuo INTEGER NOT NULL,
                            idperiodo INTEGER NOT NULL DEFAULT 0,
                            idasignatura INTEGER NOT NULL DEFAULT 0,
                            idprograma INTEGER NOT NULL DEFAULT 0,
                            idsede INTEGER NOT NULL DEFAULT 0,
                            semestre TEXT NOT NULL DEFAULT '',
                            grupo TEXT NOT NULL DEFAULT '',
                            idtercero INTEGER NOT NULL DEFAULT 0,
                            id_aula INTEGER NOT NULL DEFAULT 0,
                            id_bloque_horario INTEGER NOT NULL DEFAULT 0,
                            dia_semana INTEGER NOT NULL DEFAULT 0,
                            horas_semanales REAL NOT NULL DEFAULT 0,
                            total_horas_asignatura REAL NOT NULL DEFAULT 0,
                            fecha_inicio_estimada TEXT NOT NULL DEFAULT '',
                            fecha_fin_estimada TEXT NOT NULL DEFAULT '',
                            es_fijo INTEGER NOT NULL DEFAULT 0,
                            es_vacante INTEGER NOT NULL DEFAULT 0,
                            motivo_vacante TEXT NOT NULL DEFAULT '',
                            dentro_de_ventana INTEGER NOT NULL DEFAULT 0,
                            estado INTEGER NOT NULL DEFAULT 0,
                            subunidad TEXT NOT NULL DEFAULT '',
                            indice_dia_patron INTEGER NOT NULL DEFAULT 0,
                            patron_semanal TEXT NOT NULL DEFAULT '',
                            dias_patron TEXT NOT NULL DEFAULT '',
                            clave_grupo TEXT NOT NULL DEFAULT '',
                            horas_gen REAL NOT NULL DEFAULT 0,
                            fechas_gen TEXT NOT NULL DEFAULT ''
                        );
                        CREATE TABLE ag_violacion (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            id_individuo INTEGER NOT NULL,
                            id_gen INTEGER NULL,
                            regla TEXT NOT NULL DEFAULT '',
                            es_dura INTEGER NOT NULL DEFAULT 1,
                            mensaje TEXT NOT NULL DEFAULT '',
                            detalle_json TEXT NOT NULL DEFAULT '{}',
                            subunidad TEXT NOT NULL DEFAULT ''
                        );
                    ");
                    conn.Execute("INSERT INTO ag_corrida (id, estado, vacantes) VALUES (1, 1, 0);");
                }

                var repo = new AgScheduleRepository(() => new SQLiteConnection(connStr));

                // Prueba UpdateCorridaAsync persistiendo vacantes
                var corrida = (await repo.GetCorridaByIdAsync(1, CancellationToken.None)).Value;
                corrida.Vacantes = 7;
                corrida.HardViolations = 0;
                corrida.Estado = AgEstadoCorrida.Completada;
                var resUpd = await repo.UpdateCorridaAsync(corrida, CancellationToken.None);
                resUpd.IsSuccess.Should().BeTrue();

                var corridaLeida = (await repo.GetCorridaByIdAsync(1, CancellationToken.None)).Value;
                corridaLeida.Vacantes.Should().Be(7);

                // Prueba SaveIndividuoAsync mapeando indicesGenes a id_gen
                var indPersist = new AgIndividuo { IdCorrida = 1, Generacion = 1, Fitness = 0.95 };
                indPersist.Genes.Add(new AgGen { IdAsignatura = 10, HorasGen = 4.0, IdBloqueHorario = 1 });
                indPersist.Genes.Add(new AgGen { IdAsignatura = 20, HorasGen = 4.0, IdBloqueHorario = 1 });

                var violacionConIndices = new AgViolacion(AgReglaViolacion.R2_SuperposicionAula, true, "Cruce de aula",
                    detalleJson: "{\"indicesGenes\":[1],\"fecha\":\"2027-04-16\"}");
                indPersist.Violaciones.Add(violacionConIndices);

                var saveRes = await repo.SaveIndividuoAsync(indPersist, CancellationToken.None);
                saveRes.IsSuccess.Should().BeTrue();

                // Verificar que el gen 1 obtuvo un ID y la violacion fue enlazada a él
                long idGenEsperado = indPersist.Genes[1].Id;
                idGenEsperado.Should().BeGreaterThan(0);
                violacionConIndices.IdGen.Should().Be(idGenEsperado);
                violacionConIndices.DetalleJson.Should().Contain("\"idGenes\":[" + idGenEsperado + "]");
            }
            finally
            {
                try { if (File.Exists(dbPath)) File.Delete(dbPath); } catch { }
            }
        }

        #region Métodos Auxiliares

        private static List<AgBloqueHorario> CrearBloquesLunesViernes(double duracion, string franja)
        {
            var bloques = new List<AgBloqueHorario>();
            for (int dia = 1; dia <= 5; dia++)
            {
                bloques.Add(new AgBloqueHorario
                {
                    Id = dia * 10,
                    DiaSemanaNumero = dia,
                    DuracionHoras = duracion,
                    Franja = franja,
                    HoraInicio = "14:00",
                    HoraFin = "18:00",
                    Activo = 1
                });
            }
            return bloques;
        }

        private static AgPlanOleada CrearPlan(int idAsignatura, List<AgSesionPlanificada> sesiones, AgRanura ranura, DateTime ini, DateTime fin, int requeridas)
        {
            return new AgPlanOleada(
                idAsignatura,
                0,
                0,
                new List<int> { ranura.Dia },
                ini,
                fin,
                1,
                requeridas,
                10,
                new List<AgRanura> { ranura },
                sesiones);
        }

        private static AgContexto CrearContexto(int[] horas)
        {
            var ctx = new AgContexto
            {
                IdSede = 1,
                Ventana = new VentanaSemestre
                {
                    FechaInicio = new DateTime(2027, 2, 5),
                    FechaFin = new DateTime(2027, 5, 30)
                }
            };
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PRESENCIAL" });
            ctx.Aulas.Add(new AgAula
            {
                Id = 1,
                IdSede = 1,
                IdPrograma = 1,
                AulaRef = "CALI-TEST",
                Semestre = "III",
                Jornada = "TARDE",
                HorasDia = 4,
                DiasVentana = "LUN-VIE",
                Estado = AgEstadoAula.Disponible
            });
            for (int dia = 1; dia <= 7; dia++)
            {
                ctx.Bloques.Add(new AgBloqueHorario
                {
                    Id = dia * 10,
                    DiaSemanaNumero = dia,
                    DuracionHoras = 4,
                    Franja = dia >= 6 ? "MANANA" : "TARDE",
                    HoraInicio = dia >= 6 ? "08:00" : "14:00",
                    HoraFin = dia >= 6 ? "12:00" : "18:00",
                    Activo = 1
                });
                if (dia == 6)
                {
                    ctx.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 61,
                        DiaSemanaNumero = 6,
                        DuracionHoras = 4,
                        Franja = "TARDE",
                        HoraInicio = "13:00",
                        HoraFin = "17:00",
                        Activo = 1
                    });
                }
            }
            for (int i = 0; i < horas.Length; i++)
            {
                ctx.Asignaturas.Add(new Asignatura { Id = i + 1, IdPrograma = 1, Semestre = "III", Horas = horas[i] });
                ctx.AulaAsignaturas.Add(new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i + 1,
                    HorasObjetivo = horas[i],
                    OrdenCurricular = i + 1,
                    OrigenPatron = "AUTOMATICO",
                    DiasPatron = "5"
                });
            }
            ctx.Festivos.Add((new DateTime(2027, 3, 22), new DateTime(2027, 3, 26)));
            ctx.ConstruirIndices();
            return ctx;
        }

        #endregion
    }
}

