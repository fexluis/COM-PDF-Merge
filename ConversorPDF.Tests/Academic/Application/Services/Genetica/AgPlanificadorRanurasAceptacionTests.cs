using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using CSharpFunctionalExtensions;
using FluentAssertions;
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
        /// (13/05 para III, 19/05 para IV), cero sesiones posteriores al 30/05/2027 y conservan las
        /// fechas exactas de las 5 materias restringidas previas.
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

            // 4. Fechas de las cinco materias restringidas conservadas por el completado
            for (int asigId = 1; asigId <= 5; asigId++)
            {
                var fechasPrevias = resPrevio.Value.Planes.Single(p => p.IdAsignatura == asigId).Sesiones.Select(s => s.Fecha).OrderBy(d => d).ToList();
                var fechasMixtas = resMixto.Value.Planes.Single(p => p.IdAsignatura == asigId).Sesiones.Select(s => s.Fecha).OrderBy(d => d).ToList();
                fechasMixtas.Should().Equal(fechasPrevias);
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
        /// Criterio 11: Fidelidad de fechas entre la fábrica de individuos (FechasGen) y el expansor
        /// de sesiones (AgExpansorSesiones), garantizando que no existan mutaciones espurias ni truncamientos.
        /// </summary>
        [Fact]
        public void Criterio11_FidelidadFechas_FabricaYExpansor()
        {
            // Arrange
            var ctx = CrearContexto(new int[] { 32, 48, 48 });
            var planificador = new AgPlanificadorOleadas();
            var fabrica = new AgFabricaIndividuos(new Random(42), null, planificador);
            var expansor = new AgExpansorSesiones();

            // Act
            var individuo = fabrica.CrearAleatorio(ctx, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            var genesActivos = individuo.Genes.Where(g => g.HorasGen > 0.0).ToList();
            genesActivos.Should().NotBeEmpty();

            foreach (var gen in genesActivos)
            {
                gen.FechasGen.Should().NotBeNullOrEmpty();
                double unidad = ctx.BloquesPorId[gen.IdBloqueHorario].DuracionHoras;
                int sesionesEsperadas = (int)Math.Round(gen.HorasGen / unidad);
                gen.FechasGen.Count.Should().Be(sesionesEsperadas);

                // Expansión a sesiones cronológicas reales
                var resExp = expansor.Expandir(gen, 1, ctx);
                resExp.IsSuccess.Should().BeTrue();

                // Las fechas de las sesiones coinciden exactamente con FechasGen
                var fechasExpandidas = resExp.Value.Select(s => s.FechaSesion.Value).OrderBy(d => d).ToList();
                var fechasGenOrdenadas = gen.FechasGen.OrderBy(d => d).ToList();
                fechasExpandidas.Should().Equal(fechasGenOrdenadas);

                // La suma de duraciones coincide exactamente con HorasGen
                resExp.Value.Sum(s => s.Horas).Should().Be(gen.HorasGen);
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

