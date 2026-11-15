using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Pruebas de regresión pura para la planificación de ranuras y oleadas mixtas (Fase 1798).
    /// Cubre escenarios de Cali III, Cali IV y Cartago X-A con asignaturas restringidas por patrón y materias automáticas.
    /// </summary>
    public class AgPlanificadorRanurasRegresionTests
    {
        private static readonly Func<DateTime, bool> EsFestivoCanonico = d =>
            (d.Date >= new DateTime(2027, 3, 22) && d.Date <= new DateTime(2027, 3, 28))
            || d.Date == new DateTime(2027, 5, 10);

        [Theory]
        [InlineData(false, 13)] // Cali III: 32, 48, 48, 48, 32, 48 h -> límite 13/05/2027
        [InlineData(true, 19)]  // Cali IV:  32, 48, 48, 48, 48, 48 h -> límite 19/05/2027
        public void Mixta_ConservaRestriccionesYCompletaAntesDelLimite(bool cuarto, int diaFin)
        {
            // Arrange
            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);
            var fLimite = new DateTime(2027, 5, diaFin);

            var bloques = new List<AgBloqueHorario>();
            for (int dia = 1; dia <= 5; dia++)
            {
                bloques.Add(new AgBloqueHorario
                {
                    Id = dia * 10,
                    DiaSemanaNumero = dia,
                    DuracionHoras = 4.0,
                    Franja = "TARDE",
                    HoraInicio = "14:00",
                    HoraFin = "18:00",
                    Activo = 1
                });
            }

            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 48.0, 2),
                (3, 48.0, 3),
                (4, 48.0, 4),
                (5, cuarto ? 48.0 : 32.0, 5),
                (6, 48.0, 6)
            };

            var patrones = new Dictionary<int, string>
            {
                { 1, "1,2" },
                { 2, "3,4" },
                { 3, "5" },
                { 4, "1,2" },
                { 5, "3,4" }
                // Asignatura 6: AUTOMATICO (sin entrada en patrones)
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resConstruir = sut.Construir(
                demandas, "LUN-VIE", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "TARDE", patrones);

            var resEvaluacion = sut.ConstruirEvaluacion(
                demandas, "LUN-VIE", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "TARDE", patrones);

            // Assert
            resConstruir.IsSuccess.Should().BeTrue();
            resConstruir.Value.Infactibles.Should().BeEmpty();

            resEvaluacion.IsSuccess.Should().BeTrue();
            resEvaluacion.Value.Infactibles.Should().BeEmpty();

            int sesionesEsperadas = cuarto ? 68 : 64;
            var sesionesConstruir = resConstruir.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            var sesionesEval = resEvaluacion.Value.Planes.SelectMany(p => p.Sesiones).ToList();

            sesionesConstruir.Should().HaveCount(sesionesEsperadas);
            sesionesEval.Should().HaveCount(sesionesEsperadas);

            // Paridad estricta entre Construir y ConstruirEvaluacion
            var llavesConstruir = sesionesConstruir
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia)
                .OrderBy(x => x)
                .ToList();
            var llavesEval = sesionesEval
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia)
                .OrderBy(x => x)
                .ToList();
            llavesConstruir.Should().Equal(llavesEval);

            // Completitud antes del límite estipulado
            var sesionesAsig6 = sesionesConstruir.Where(s => s.IdAsignatura == 6).ToList();
            sesionesAsig6.Should().NotBeEmpty();
            sesionesAsig6.Max(s => s.Fecha).Should().BeOnOrBefore(fLimite);
            sesionesConstruir.Max(s => s.Fecha).Should().BeOnOrBefore(fLimite);

            // Preservación estricta de restricciones de días según patrón
            sesionesConstruir.Where(s => s.IdAsignatura == 1).Should().OnlyContain(s => s.Ranura.Dia == 1 || s.Ranura.Dia == 2);
            sesionesConstruir.Where(s => s.IdAsignatura == 2).Should().OnlyContain(s => s.Ranura.Dia == 3 || s.Ranura.Dia == 4);
            sesionesConstruir.Where(s => s.IdAsignatura == 3).Should().OnlyContain(s => s.Ranura.Dia == 5);
            sesionesConstruir.Where(s => s.IdAsignatura == 4).Should().OnlyContain(s => s.Ranura.Dia == 1 || s.Ranura.Dia == 2);
            sesionesConstruir.Where(s => s.IdAsignatura == 5).Should().OnlyContain(s => s.Ranura.Dia == 3 || s.Ranura.Dia == 4);
        }

        [Fact]
        public void Construir_WhenCartagoXA_ShouldCompletarSinInfactiblesYRespetarPatrones()
        {
            // Arrange
            // Cartago X-A: 32, 32, 64, 112 h en ventana JUE-DOM (período 2027-02-05 a 2027-05-30)
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
                (4, 112.0, 4)
            };

            var patrones = new Dictionary<int, string>
            {
                { 1, "4,5" },
                { 2, "6" },
                { 3, "4,5" }
                // Asignatura 4: AUTOMATICO (sin entrada en patrones)
            };

            var sut = new AgPlanificadorOleadas();

            // Act
            var resConstruir = sut.Construir(
                demandas, "JUE-DOM", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "NOCHE", patrones);

            var resEvaluacion = sut.ConstruirEvaluacion(
                demandas, "JUE-DOM", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "NOCHE", patrones);

            // Assert
            resConstruir.IsSuccess.Should().BeTrue();
            resConstruir.Value.Infactibles.Should().BeEmpty();

            resEvaluacion.IsSuccess.Should().BeTrue();
            resEvaluacion.Value.Infactibles.Should().BeEmpty();

            // Demanda total = 240 horas = 60 sesiones de 4 horas
            var sesionesConstruir = resConstruir.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            var sesionesEval = resEvaluacion.Value.Planes.SelectMany(p => p.Sesiones).ToList();

            sesionesConstruir.Should().HaveCount(60);
            sesionesEval.Should().HaveCount(60);

            // Sin colisiones horarias en el calendario
            sesionesConstruir.Select(s => s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .Should().OnlyHaveUniqueItems();

            // Respeto de patrones en las asignaturas restringidas
            sesionesConstruir.Where(s => s.IdAsignatura == 1).Should().OnlyContain(s => s.Ranura.Dia == 4 || s.Ranura.Dia == 5);
            sesionesConstruir.Where(s => s.IdAsignatura == 2).Should().OnlyContain(s => s.Ranura.Dia == 6);
            sesionesConstruir.Where(s => s.IdAsignatura == 3).Should().OnlyContain(s => s.Ranura.Dia == 4 || s.Ranura.Dia == 5);

            // Asignatura 4 completa todas sus 28 sesiones dentro del plazo
            var sesionesAsig4 = sesionesConstruir.Where(s => s.IdAsignatura == 4).ToList();
            sesionesAsig4.Should().HaveCount(28);
            sesionesConstruir.Max(s => s.Fecha).Should().BeOnOrBefore(fFin);

            // Paridad entre Construir y ConstruirEvaluacion
            var llavesConstruir = sesionesConstruir
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .OrderBy(x => x)
                .ToList();
            var llavesEval = sesionesEval
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Dia + "|" + s.Ranura.Franja)
                .OrderBy(x => x)
                .ToList();
            llavesConstruir.Should().Equal(llavesEval);
        }

        [Fact]
        public void CompararSesionesRestringidas_AntesYDespues_ConservaDistribucionOriginal()
        {
            // Arrange
            // Caso Cartago X-A con 3 asignaturas restringidas iniciales ("Antes")
            // y posterior incorporación de la 4ta asignatura AUTOMATICO ("Después")
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

            var demandasAntes = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 64.0, 3)
            };

            var demandasDespues = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1),
                (2, 32.0, 2),
                (3, 64.0, 3),
                (4, 112.0, 4)
            };

            var patrones = new Dictionary<int, string>
            {
                { 1, "4,5" },
                { 2, "6" },
                { 3, "4,5" }
            };

            var sut = new AgPlanificadorOleadas();

            // Act - "Antes": Planificación de solo las materias restringidas
            var resAntes = sut.Construir(
                demandasAntes, "JUE-DOM", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "NOCHE", patrones);

            // Act - "Después": Plan nuevo por prioridad de horas, sin fechas implícitas de otra llamada.
            var resDespues = sut.Construir(
                demandasDespues, "JUE-DOM", fInicio, fFin, 4.0,
                EsFestivoCanonico, bloques, "NOCHE", patrones);

            // Assert
            resAntes.IsSuccess.Should().BeTrue();
            resAntes.Value.Infactibles.Should().BeEmpty();

            resDespues.IsSuccess.Should().BeTrue();
            resDespues.Value.Infactibles.Should().BeEmpty();

            var sesionesAntes = resAntes.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            var sesionesDespues = resDespues.Value.Planes.SelectMany(p => p.Sesiones).ToList();

            // Cada materia restringida (1, 2, 3) debe mantener su cuota completa de sesiones
            sesionesDespues.Where(s => s.IdAsignatura == 1).Should().HaveCount(8);
            sesionesDespues.Where(s => s.IdAsignatura == 2).Should().HaveCount(8);
            sesionesDespues.Where(s => s.IdAsignatura == 3).Should().HaveCount(16);

            // Todas las sesiones de las materias restringidas en "Después" deben conservar estrictamente sus días válidos
            sesionesDespues.Where(s => s.IdAsignatura == 1).Should().OnlyContain(s => s.Ranura.Dia == 4 || s.Ranura.Dia == 5);
            sesionesDespues.Where(s => s.IdAsignatura == 2).Should().OnlyContain(s => s.Ranura.Dia == 6);
            sesionesDespues.Where(s => s.IdAsignatura == 3).Should().OnlyContain(s => s.Ranura.Dia == 4 || s.Ranura.Dia == 5);

            sesionesDespues.Where(s => s.IdAsignatura == 4).Should().HaveCount(28);
            sesionesDespues.Should().HaveCount(60);

            // El completado recibe explícitamente "Antes": ahora sí debe preservar sus fechas exactas.
            var ranuras = AgRanuras.PorVentana("JUE-DOM", bloques, "NOCHE", 4.0);
            // La demanda ampliada declara pendientes solo las materias ausentes de la referencia previa.
            var antes = resAntes.Value;
            var nuevasPendientes = demandasDespues.Select(d => d.IdAsignatura)
                .Except(antes.Planes.Select(p => p.IdAsignatura).Union(antes.Infactibles));
            var referenciaAmpliada = new AgPlanOleadasResultado(antes.Planes,
                antes.Infactibles.Union(nuevasPendientes).ToList(), antes.DetalleInfactibles,
                antes.Particion, antes.VentanaSugerida, antes.PatronesRenormalizados,
                antes.OcupacionPorcentaje, antes.BusquedaAgotada);
            var completado = AgPlanificadorRanuras.CompletarReferencia(demandasDespues, ranuras,
                fInicio, fFin, 4.0, EsFestivoCanonico, patrones, referenciaAmpliada,
                System.Threading.CancellationToken.None);
            completado.IsSuccess.Should().BeTrue(completado.IsFailure ? completado.Error : string.Empty);
            completado.Value.Infactibles.Should().BeEmpty();
            var sesionesCompletadas = completado.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            sesionesCompletadas.Should().HaveCount(60);
            sesionesCompletadas.Where(s => s.IdAsignatura == 4).Should().HaveCount(28);

            var fechasAntes1 = sesionesAntes.Where(s => s.IdAsignatura == 1).Select(s => s.Fecha).OrderBy(d => d).ToList();
            var fechasDespues1 = sesionesCompletadas.Where(s => s.IdAsignatura == 1).Select(s => s.Fecha).OrderBy(d => d).ToList();
            fechasDespues1.Should().Equal(fechasAntes1);

            var fechasAntes2 = sesionesAntes.Where(s => s.IdAsignatura == 2).Select(s => s.Fecha).OrderBy(d => d).ToList();
            var fechasDespues2 = sesionesCompletadas.Where(s => s.IdAsignatura == 2).Select(s => s.Fecha).OrderBy(d => d).ToList();
            fechasDespues2.Should().Equal(fechasAntes2);

            var fechasAntes3 = sesionesAntes.Where(s => s.IdAsignatura == 3).Select(s => s.Fecha).OrderBy(d => d).ToList();
            var fechasDespues3 = sesionesCompletadas.Where(s => s.IdAsignatura == 3).Select(s => s.Fecha).OrderBy(d => d).ToList();
            fechasDespues3.Should().Equal(fechasAntes3);
        }
    }
}
