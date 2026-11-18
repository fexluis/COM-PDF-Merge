using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Implementación de la expansión de genes cromosómicos a sesiones cronológicas individuales.
    /// Garantiza la exclusión de festivos, acotación en la ventana semestral y cálculo de viáticos.
    /// </summary>
    public class AgExpansorSesiones : IAgExpansorSesiones
    {
        private readonly AgPlanificadorTemporal _planificador;

        public AgExpansorSesiones(AgPlanificadorTemporal planificador = null)
        {
            _planificador = planificador ?? new AgPlanificadorTemporal();
        }

        public static Result<List<DateTime>> ProyectarFechas(AgGen gen, AgContexto contexto, AgPlanificadorTemporal planificador = null)
        {
            if (gen == null) throw new ArgumentNullException(nameof(gen));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            // Fase 573 & 576: sin carga => 0 sesiones directamente, sin fallar por ventana/bloque
            if (gen.HorasGen <= 0.0)
            {
                return Result.Success(new List<DateTime>());
            }

            if (contexto.Ventana == null)
            {
                return Result.Failure<List<DateTime>>("Ventana semestral no resuelta.");
            }

            if (contexto.BloquesPorId == null || !contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bloque))
            {
                return Result.Failure<List<DateTime>>($"Bloque horario {gen.IdBloqueHorario} no encontrado.");
            }

            DateTime fInicio = gen.FechaInicioEstimada.HasValue
                ? gen.FechaInicioEstimada.Value.Date
                : contexto.Ventana.FechaInicio.Date;
            DateTime fFin = gen.FechaFinEstimada.HasValue
                ? gen.FechaFinEstimada.Value.Date
                : contexto.Ventana.FechaFin.Date;

            // Corrida 92 (T6): Cuando usaFechasGen es true, el rango efectivo de validación fInicio..fFin
            // debe contener al menos los extremos de FechasGen, siempre dentro de contexto.Ventana,
            // para que una sub-ventana estrecha no rechace sesiones legítimas ubicadas por el planificador fechado en huecos.
            bool usaFechasGen = gen.FechasGen != null && gen.FechasGen.Count > 0;
            if (usaFechasGen)
            {
                DateTime minFecha = gen.FechasGen.Min(d => d.Date);
                DateTime maxFecha = gen.FechasGen.Max(d => d.Date);

                if (fInicio > minFecha)
                {
                    fInicio = minFecha;
                }
                if (fFin < maxFecha)
                {
                    fFin = maxFecha;
                }

                if (fInicio < contexto.Ventana.FechaInicio.Date)
                {
                    fInicio = contexto.Ventana.FechaInicio.Date;
                }
                if (fFin > contexto.Ventana.FechaFin.Date)
                {
                    fFin = contexto.Ventana.FechaFin.Date;
                }
            }

            if (!contexto.Ventana.ContieneRango(fInicio, fFin))
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id} fuera de ventana: {fInicio:yyyy-MM-dd}..{fFin:yyyy-MM-dd}.");
            }

            double horasObjetivoGen = gen.HorasGen;

            // Fase 964: Cuantización estricta de sesiones (sin residuos fraccionarios Math.Min)
            double unidad = bloque.DuracionHoras;
            if (unidad <= 0)
            {
                return Result.Failure<List<DateTime>>($"Gen {gen.Id}: bloque con duración inválida ({unidad} h).");
            }

            // Ronda 206 / Fase 1056 (D4/I6): Defensa en profundidad de duración esperada según nivel académico del aula/programa
            AgAula aula = null;
            if (contexto.AulasPorId != null) contexto.AulasPorId.TryGetValue(gen.IdAula, out aula);
            if (aula != null)
            {
                int idProg = gen.IdPrograma > 0 ? gen.IdPrograma : (aula.IdPrograma > 0 ? aula.IdPrograma : contexto.IdPrograma);
                Programa progGen = null;
                if (contexto.ProgramasPorId != null && contexto.ProgramasPorId.TryGetValue(idProg, out var pg))
                {
                    progGen = pg;
                }
                else
                {
                    progGen = contexto.Programas?.FirstOrDefault(p => p.Id == idProg);
                }
                var nivel = AgNivelAcademicoCatalogo.Desde(progGen?.Modalidad);
                double duracionEsperada = AgUnidadesSemanales.HorasSesionAula(aula, nivel);
                if (Math.Abs(unidad - duracionEsperada) > 0.001)
                {
                    return Result.Failure<List<DateTime>>(
                        $"Gen {gen.Id}: bloque con duración ({unidad} h) no coincide con la esperada ({duracionEsperada} h) para nivel {nivel}.");
                }
            }

            double sesionesExactas = horasObjetivoGen / unidad;
            if (Math.Abs(sesionesExactas - Math.Round(sesionesExactas)) > 0.001)
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: {horasObjetivoGen:0.#} h no es múltiplo de {unidad:0.#} h (unidad de sesión).");
            }
            int sesionesRequeridas = (int)Math.Round(sesionesExactas);

            // Fase 414: Usar FechasGen si fueron precalculadas por el planificador semanal
            List<DateTime> fechas = usaFechasGen
                ? gen.FechasGen
                : (planificador ?? new AgPlanificadorTemporal()).GenerarFechasValidas(
                    bloque.DiaSemana,
                    fInicio,
                    fFin,
                    bloque.DuracionHoras,
                    horasObjetivoGen,
                    contexto.EsFestivo);

            if (fechas == null || fechas.Count == 0)
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: sin fechas disponibles para cubrir {horasObjetivoGen:0.#} h.");
            }

            // Validaciones estrictas antes de materializar:
            // 1. Plazo semestral
            if (fechas.Any(f => f.Date < contexto.Ventana.FechaInicio.Date || f.Date > contexto.Ventana.FechaFin.Date))
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id} contiene sesiones fuera del plazo semestral ({contexto.Ventana.FechaInicio:yyyy-MM-dd}..{contexto.Ventana.FechaFin:yyyy-MM-dd}).");
            }

            // 2. Rango del gen
            if (fechas.Any(f => f.Date < fInicio || f.Date > fFin))
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id} contiene sesiones fuera de su rango ({fInicio:yyyy-MM-dd}..{fFin:yyyy-MM-dd}).");
            }

            // 3. Día de la semana del bloque
            if (fechas.Any(f => f.DayOfWeek != bloque.DiaSemana))
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: fecha incompatible con el día del bloque ({bloque.DiaSemana}).");
            }

            // 4. Festivos
            if (contexto != null && fechas.Any(f => contexto.EsFestivo(f.Date)))
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: contiene fechas festivas.");
            }

            // 5. Duplicados
            if (fechas.Select(f => f.Date).Distinct().Count() != fechas.Count)
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: contiene fechas duplicadas.");
            }

            // 6. Cardinalidad
            if (fechas.Count != sesionesRequeridas)
            {
                return Result.Failure<List<DateTime>>(
                    $"Gen {gen.Id}: cardinalidad de fechas ({fechas.Count}) no coincide con sesiones requeridas ({sesionesRequeridas}).");
            }

            double horasAcumuladas = 0.0;
            var fechasValidas = new List<DateTime>(fechas.Count);
            foreach (var f in fechas)
            {
                if (horasAcumuladas + 0.001 >= horasObjetivoGen) break;
                fechasValidas.Add(f.Date);
                horasAcumuladas += unidad;
            }

            const double Tolerancia = 0.001;
            if (horasAcumuladas + Tolerancia < horasObjetivoGen)
            {
                return Result.Failure<List<DateTime>>(
                    $"Cobertura incompleta gen {gen.Id}: {horasAcumuladas}/{horasObjetivoGen} h.");
            }

            return Result.Success(fechasValidas);
        }

        public Result<List<AgSesion>> Expandir(AgGen gen, int idVinculacion, AgContexto contexto)
        {
            if (gen == null) throw new ArgumentNullException(nameof(gen));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));
            if (LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Debug($"[AG-RANURAS-TMP] EXPANDIR_ENTRADA corrida={contexto.IdCorrida} individuo={gen.IdIndividuo} gen={gen.Id} aula={gen.IdAula} asig={gen.IdAsignatura} bloque={gen.IdBloqueHorario} horas={gen.HorasGen} fuente={(gen.FechasGen != null && gen.FechasGen.Count > 0 ? "FechasGen" : "calendario_legacy")} fechas={gen.FechasGen?.Count ?? 0} ultima={gen.FechasGen?.Select(d => (DateTime?)d).Max():yyyy-MM-dd} rango={gen.FechaInicioEstimada:yyyy-MM-dd}..{gen.FechaFinEstimada:yyyy-MM-dd}");
            }

            var resFechas = ProyectarFechas(gen, contexto, _planificador);
            if (resFechas.IsFailure)
            {
                return Result.Failure<List<AgSesion>>(resFechas.Error);
            }

            var fechas = resFechas.Value;
            if (fechas.Count == 0)
            {
                return Result.Success(new List<AgSesion>());
            }

            AgBloqueHorario bloque = null;
            if (contexto.BloquesPorId != null) contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out bloque);
            AgAula aula = null;
            if (contexto.AulasPorId != null) contexto.AulasPorId.TryGetValue(gen.IdAula, out aula);
            Tercero docente = null;
            if (contexto.DocentesPorId != null) contexto.DocentesPorId.TryGetValue(gen.IdTercero, out docente);
            Sede sede = null;
            if (contexto.SedesPorId != null) contexto.SedesPorId.TryGetValue(gen.IdSede, out sede);

            // Fase 217 & 786: Si el aula es virtual (EsVirtual == 1), forzar modalidad Plataforma para que CalculadoraViaticaSesion anule viáticos.
            double unidad = bloque != null ? bloque.DuracionHoras : 4.0;
            string modalidad = (aula != null && aula.EsVirtual == 1)
                ? ModalidadTutoriaCatalog.Plataforma
                : ModalidadTutoriaCatalog.NormalizarTutoria(aula?.ModalidadVinculacion ?? ModalidadTutoriaCatalog.Presencial);
            string ciudadSede = sede != null ? sede.Ciudad : string.Empty;
            string residenciaDocente = docente != null ? docente.Residencia : string.Empty;

            var sesiones = new List<AgSesion>(fechas.Count);
            foreach (var f in fechas)
            {
                bool viatica = CalculadoraViaticaSesion.DeterminarViatica(
                    f,
                    modalidad,
                    ciudadSede,
                    residenciaDocente,
                    contexto.Presenciales);

                int idTerceroSesion = (gen.EsVacante == 1 && contexto.IdTerceroVacante > 0) ? contexto.IdTerceroVacante : gen.IdTercero;

                var sesion = new AgSesion
                {
                    IdVinculacion = idVinculacion,
                    IdTercero = idTerceroSesion,
                    IdAula = gen.IdAula,
                    FechaSesion = f,
                    HorarioSesion = bloque != null ? bloque.ObtenerHorarioTexto() : string.Empty,
                    Horas = unidad,
                    Viatica = viatica,
                    Subunidad = contexto.Subunidad,
                    IdIndividuo = gen.IdIndividuo,
                    IdGen = gen.Id,
                    EsVacante = gen.EsVacante,
                    MotivoVacante = gen.MotivoVacante ?? string.Empty,
                    CandidatosPool = gen.CandidatosPool
                };

                sesiones.Add(sesion);
            }

            return Result.Success(sesiones);
        }

        public Result<List<AgSesion>> ExpandirIndividuo(AgIndividuo individuo, Dictionary<long, int> mapaVinculaciones, AgContexto contexto)
        {
            if (individuo == null) throw new ArgumentNullException(nameof(individuo));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            var todasSesiones = new List<AgSesion>();
            var errores = new List<string>();
            foreach (var gen in individuo.Genes)
            {
                int idVinc = (mapaVinculaciones != null && mapaVinculaciones.TryGetValue(gen.Id, out int mappedId)) ? mappedId : 0;

                var expRes = Expandir(gen, idVinc, contexto);
                if (expRes.IsFailure)
                {
                    errores.Add($"gen {gen.Id}: {expRes.Error}");
                    continue;   // Fase 969: aislar el gen, no cancelar todo el individuo
                }

                todasSesiones.AddRange(expRes.Value);
            }

            if (todasSesiones.Count == 0 && errores.Count > 0)
            {
                return Result.Failure<List<AgSesion>>($"Sin sesiones expandibles ({errores.Count}): {string.Join(" | ", errores.Take(5))}");
            }

            // Fase 1296, Fase 1302 & Fase 1310: Diagnóstico explícito de colisión I7 sin descartar sesiones (evita déficit I2)
            var mapaGrupos = (individuo.Genes ?? new List<AgGen>())
                .GroupBy(g => g.Id)
                .ToDictionary(g => g.Key, g => g.First().ClaveGrupo ?? string.Empty);

            var sesionColision = todasSesiones
                .Where(s => s.FechaSesion.HasValue && !string.IsNullOrEmpty(s.HorarioSesion))
                .GroupBy(s => new
                {
                    ClaveGrupo = mapaGrupos.TryGetValue(s.IdGen, out var cg) ? cg : string.Empty,
                    Fecha = s.FechaSesion.Value.Date,
                    s.HorarioSesion
                })
                .FirstOrDefault(gr => gr.Count() > 1 && !string.IsNullOrEmpty(gr.Key.ClaveGrupo));

            if (sesionColision != null)
            {
                var genesColision = sesionColision.Select(s => s.IdGen).Distinct().OrderBy(x => x).ToList();
                LoggerHelper.WarningLimitado(
                    $"I7Colision|{sesionColision.Key.ClaveGrupo}|{sesionColision.Key.Fecha:yyyy-MM-dd}|{sesionColision.Key.HorarioSesion}",
                    $"[AgExpansorSesiones] I7: grupo '{sesionColision.Key.ClaveGrupo}' con {genesColision.Count} gen(es) " +
                    $"[{string.Join(",", genesColision)}] en {sesionColision.Key.Fecha:yyyy-MM-dd} {sesionColision.Key.HorarioSesion}; " +
                    "no se descartan sesiones (evita déficit de horas I2).");
            }

            if (errores.Count > 0)
            {
                LoggerHelper.Warning($"[AgExpansorSesiones] Individuo {individuo.Id}: {errores.Count} gen(es) omitidos: {string.Join(" | ", errores.Take(5))}");
            }

            return Result.Success(todasSesiones);
        }
    }
}
