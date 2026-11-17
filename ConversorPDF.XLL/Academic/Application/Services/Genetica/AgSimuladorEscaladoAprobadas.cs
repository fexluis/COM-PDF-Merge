using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Preparación puramente informativa: calendario, no factibilidad docente ni consentimiento T4.</summary>
    internal sealed class AgSimuladorEscaladoAprobadas
    {
        internal static bool OrigenIdentificable(AgContexto c, AgAula aula)
        {
            var m = c?.ManifiestoCobertura; var s = c?.PlanAprobadoSnapshot;
            if (aula == null || aula.Estado != AgEstadoAula.Aprobado || string.IsNullOrWhiteSpace(aula.AulaRef)
                || m == null || s == null || c.Ventana == null || m.Huella != s.Manifiesto.Huella
                || c.Ventana.FechaInicio.Date != m.Inicio || c.Ventana.FechaFin.Date != m.Fin
                || !m.Relaciones.Any(r => r.Clave.IdAula == aula.Id)
                || !c.Aulas.Contains(aula) || c.Aulas.Count(a => a.Id == aula.Id) != 1
                || c.Programas.Count(p => p.Id == aula.IdPrograma) != 1
                || AgDiasVentana.Dias(aula.DiasVentana).Count == 0) return false;
            var territorial = m.Relaciones.Select(r => r.Clave.IdTerritorial).Distinct().ToList();
            if (territorial.Count != 1) return false;
            var actual = new AgConstructorManifiestoCobertura().Construir(c,
                m.Relaciones.Select(r => r.Clave.IdAula).Distinct(), territorial[0], c.Subunidad, m.Inicio, m.Fin);
            if (actual.IsFailure || actual.Value.Huella != m.Huella) return false;
            // El índice es otra entrada al planificador: no admitir relaciones divergentes u omitidas.
            var relaciones = c.AulaAsignaturas.Where(r => r.IdAula == aula.Id).ToList();
            if (relaciones.Count != c.AulaAsignaturaPorAulaAsig.Values.Count(r => r.IdAula == aula.Id)
                || relaciones.Any(r => !c.AulaAsignaturaPorAulaAsig.TryGetValue((r.IdAula, r.IdAsignatura), out var indice)
                    || !ReferenceEquals(indice, r))) return false;
            var validacion = new AgValidadorPlanAprobado().Validar(s, c);
            if (validacion.Diagnosticos.Count != 0) return false;
            var v = validacion.Aulas.Single(a => a.IdAula == aula.Id);
            // Déficit temporal no equivale a identidad desconocida. SIN_SESIONES sólo es admisible
            // si existe una vinculación única reconciliada; RELACION_AUSENTE/ambigua nunca se admite.
            string[] temporales = { "SIN_SESIONES:", "SESION_FUERA_DE_VENTANA:", "SESION_EN_FESTIVO:", "CRUCE_AULA:", "CRUCE_DOCENTE:" };
            return v.Exceso <= 0.001 && v.Relaciones.All(r => r.Diagnosticos.All(d =>
                temporales.Any(p => d.StartsWith(p, StringComparison.Ordinal))));
        }

        /// <summary>Evalúa todas las aprobadas con déficit real verificado. Nunca abre una conexión ni aplica decisiones.</summary>
        public IReadOnlyList<AgPropuestaEscaladoAprobada> Evaluar(AgContexto contexto, CancellationToken ct = default(CancellationToken))
        {
            var propuestas = new List<AgPropuestaEscaladoAprobada>();
            if (contexto?.Aulas == null) return propuestas.AsReadOnly();
            foreach (var aula in contexto.Aulas.OrderBy(a => a.Id))
            {
                ct.ThrowIfCancellationRequested();
                if (!OrigenIdentificable(contexto, aula)) continue;
                var validacion = new AgValidadorPlanAprobado().Validar(contexto.PlanAprobadoSnapshot, contexto)
                    .Aulas.Single(a => a.IdAula == aula.Id);
                // La evidencia canónica completa prevalece frente a cualquier simulación desde cero.
                if (validacion.Deficit <= 0.001) continue;
                string huella = AgHuellaPlanVentana.Calcular(contexto, aula);
                var copia = AgEscaladorRanuraAula.ClonarContextoParaSimulacion(contexto);
                var aislada = copia.Aulas.Single(a => a.Id == aula.Id);
                copia.AulaAsignaturas = copia.AulaAsignaturas.Where(r => r.Requerida == 1).ToList();
                copia.ConstruirIndices();
                double unidad = AgUnidadesSemanales.HorasSesionAula(aislada,
                    AgNivelAcademicoCatalogo.Desde(copia.Programas.Single(p => p.Id == aislada.IdPrograma).Modalidad));
                var antes = Planificar(copia, aislada, aislada.DiasVentana, unidad, ct);
                string siguiente = AgEscaleraVentanas.Siguiente(aislada.DiasVentana);
                string destino = string.IsNullOrEmpty(siguiente) ? aislada.DiasVentana : siguiente;
                var sin = Planificar(copia, aislada, destino, unidad, ct);
                // Misma búsqueda canónica de LEGADO pendiente que Disponible; sin marcar Autorizado.
                var pendientes = validacion.Relaciones.ToDictionary(r => r.Clave.IdAsignatura, r =>
                    Math.Max(0, (int)Math.Ceiling(r.HorasEsperadas / unidad) - antes.Planes
                        .Where(p => p.IdAsignatura == r.Clave.IdAsignatura).Sum(p => p.Sesiones.Count)));
                var evaluacion = new AgResultadoEvaluacionVentana { CabeEnPlazo = pendientes.Values.All(n => n == 0),
                    SesionesPendientesPorAsignatura = pendientes, IdAsignaturasSaturadas = pendientes.Where(p => p.Value > 0).Select(p => p.Key).ToList() };
                var candidatas = new AgEscaladorRanuraAula(() => null)
                    .IdentificarPatronesRedistribuibles(contexto, aula, evaluacion, ct);
                foreach (var candidata in candidatas)
                {
                    var r = copia.AulaAsignaturaPorAulaAsig[(aula.Id, candidata.IdAsignatura)];
                    r.OrigenPatron = candidata.OrigenPropuesto; r.DiasPatron = candidata.DiasPatronPropuesto;
                }
                var con = Planificar(copia, aislada, destino, unidad, ct);
                var detalles = validacion.Relaciones.Select(r => new AgDetalleEscaladoAprobada(r,
                    contexto.AulaAsignaturaPorAulaAsig[(aula.Id, r.Clave.IdAsignatura)],
                    contexto.PlanAprobadoSnapshot.Relaciones.Where(p => p.ClavesCandidatas.Count == 1
                        && p.ClavesCandidatas[0].Valor == r.Clave.Valor).SelectMany(p => p.Sesiones), antes, sin, con, unidad,
                    candidatas.Any(p => p.IdAsignatura == r.Clave.IdAsignatura))).ToList();
                ct.ThrowIfCancellationRequested();
                if (AgHuellaPlanVentana.Calcular(contexto, aula) != huella)
                    throw new InvalidOperationException("CONFIGURACION_MODIFICADA_DURANTE_PROPUESTA");
                propuestas.Add(new AgPropuestaEscaladoAprobada(aula, contexto.PlanAprobadoSnapshot.HuellaOrigen,
                    huella, siguiente, copia.Ventana.FechaFin, antes.BusquedaAgotada || sin.BusquedaAgotada || con.BusquedaAgotada, detalles));
            }
            return propuestas.AsReadOnly();
        }

        private static AgPlanOleadasResultado Planificar(AgContexto c, AgAula aula, string dias, double unidad, CancellationToken ct)
        {
            var patrones = c.AulaAsignaturaPorAulaAsig.Values.Where(r => r.IdAula == aula.Id
                && !string.Equals(r.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(r.DiasPatron)).ToDictionary(r => r.IdAsignatura, r => r.DiasPatron);
            // Misma autoridad fechada, sin la proyección fuera de plazo del evaluador de demanda.
            var resultado = new AgPlanificadorOleadas().ConstruirEvaluacion(AgEvaluadorDemandaVentana.Demandas(c, aula),
                dias, c.Ventana.FechaInicio, c.Ventana.FechaFin, unidad, c.EsFestivo, c.Bloques, aula.Jornada, patrones, ct);
            if (resultado.IsFailure) throw new InvalidOperationException("SIMULACION_APROBADA:" + resultado.Error);
            return resultado.Value;
        }
    }
}
