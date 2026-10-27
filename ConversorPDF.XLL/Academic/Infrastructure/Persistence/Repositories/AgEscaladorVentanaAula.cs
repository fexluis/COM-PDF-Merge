using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Escalador de ventana por aula (Fase 1229): decide (dry-run) y aplica el siguiente escalón de
    /// AgEscaleraVentanas cuando la demanda de sesiones excede la capacidad real de la ventana.
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal sealed class AgEscaladorVentanaAula
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgEscaladorVentanaAula(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        internal sealed class Decision
        {
            public string AulaRef { get; set; }
            public string Subunidad { get; set; }
            public string Desde { get; set; }
            public string Hacia { get; set; }
            public int DemandaSesiones { get; set; }
            public int CapacidadAntes { get; set; }
            public int CapacidadDespues { get; set; }
            public string Motivo { get; set; }

            public bool Aplica => !string.IsNullOrWhiteSpace(Hacia) && CapacidadDespues > CapacidadAntes;
        }

        /// <summary>Dry-run: no escribe en base de datos.</summary>
        public IReadOnlyList<Decision> Evaluar(AgContexto contexto)
        {
            var decisiones = new List<Decision>();
            if (contexto == null || contexto.Ventana == null || contexto.Aulas == null) return decisiones;

            var ini = contexto.Ventana.FechaInicio.Date;
            var fin = contexto.Ventana.FechaFin.Date;

            foreach (var aula in contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Disponible && !string.IsNullOrWhiteSpace(a.AulaRef)))
            {
                var programa = contexto.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma);
                var nivel = AgNivelAcademicoCatalogo.Desde(programa?.Modalidad);
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula, nivel);

                int demanda = DemandaAula(contexto, aula, unidad);
                if (demanda <= 0) continue;

                int capacidadAntes = Capacidad(aula.DiasVentana, aula.Jornada, unidad, contexto, ini, fin);
                if (demanda <= capacidadAntes) continue;

                foreach (var paso in AgEscaleraVentanas.Escalera(aula.DiasVentana))
                {
                    int capacidad = Capacidad(paso, aula.Jornada, unidad, contexto, ini, fin);
                    if (capacidad <= capacidadAntes) continue;

                    decisiones.Add(new Decision
                    {
                        AulaRef = aula.AulaRef,
                        Subunidad = aula.Subunidad,
                        Desde = aula.DiasVentana,
                        Hacia = paso,
                        DemandaSesiones = demanda,
                        CapacidadAntes = capacidadAntes,
                        CapacidadDespues = capacidad,
                        Motivo = $"demanda {demanda} sesiones > capacidad {capacidadAntes} ranuras"
                    });
                    break;   // Fase 1229: un escalón por corrida (idempotente)
                }
            }
            return decisiones;
        }

        /// <summary>Aplica las decisiones en ag_aula y ag_arbol_nodo; devuelve filas de ag_aula afectadas.</summary>
        public Task<int> AplicarAsync(IReadOnlyList<Decision> decisiones, CancellationToken ct)
            => AplicarAsync(decisiones, 0, ct);

        /// <summary>Aplica las decisiones en ag_aula y ag_arbol_nodo acotando por idArbol (Fase 1258).</summary>
        public async Task<int> AplicarAsync(IReadOnlyList<Decision> decisiones, int idArbol, CancellationToken ct)
        {
            if (decisiones == null || decisiones.Count == 0) return 0;

            const string sqlAula = @"
UPDATE ag_aula
SET dias_ventana = @D
WHERE aula_ref = @Ref
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub
  AND COALESCE(dias_ventana, '') <> @D;";

            const string sqlNodo = @"
UPDATE ag_arbol_nodo
SET dias_ventana = @D
WHERE aula_ref = @Ref
  AND (@IdArbol <= 0 OR idarbol = @IdArbol)
  AND COALESCE(dias_ventana, '') <> @D;";

            int total = 0;
            try
            {
                using (var conn = _connectionFactory())
                {
                    if (conn.State != ConnectionState.Open) conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        foreach (var d in decisiones.Where(x => x.Aplica))
                        {
                            var args = new { D = d.Hacia, Ref = d.AulaRef, Sub = d.Subunidad ?? string.Empty };
                            total += await conn.ExecuteAsync(new CommandDefinition(sqlAula, args, trans, cancellationToken: ct)).ConfigureAwait(false);
                            if (idArbol > 0)   // Fase 1272 / Fase 1276: sin árbol destino no se tocan nodos de otros árboles
                            {
                                await conn.ExecuteAsync(new CommandDefinition(sqlNodo, new { D = d.Hacia, Ref = d.AulaRef, IdArbol = idArbol }, trans, cancellationToken: ct)).ConfigureAwait(false);
                            }
                        }
                        trans.Commit();
                    }
                }

                if (total > 0)
                {
                    ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                }
                LoggerHelper.Info($"[AgEscaladorVentanaAula] Escalado aplicado en {total} aula(s).");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgEscaladorVentanaAula] Error al aplicar escalado: {ex.Message}");
                return 0;
            }
            return total;
        }

        private static int DemandaAula(AgContexto contexto, AgAula aula, double unidad)
        {
            if (contexto.AulaAsignaturaPorAulaAsig == null) return 0;

            // Fase 1259: filtrar asignaturas por programa y semestre del aula (consistente con AgValidadorVentana)
            string semAula = AgFormatoAsignatura.SemestreRomano(aula.Semestre);
            int sesiones = 0;
            foreach (var kv in contexto.AulaAsignaturaPorAulaAsig)
            {
                if (kv.Key.Item1 != aula.Id || kv.Value == null) continue;
                var asig = contexto.Asignaturas?.FirstOrDefault(a => a.Id == kv.Key.Item2);
                if (asig != null && asig.IdPrograma > 0 && aula.IdPrograma > 0 && asig.IdPrograma != aula.IdPrograma) continue;
                if (asig != null && !string.IsNullOrWhiteSpace(semAula) && semAula != "SIN ASIGNAR" &&
                    !string.Equals(AgFormatoAsignatura.SemestreRomano(asig.Semestre), semAula, StringComparison.OrdinalIgnoreCase)) continue;
                if (kv.Value.HorasObjetivo > 0)
                {
                    sesiones += (int)Math.Ceiling(kv.Value.HorasObjetivo / unidad);
                }
            }
            return sesiones;
        }

        private static int Capacidad(string diasVentana, string jornada, double unidad, AgContexto contexto, DateTime ini, DateTime fin)
        {
            var ranuras = AgRanuras.PorVentana(diasVentana, contexto.Bloques, jornada, unidad);
            if (ranuras == null || ranuras.Count == 0) return 0;

            int total = 0;
            foreach (var r in ranuras)
            {
                total += AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo);
            }
            return total;
        }
    }
}
