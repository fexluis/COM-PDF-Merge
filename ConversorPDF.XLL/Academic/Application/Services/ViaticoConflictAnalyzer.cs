using System;
using System.Collections.Generic;
using System.Text;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación para detectar solapamientos de fechas entre bloques de viáticos.
    /// </summary>
    public static class ViaticoConflictAnalyzer
    {
        /// <summary>
        /// Detecta qué bloques de un docente tienen fechas solapadas
        /// con bloques de OTRA sede distinta.
        /// </summary>
        /// <param name="bloquesDelDocente">Lista de tuplas (Sede, Bloque) de un docente.</param>
        /// <param name="salvedades">Colección opcional de salvedades para ignorar conflictos aprobados previamente.</param>
        /// <returns>Conjunto de bloques conflictivos.</returns>
        public static HashSet<ViaticoBloque> DetectarBloquesCruzados(
            List<(string Sede, ViaticoBloque Bloque)> bloquesDelDocente,
            IReadOnlyCollection<SalvedadConflicto> salvedades = null)
        {
            var conflictivos = new HashSet<ViaticoBloque>();

            if (bloquesDelDocente == null || bloquesDelDocente.Count < 2)
                return conflictivos;

            for (int i = 0; i < bloquesDelDocente.Count; i++)
            {
                var (sedeA, bloqueA) = bloquesDelDocente[i];

                for (int j = i + 1; j < bloquesDelDocente.Count; j++)
                {
                    var (sedeB, bloqueB) = bloquesDelDocente[j];

                    // Solo detectar cruce entre sedes DIFERENTES
                    if (string.Equals(sedeA, sedeB, System.StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Delegar al helper privado — DRY (R3)
                    if (SeSolapan(bloqueA, bloqueB))
                    {
                        if (!EsCruceSalvado(bloqueA, bloqueB, salvedades))
                        {
                            conflictivos.Add(bloqueA);
                            conflictivos.Add(bloqueB);
                        }
                    }
                }
            }

            return conflictivos;
        }

        private static bool EsCruceSalvado(ViaticoBloque bloqueA, ViaticoBloque bloqueB, IReadOnlyCollection<SalvedadConflicto> salvedades)
        {
            if (salvedades == null || salvedades.Count == 0)
                return false;

            var paresMismoDia = new List<(SesionListItem a, SesionListItem b)>();

            foreach (var sA in bloqueA.Sesiones)
            {
                if (!sA.FechaSesion.HasValue) continue;
                
                foreach (var sB in bloqueB.Sesiones)
                {
                    if (!sB.FechaSesion.HasValue) continue;

                    if (sA.FechaSesion.Value.Date == sB.FechaSesion.Value.Date)
                    {
                        paresMismoDia.Add((sA, sB));
                    }
                }
            }

            if (paresMismoDia.Count > 0)
            {
                foreach (var (sA, sB) in paresMismoDia)
                {
                    int minId = System.Math.Min(sA.IdSesion, sB.IdSesion);
                    int maxId = System.Math.Max(sA.IdSesion, sB.IdSesion);
                    
                    bool tieneSalvedad = false;
                    foreach (var s in salvedades)
                    {
                        if (s.IdSesionA == minId && s.IdSesionB == maxId)
                        {
                            tieneSalvedad = true;
                            break;
                        }
                    }
                    
                    if (!tieneSalvedad)
                        return false;
                }
                return true;
            }

            return false;
        }

        /// <summary>
        /// Genera texto detallado para el Tooltip describiendo con qué bloques se solapa el bloque actual.
        /// </summary>
        /// <param name="bloqueA">Bloque con conflicto.</param>
        /// <param name="todosLosBloques">Lista completa de bloques del docente.</param>
        /// <param name="sedeActual">Sede del bloque con conflicto.</param>
        /// <returns>Texto formateado para mostrar en el Tooltip, o <see cref="string.Empty"/> si no hay cruce.</returns>
        public static string ObtenerDetalleCruce(
            ViaticoBloque bloqueA,
            List<(string Sede, ViaticoBloque Bloque)> todosLosBloques,
            string sedeActual)
        {
            if (bloqueA == null || todosLosBloques == null)
                return string.Empty;

            // Guard nulo — R6
            string sedeDisplay = sedeActual?.ToUpper() ?? "SIN SEDE";

            var sb = new StringBuilder();
            sb.AppendLine("⚠️ CONFLICTO DE FECHAS");
            sb.AppendLine($"Este bloque en {sedeDisplay} se solapa con:");

            bool huboCruce = false;
            foreach (var (sedeB, bloqueB) in todosLosBloques)
            {
                if (bloqueB == bloqueA) continue;

                // Solo reportar cruce entre sedes DIFERENTES
                if (string.Equals(sedeActual, sedeB, System.StringComparison.OrdinalIgnoreCase))
                    continue;

                // Reutilizar helper — DRY (R3)
                if (SeSolapan(bloqueA, bloqueB))
                {
                    huboCruce = true;
                    sb.AppendLine();
                    sb.AppendLine($"📍 {sedeB.ToUpper()}");
                    DateTime? fFinB = bloqueB.FechaFinComision ?? bloqueB.FechaMaxima;
                    sb.AppendLine($"   {bloqueB.FechaMinima:dd/MM/yyyy} → {fFinB:dd/MM/yyyy}");
                    sb.AppendLine($"   {bloqueB.TotalHoras}h | {bloqueB.DuracionTotal} días");
                }
            }

            return huboCruce ? sb.ToString() : string.Empty;
        }

        /// <summary>
        /// Determina si dos bloques tienen fechas solapadas de clase efectiva (intersección de intervalos de sesiones).
        /// Se evalúa sobre FechaMaxima (días de clase) para permitir salvedades a nivel de sesión y evitar falsos positivos
        /// en días de solo traslado/pernocta.
        /// Dos rangos [A_start, A_end] y [B_start, B_end] se solapan si:
        ///   A_start &lt;= B_end AND B_start &lt;= A_end
        /// </summary>
        private static bool SeSolapan(ViaticoBloque a, ViaticoBloque b)
        {
            var aFin = a.FechaMaxima;   // días de clase efectivos
            var bFin = b.FechaMaxima;
            if (!a.FechaMinima.HasValue || !aFin.HasValue) return false;
            if (!b.FechaMinima.HasValue || !bFin.HasValue) return false;

            return a.FechaMinima.Value <= bFin.Value
                && b.FechaMinima.Value <= aFin.Value;
        }
    }
}
