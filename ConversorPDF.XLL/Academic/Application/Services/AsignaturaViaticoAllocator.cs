using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de dominio para la distribución óptima de bloques de viáticos por asignatura y sede en el reporte Excel.
    /// Implementa las reglas de negocio R1 a R7 garantizando un máximo de 8 comisiones por fila,
    /// preservación de bloques multi-sede intactos y generación de filas de continuación únicamente como último recurso.
    /// Emplea un algoritmo de dos pases para reservar cupo a asignaciones naturales (R1) antes de resolver desbordes (R4).
    /// </summary>
    public static class AsignaturaViaticoAllocator
    {
        /// <summary>
        /// Máximo número de bloques de comisiones permitidos por fila en la plantilla Excel (COMISIÓN 1..8 = AK–BG).
        /// </summary>
        public const int MaxBloquesPorFila = 8;

        private static readonly string[] FormatosFecha =
        {
            "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss",
            "dd/MM/yyyy", "d/M/yyyy",
            "dd/MM/yyyy H:mm:ss", "dd/MM/yyyy HH:mm:ss", "d/M/yyyy H:mm:ss",
            "yyyy/MM/dd", "dd-MM-yyyy"
        };

        /// <summary>
        /// Distribuye los bloques globales de viáticos de un docente entre sus filas de asignaturas/sedes/ventanas.
        /// </summary>
        /// <param name="asignaturas">Colección de grupos de sesiones por asignatura/sede/ventana.</param>
        /// <param name="bloquesGlobales">Bloques de viáticos calculados para el docente.</param>
        /// <param name="advertencias">Colección opcional donde se registran advertencias de reasignación, desbordes o ventanas.</param>
        /// <returns>Lista de filas de exportación con sus respectivos bloques asignados ordenados cronológicamente.</returns>
        public static IReadOnlyList<AsignaturaViaticoFila> Distribuir(
            IReadOnlyList<IReadOnlyList<SesionListItem>> asignaturas,
            IReadOnlyList<ViaticoBloque> bloquesGlobales,
            ICollection<string> advertencias = null)
        {
            if (asignaturas == null || asignaturas.Count == 0)
            {
                return new List<AsignaturaViaticoFila>().AsReadOnly();
            }

            // Preparar filas base: IdVinculacion -> fila, sede normalizada, ventana (U, V) parseada
            var filasBase = new List<FilaContexto>(asignaturas.Count);
            foreach (var asig in asignaturas)
            {
                var primera = asig.FirstOrDefault();
                TryParseFecha(primera?.VinculacionFechaInicio, out DateTime u);
                TryParseFecha(primera?.VinculacionFechaFin, out DateTime v);

                DateTime? uVal = u != default(DateTime) ? (DateTime?)u : null;
                DateTime? vVal = v != default(DateTime) ? (DateTime?)v : null;

                var filaBase = new AsignaturaViaticoFila
                {
                    Asignatura = asig,
                    EsContinuacion = false
                };

                var ctx = new FilaContexto
                {
                    Fila = filaBase,
                    IdsVinculacion = new HashSet<int>(asig.Where(s => s.IdVinculacion > 0).Select(s => s.IdVinculacion)),
                    SedeNormalizada = NormalizarSede(primera?.SedeCiudad, primera?.Subsede),
                    AsignaturaNombre = primera?.AsignaturaNombre ?? string.Empty,
                    U = uVal,
                    V = vVal
                };

                filasBase.Add(ctx);
            }

            if (bloquesGlobales == null || bloquesGlobales.Count == 0)
            {
                return filasBase.Select(f => f.Fila).ToList().AsReadOnly();
            }

            // Para cada bloque por FechaMinima
            var bloquesOrdenados = bloquesGlobales
                .OrderBy(b => b.FechaMinima ?? DateTime.MaxValue)
                .ToList();

            var pendientes = new List<(ViaticoBloque Bloque, List<FilaContexto> Candidatas)>();

            // Pase 1 — R1: Reservar cupo a la fila natural con precedencia estricta
            foreach (var bloque in bloquesOrdenados)
            {
                var idsBloque = new HashSet<int>(bloque.IdsVinculacion);
                var candidatas = filasBase
                    .Where(f => f.IdsVinculacion.Overlaps(idsBloque))
                    .ToList();

                if (candidatas.Count == 1 && candidatas[0].TieneCupo)
                {
                    candidatas[0].Fila.Bloques.Add(bloque);
                    AdvertirVentana(candidatas[0], bloque, advertencias);
                }
                else
                {
                    pendientes.Add((bloque, candidatas));
                }
            }

            // Pase 2 — R2 / R4 / R5 sobre bloques pendientes
            foreach (var item in pendientes)
            {
                var bloque = item.Bloque;
                var candidatas = item.Candidatas;
                DateTime fecha = bloque.FechaMinima ?? DateTime.MinValue;
                string fechaStr = bloque.FechaMinima.HasValue ? bloque.FechaMinima.Value.ToString("yyyy-MM-dd") : "S/F";
                var idsBloque = new HashSet<int>(bloque.IdsVinculacion);
                var sedesRutaBloque = new HashSet<string>(
                    bloque.Sesiones
                        .Select(s => NormalizarSede(s.SedeCiudad, s.Subsede))
                        .Where(s => !string.IsNullOrEmpty(s)),
                    StringComparer.OrdinalIgnoreCase);

                FilaContexto destinoCtx = null;

                // R2: Bloque con sesiones de >= 2 vinculaciones (cruza sedes / vinculaciones)
                if (candidatas.Count >= 2)
                {
                    var candidatasConCupo = candidatas
                        .Where(c => c.TieneCupo)
                        .OrderByDescending(c => bloque.Sesiones.Count(s => c.IdsVinculacion.Contains(s.IdVinculacion)))
                        .ThenBy(c => c.U ?? DateTime.MaxValue)
                        .ToList();

                    if (candidatasConCupo.Count > 0)
                    {
                        destinoCtx = candidatasConCupo[0];
                        string idsStr = string.Join(", ", idsBloque);
                        advertencias?.Add($"Bloque ({fechaStr}) contiene sesiones de múltiples vinculaciones ({idsStr}). Se asignó entero a la fila {destinoCtx.SedeNormalizada}/{destinoCtx.AsignaturaNombre}.");
                    }
                }

                // R4: Si la fila natural no tiene cupo, buscar otra fila con cupo según prioridad
                if (destinoCtx == null)
                {
                    var filasConCupo = filasBase
                        .Where(f => f.TieneCupo)
                        .OrderByDescending(f => f.U.HasValue && f.V.HasValue && fecha >= f.U.Value && fecha <= f.V.Value)
                        .ThenByDescending(f => sedesRutaBloque.Contains(f.SedeNormalizada))
                        .ThenBy(f => f.U.HasValue ? Math.Abs((f.U.Value - fecha).TotalDays) : double.MaxValue)
                        .ThenBy(f => f.U ?? DateTime.MaxValue)
                        .ToList();

                    if (filasConCupo.Count > 0)
                    {
                        destinoCtx = filasConCupo[0];
                        advertencias?.Add(candidatas.Count == 0
                            ? $"Bloque ({fechaStr}) sin vinculación coincidente; asignado a la fila {destinoCtx.SedeNormalizada}/{destinoCtx.AsignaturaNombre}."
                            : $"Bloque ({fechaStr}) reasignado por saturación a la fila {destinoCtx.SedeNormalizada}/{destinoCtx.AsignaturaNombre}.");
                    }
                }

                // R5: Si ninguna fila tiene cupo -> fila de continuación (último recurso)
                if (destinoCtx == null)
                {
                    FilaContexto filaBase = null;
                    if (candidatas.Count > 0)
                    {
                        filaBase = candidatas
                            .OrderByDescending(c => bloque.Sesiones.Count(s => c.IdsVinculacion.Contains(s.IdVinculacion)))
                            .ThenBy(c => c.U ?? DateTime.MaxValue)
                            .First();
                    }
                    else
                    {
                        filaBase = filasBase[0];
                    }

                    destinoCtx = filaBase.Continuaciones.FirstOrDefault(c => c.TieneCupo);
                    if (destinoCtx == null)
                    {
                        var nuevaFila = new AsignaturaViaticoFila
                        {
                            Asignatura = filaBase.Fila.Asignatura,
                            EsContinuacion = true
                        };
                        destinoCtx = new FilaContexto
                        {
                            Fila = nuevaFila,
                            IdsVinculacion = filaBase.IdsVinculacion,
                            SedeNormalizada = filaBase.SedeNormalizada,
                            AsignaturaNombre = filaBase.AsignaturaNombre,
                            U = filaBase.U,
                            V = filaBase.V
                        };
                        filaBase.Continuaciones.Add(destinoCtx);
                    }

                    advertencias?.Add($"Bloque ({fechaStr}) asignado a fila de continuación para {filaBase.SedeNormalizada}/{filaBase.AsignaturaNombre} por falta de cupo global.");
                }

                // Asignar el bloque a la fila de destino
                destinoCtx.Fila.Bloques.Add(bloque);

                // R6: Validación de fecha dentro de la ventana de vinculación [U, V]
                AdvertirVentana(destinoCtx, bloque, advertencias);
            }

            // Ordenar bloques de cada fila por FechaMinima y emitir fila base seguida de sus continuaciones
            var resultado = new List<AsignaturaViaticoFila>();
            foreach (var baseCtx in filasBase)
            {
                baseCtx.Fila.Bloques.Sort((a, b) => Nullable.Compare(a.FechaMinima, b.FechaMinima));
                resultado.Add(baseCtx.Fila);

                foreach (var contCtx in baseCtx.Continuaciones)
                {
                    contCtx.Fila.Bloques.Sort((a, b) => Nullable.Compare(a.FechaMinima, b.FechaMinima));
                    resultado.Add(contCtx.Fila);
                }
            }

            return resultado.AsReadOnly();
        }

        private static void AdvertirVentana(FilaContexto ctx, ViaticoBloque bloque, ICollection<string> advertencias)
        {
            if (ctx.U.HasValue && ctx.V.HasValue && bloque.FechaMinima.HasValue)
            {
                if (bloque.FechaMinima.Value < ctx.U.Value || bloque.FechaMinima.Value > ctx.V.Value)
                {
                    string fechaStr = bloque.FechaMinima.Value.ToString("yyyy-MM-dd");
                    advertencias?.Add($"Bloque ({fechaStr}) fuera de la ventana de vinculación [{ctx.U.Value:yyyy-MM-dd} a {ctx.V.Value:yyyy-MM-dd}] en fila {ctx.SedeNormalizada}/{ctx.AsignaturaNombre}.");
                }
            }
        }

        private static string NormalizarSede(string ciudad, string subsede)
        {
            string s = string.IsNullOrWhiteSpace(ciudad) ? subsede : ciudad;
            return s?.Trim().ToUpperInvariant() ?? string.Empty;
        }

        private static bool TryParseFecha(string fechaStr, out DateTime fecha)
        {
            if (string.IsNullOrWhiteSpace(fechaStr))
            {
                fecha = default(DateTime);
                return false;
            }

            string f = fechaStr.Trim();
            if (DateTime.TryParseExact(f, FormatosFecha, CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
            {
                return true;
            }

            var cultura = f.Contains("/") ? CultureInfo.GetCultureInfo("es-CO") : CultureInfo.InvariantCulture;
            return DateTime.TryParse(f, cultura, DateTimeStyles.None, out fecha);
        }

        private sealed class FilaContexto
        {
            public AsignaturaViaticoFila Fila { get; set; }
            public HashSet<int> IdsVinculacion { get; set; }
            public string SedeNormalizada { get; set; }
            public string AsignaturaNombre { get; set; }
            public DateTime? U { get; set; }
            public DateTime? V { get; set; }
            public List<FilaContexto> Continuaciones { get; } = new List<FilaContexto>();
            public bool TieneCupo => Fila.Bloques.Count < MaxBloquesPorFila;
        }
    }
}
