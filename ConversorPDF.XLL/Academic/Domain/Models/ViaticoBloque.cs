using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Agrupa sesiones con viática en bloques de días consecutivos.
    /// DuracionTotal = (FechaMaxima - FechaMinima).Days + 0.5
    /// </summary>
    public class ViaticoBloque
    {
        private double? _cachedTotalHoras;
        private double? _cachedDuracionTotal;
        private bool? _cachedPernocta;
        private bool _dirty = true;

        /// <summary>
        /// Lista de sesiones agrupadas en este bloque.
        /// </summary>
        public List<SesionListItem> Sesiones { get; } = new List<SesionListItem>();

        /// <summary>
        /// Fecha inicial (mínima) de todas las sesiones en el bloque.
        /// </summary>
        public DateTime? FechaMinima { get; private set; }

        /// <summary>
        /// Fecha final (máxima) de todas las sesiones en el bloque.
        /// Se mantiene intacta para la lógica de consecutividad y algoritmos de agrupación.
        /// </summary>
        public DateTime? FechaMaxima { get; private set; }

        /// <summary>
        /// Regla A: Indica si cualquier sesión del bloque proviene de una vinculación con pernocta último día.
        /// </summary>
        public bool TienePernoctaUltimoDia => Sesiones != null && Sesiones.Any(s => s.IsPernoctaUltimoDia == 1);

        /// <summary>
        /// Fecha de regreso de la comisión: último día de clase + 1 día si hay pernocta en el último día.
        /// </summary>
        public DateTime? FechaFinComision => FechaMaxima?.AddDays(TienePernoctaUltimoDia ? 1 : 0);

        /// <summary>
        /// Lista única de identificadores de vinculación a los que pertenecen las sesiones del bloque.
        /// </summary>
        public IReadOnlyList<int> IdsVinculacion => Sesiones
            .Where(s => s.IdVinculacion > 0)
            .Select(s => s.IdVinculacion)
            .Distinct()
            .ToList();

        /// <summary>
        /// Agrega una sesión al bloque y actualiza las fechas extremas.
        /// </summary>
        /// <param name="sesion">Sesión a integrar.</param>
        public void Agregar(SesionListItem sesion)
        {
            if (sesion?.FechaSesion == null) return;

            Sesiones.Add(sesion);
            var fecha = sesion.FechaSesion.Value.Date;

            if (Sesiones.Count == 1)
            {
                FechaMinima = fecha;
                FechaMaxima = fecha;
            }
            else
            {
                if (FechaMinima.HasValue && fecha < FechaMinima.Value) FechaMinima = fecha;
                if (FechaMaxima.HasValue && fecha > FechaMaxima.Value) FechaMaxima = fecha;
            }

            _dirty = true; // Invalida la caché al modificar la colección
            _cachedTotalHoras = null;
            _cachedDuracionTotal = null;
            _cachedPernocta = null;
        }

        /// <summary>
        /// Invalida la caché de cálculos. Llamar si se modifica <see cref="Sesiones"/> externamente.
        /// </summary>
        public void InvalidateCache()
        {
            _dirty = true;
            _cachedTotalHoras = null;
            _cachedDuracionTotal = null;
            _cachedPernocta = null;
        }

        /// <summary>
        /// Obtiene la suma total de horas de todas las sesiones en el bloque.
        /// Utiliza caché interna.
        /// </summary>
        public double TotalHoras
        {
            get
            {
                if (!_dirty && _cachedTotalHoras.HasValue) return _cachedTotalHoras.Value;
                _cachedTotalHoras = Math.Round(Sesiones.Sum(s => s.Horas), 1);
                _dirty = false;
                return _cachedTotalHoras.Value;
            }
        }

        /// <summary>
        /// Obtiene la duración total en días (calculada como días calendario de comisión + 0.5 de holgura).
        /// Considera FechaFinComision (+1 día si aplica pernocta en último día).
        /// Utiliza caché interna.
        /// </summary>
        public double DuracionTotal
        {
            get
            {
                bool pernoctaActual = TienePernoctaUltimoDia;
                if (!_dirty && _cachedDuracionTotal.HasValue && _cachedPernocta == pernoctaActual) 
                    return _cachedDuracionTotal.Value;

                _cachedPernocta = pernoctaActual;
                _cachedDuracionTotal = (FechaMinima.HasValue && FechaFinComision.HasValue)
                    ? Math.Round((FechaFinComision.Value - FechaMinima.Value).Days + 0.5, 1)
                    : 0;
                _dirty = false;
                return _cachedDuracionTotal.Value;
            }
        }

        /// <summary>
        /// Texto formateado para visualización en el árbol UI utilizando FechaFinComision.
        /// </summary>
        public string DisplayText =>
            FechaMinima.HasValue && FechaFinComision.HasValue
            ? $"📅 {FechaMinima.Value:dd/MM/yyyy} - {FechaFinComision.Value:dd/MM/yyyy}  |  {TotalHoras}h  |  {DuracionTotal} días"
            : "📅 S/N";

        /// <summary>
        /// Algoritmo de agrupación: toma una lista plana de sesiones y las agrupa en bloques 
        /// si ocurren en días consecutivos o el mismo día.
        /// </summary>
        /// <param name="sesiones">Colección de sesiones filtradas para viática.</param>
        /// <returns>Lista de bloques de viáticos generada.</returns>
        public static List<ViaticoBloque> AgruparConsecutivos(List<SesionListItem> sesiones)
        {
            var bloques = new List<ViaticoBloque>();
            if (sesiones == null || !sesiones.Any()) return bloques;

            // 1. Filtrar sesiones con fecha y ordenar
            var filtradas = sesiones
                .Where(s => s.FechaSesion.HasValue)
                .OrderBy(s => s.FechaSesion.Value.Date)
                .ToList();

            if (!filtradas.Any()) return bloques;

            // 2. Iniciar primer bloque
            var bloqueActual = new ViaticoBloque();
            bloqueActual.Agregar(filtradas[0]);

            // 3. Iterar y agrupar consecutivos (o mismo día)
            for (int i = 1; i < filtradas.Count; i++)
            {
                DateTime? ultima = bloqueActual.FechaMaxima;
                DateTime siguiente = filtradas[i].FechaSesion.Value.Date;

                // Si es el mismo día o el día siguiente, se agrega al bloque actual
                if (ultima.HasValue && siguiente <= ultima.Value.AddDays(1).Date)
                {
                    bloqueActual.Agregar(filtradas[i]);
                }
                else
                {
                    bloques.Add(bloqueActual);
                    bloqueActual = new ViaticoBloque();
                    bloqueActual.Agregar(filtradas[i]);
                }
            }

            bloques.Add(bloqueActual);
            return bloques;
        }
    }
}
