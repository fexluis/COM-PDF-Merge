using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio puro de resolución y filtrado de elementos de árbol académico (Fase 1314 / 1354).
    /// </summary>
    internal static class AgFiltroArbolResolver
    {
        /// <summary>
        /// Resuelve los IDs válidos intersectando con los IDs vigentes en catálogo.
        /// Opcionalmente extrae los IDs obsoletos para auditoría o purga.
        /// </summary>
        public static List<int> Resolver(IEnumerable<int> ids, IEnumerable<int> vigentes, List<int> obsoletos = null)
        {
            if (ids == null) return new List<int>();
            var setVigentes = new HashSet<int>(vigentes ?? Enumerable.Empty<int>());
            var resueltos = new List<int>();

            foreach (var id in ids)
            {
                if (setVigentes.Contains(id))
                {
                    resueltos.Add(id);
                }
                else
                {
                    obsoletos?.Add(id);
                }
            }

            return resueltos;
        }

        /// <summary>
        /// Comprueba si un ID coincide con el filtro seleccionado.
        /// Regla "Vacío = Todos": si la lista de seleccionados está vacía o es null, coincide siempre.
        /// </summary>
        public static bool CoincideConFiltro(int id, IReadOnlyCollection<int> seleccionados)
        {
            if (seleccionados == null || seleccionados.Count == 0)
            {
                return true;
            }
            return seleccionados.Contains(id);
        }

        public static List<Periodo> FiltrarPeriodos(IEnumerable<Periodo> periodos, IReadOnlyCollection<int> sel)
        {
            if (periodos == null) return new List<Periodo>();
            if (sel == null || sel.Count == 0)
            {
                var activos = periodos.Where(p => p.Activo).ToList();
                return activos.Any() ? activos : periodos.ToList();
            }
            return periodos.Where(p => sel.Contains(p.Id)).ToList();
        }

        public static List<Sede> FiltrarSedes(IEnumerable<Sede> sedes, IReadOnlyCollection<int> sel)
        {
            if (sedes == null) return new List<Sede>();
            if (sel == null || sel.Count == 0)
            {
                var activos = sedes.Where(s => s.Activo == 1).ToList();
                return activos.Any() ? activos : sedes.ToList();
            }
            return sedes.Where(s => sel.Contains(s.Id)).ToList();
        }

        public static List<Programa> FiltrarProgramas(IEnumerable<Programa> programas, string tipo, IReadOnlyCollection<int> sel)
        {
            return FiltrarProgramas(programas, sel, tipo);
        }

        public static List<Programa> FiltrarProgramas(IEnumerable<Programa> programas, IReadOnlyCollection<int> sel, string tipo = null)
        {
            if (programas == null) return new List<Programa>();
            var lista = programas;

            if (!string.IsNullOrWhiteSpace(tipo))
            {
                lista = AgTipoProgramaCatalog.Filtrar(lista, tipo);
            }

            if (sel == null || sel.Count == 0)
            {
                return lista.ToList();
            }

            return lista.Where(p => sel.Contains(p.Id)).ToList();
        }
    }
}
