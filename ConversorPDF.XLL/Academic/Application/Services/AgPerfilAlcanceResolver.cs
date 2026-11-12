using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Alcance de catálogo que debe adoptar el árbol académico al cargar un perfil guardado (Fase 1580).
    /// </summary>
    internal sealed class AgPerfilAlcance
    {
        /// <summary> Períodos del perfil presentes en el catálogo. </summary>
        public IReadOnlyList<int> Periodos { get; }

        /// <summary> Sedes del perfil presentes en el catálogo de sedes activas. </summary>
        public IReadOnlyList<int> Sedes { get; }

        /// <summary> Programas del perfil presentes en el catálogo. </summary>
        public IReadOnlyList<int> Programas { get; }

        /// <summary> Tipo de programa mínimo que cubre todas las combinaciones (TODOS / PREGRADO / POSGRADO). </summary>
        public string TipoPrograma { get; }

        /// <summary> Combinaciones del perfil cuya tupla período/sede/programa no existe en los catálogos. </summary>
        public int CombosFueraDeCatalogo { get; }

        internal AgPerfilAlcance(
            IReadOnlyList<int> periodos,
            IReadOnlyList<int> sedes,
            IReadOnlyList<int> programas,
            string tipoPrograma,
            int combosFueraDeCatalogo)
        {
            Periodos = periodos ?? new List<int>();
            Sedes = sedes ?? new List<int>();
            Programas = programas ?? new List<int>();
            TipoPrograma = string.IsNullOrWhiteSpace(tipoPrograma) ? AgTipoProgramaCatalog.Todos : tipoPrograma;
            CombosFueraDeCatalogo = combosFueraDeCatalogo;
        }
    }

    /// <summary>
    /// Resuelve el alcance (períodos, sedes, programas y tipo de programa) que un perfil necesita en el árbol
    /// de 6 niveles para que sus combinaciones no se omitan por el filtro vigente (Fase 1580).
    /// </summary>
    internal static class AgPerfilAlcanceResolver
    {
        /// <summary>
        /// Calcula el alcance del perfil intersectado con los catálogos cargados.
        /// Las combinaciones cuyo período, sede o programa no existen en catálogo se contabilizan
        /// en <see cref="AgPerfilAlcance.CombosFueraDeCatalogo"/> y no se incorporan al alcance.
        /// </summary>
        public static AgPerfilAlcance Resolver(
            IEnumerable<AgContextoSeleccionDto> nodos,
            IEnumerable<Periodo> periodos,
            IEnumerable<Sede> sedes,
            IEnumerable<Programa> programas)
        {
            var lista = nodos?.Where(n => n != null).ToList() ?? new List<AgContextoSeleccionDto>();

            var setPeriodos = new HashSet<int>((periodos ?? Enumerable.Empty<Periodo>())
                .Where(p => p != null)
                .Select(p => p.Id));

            var setSedes = new HashSet<int>((sedes ?? Enumerable.Empty<Sede>())
                .Where(s => s != null)
                .Select(s => s.Id));

            var mapProgramas = (programas ?? Enumerable.Empty<Programa>())
                .Where(p => p != null)
                .GroupBy(p => p.Id)
                .ToDictionary(g => g.Key, g => g.First());

            var periodosAlcance = new List<int>();
            var sedesAlcance = new List<int>();
            var programasAlcance = new List<int>();
            var vistosPeriodos = new HashSet<int>();
            var vistosSedes = new HashSet<int>();
            var vistosProgramas = new HashSet<int>();
            int fueraDeCatalogo = 0;

            foreach (var dto in lista)
            {
                // Catálogo vacío = sin información para validar; se conserva el id del perfil.
                bool periodoOk = dto.IdPeriodo <= 0 || setPeriodos.Count == 0 || setPeriodos.Contains(dto.IdPeriodo);
                bool sedeOk = dto.IdSede <= 0 || setSedes.Count == 0 || setSedes.Contains(dto.IdSede);
                bool programaOk = dto.IdPrograma <= 0 || mapProgramas.Count == 0 || mapProgramas.ContainsKey(dto.IdPrograma);

                if (!periodoOk || !sedeOk || !programaOk)
                {
                    fueraDeCatalogo++;
                    continue;
                }

                if (dto.IdPeriodo > 0 && vistosPeriodos.Add(dto.IdPeriodo)) periodosAlcance.Add(dto.IdPeriodo);
                if (dto.IdSede > 0 && vistosSedes.Add(dto.IdSede)) sedesAlcance.Add(dto.IdSede);
                if (dto.IdPrograma > 0 && vistosProgramas.Add(dto.IdPrograma)) programasAlcance.Add(dto.IdPrograma);
            }

            string tipo = ResolverTipoPrograma(programasAlcance, mapProgramas);
            return new AgPerfilAlcance(periodosAlcance, sedesAlcance, programasAlcance, tipo, fueraDeCatalogo);
        }

        /// <summary>
        /// Determina el tipo de programa que cubre todos los programas del perfil:
        /// un solo nivel cuando son homogéneos y <see cref="AgTipoProgramaCatalog.Todos"/> cuando son mixtos.
        /// </summary>
        private static string ResolverTipoPrograma(
            IReadOnlyCollection<int> idsProgramas,
            IReadOnlyDictionary<int, Programa> mapProgramas)
        {
            if (idsProgramas == null || idsProgramas.Count == 0) return AgTipoProgramaCatalog.Todos;

            bool hayPregrado = false;
            bool hayPosgrado = false;

            foreach (var id in idsProgramas)
            {
                if (!mapProgramas.TryGetValue(id, out var programa) || programa == null) continue;

                if (AgTipoProgramaCatalog.Nivel(programa.Modalidad) == AgNivelAcademico.Posgrado)
                {
                    hayPosgrado = true;
                }
                else
                {
                    hayPregrado = true;
                }
            }

            if (hayPregrado && hayPosgrado) return AgTipoProgramaCatalog.Todos;
            if (hayPosgrado) return AgTipoProgramaCatalog.Posgrado;
            if (hayPregrado) return AgTipoProgramaCatalog.Pregrado;
            return AgTipoProgramaCatalog.Todos;
        }
    }
}
