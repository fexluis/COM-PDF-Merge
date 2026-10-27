using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Catálogo canónico de territoriales para el filtro de terceros del módulo AG (Fase 1471).
    /// Unifica la tabla <c>territoriales</c> con los valores realmente presentes en
    /// <c>terceros.departamento</c> / <c>terceros.subunidad</c>, que pueden venir compuestos ("BOLIVAR,VALLE").
    /// </summary>
    internal static class AgTerritorialCatalog
    {
        /// <summary>Opción "sin filtro" del combo territorial (misma etiqueta usada por el filtro).</summary>
        public const string Todas = "TODAS LAS TERRITORIALES";

        /// <summary>
        /// Separa un valor compuesto por comas y descarta vacíos ("BOLIVAR,VALLE" → "BOLIVAR", "VALLE").
        /// </summary>
        public static IEnumerable<string> Separar(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor)) yield break;

            foreach (var parte in valor.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var limpio = parte.Trim();
                if (limpio.Length > 0) yield return limpio;
            }
        }

        /// <summary>
        /// Construye la lista ordenada y sin duplicados del combo: catálogo oficial ∪ valores de terceros,
        /// normalizados a mayúsculas y con los compuestos ya separados.
        /// </summary>
        public static List<string> Construir(IEnumerable<string> catalogo, IEnumerable<string> valoresTerceros)
        {
            var set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var valor in catalogo ?? Enumerable.Empty<string>())
            {
                foreach (var parte in Separar(valor)) set.Add(parte.ToUpperInvariant());
            }

            foreach (var valor in valoresTerceros ?? Enumerable.Empty<string>())
            {
                foreach (var parte in Separar(valor)) set.Add(parte.ToUpperInvariant());
            }

            return set.ToList();
        }
    }
}
