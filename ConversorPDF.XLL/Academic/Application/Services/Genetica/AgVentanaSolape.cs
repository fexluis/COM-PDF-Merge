using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Utilidades de evaluación y agrupamiento por solape temporal entre ventanas de asignaturas (Fase 1568).
    /// Utiliza semántica conservadora: si falta cualquier fecha en la comparación, se asume que solapan.
    /// </summary>
    public static class AgVentanaSolape
    {
        /// <summary>
        /// Evalúa si dos ventanas temporales se intersecan en fechas calendario (inclusivo).
        /// Semántica conservadora: si cualquiera de los 4 extremos es null, devuelve <c>true</c>.
        /// </summary>
        public static bool SeSolapan(DateTime? iniA, DateTime? finA, DateTime? iniB, DateTime? finB)
        {
            if (!iniA.HasValue || !finA.HasValue || !iniB.HasValue || !finB.HasValue)
            {
                return true;
            }

            return iniA.Value.Date <= finB.Value.Date && iniB.Value.Date <= finA.Value.Date;
        }

        /// <summary>
        /// Agrupa los elementos en componentes conexas de solape temporal (clusters).
        /// Dos elementos pertenecen al mismo cluster si se solapan directamente o a través de una cadena transitiva.
        /// </summary>
        public static List<List<T>> AgruparPorSolape<T>(
            IEnumerable<T> items, Func<T, DateTime?> ini, Func<T, DateTime?> fin)
        {
            var resultado = new List<List<T>>();
            if (items == null) return resultado;

            var lista = items.ToList();
            int n = lista.Count;
            if (n == 0) return resultado;
            if (n == 1)
            {
                resultado.Add(lista);
                return resultado;
            }

            // Grafo de adyacencia para encontrar componentes conexas
            var adj = new List<int>[n];
            for (int i = 0; i < n; i++) adj[i] = new List<int>();

            for (int i = 0; i < n; i++)
            {
                var iIni = ini(lista[i]);
                var iFin = fin(lista[i]);
                for (int j = i + 1; j < n; j++)
                {
                    if (SeSolapan(iIni, iFin, ini(lista[j]), fin(lista[j])))
                    {
                        adj[i].Add(j);
                        adj[j].Add(i);
                    }
                }
            }

            var visitados = new bool[n];
            for (int i = 0; i < n; i++)
            {
                if (visitados[i]) continue;

                var cluster = new List<T>();
                var cola = new Queue<int>();
                cola.Enqueue(i);
                visitados[i] = true;

                while (cola.Count > 0)
                {
                    int curr = cola.Dequeue();
                    cluster.Add(lista[curr]);

                    foreach (var vecino in adj[curr])
                    {
                        if (!visitados[vecino])
                        {
                            visitados[vecino] = true;
                            cola.Enqueue(vecino);
                        }
                    }
                }

                resultado.Add(cluster);
            }

            return resultado;
        }
    }
}

