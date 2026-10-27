using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Operador de selección por torneo (Tournament Selection) para preservar presión selectiva hacia mayor fitness.
    /// </summary>
    public class AgOperadorSeleccion
    {
        private readonly Random _random;

        public AgOperadorSeleccion(Random random = null)
        {
            _random = random ?? new Random();
        }

        public AgIndividuo SeleccionarPorTorneo(List<AgIndividuo> poblacion, int tamanoTorneo = 3)
        {
            if (poblacion == null || poblacion.Count == 0) throw new ArgumentException("La población no puede estar vacía.");

            AgIndividuo mejor = null;

            for (int i = 0; i < tamanoTorneo; i++)
            {
                var aspirante = poblacion[_random.Next(poblacion.Count)];
                if (mejor == null || aspirante.Fitness > mejor.Fitness)
                {
                    mejor = aspirante;
                }
            }

            return mejor;
        }
    }
}
