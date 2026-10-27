using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Construye diagnósticos accionables ante aulas que carecen de dimensiones de decisión
    /// (Jornada, Días de Ventana, Horas/Día) persistidas en ag_aula (Fase 1192 / Fase 2).
    /// </summary>
    public static class AgMensajeDimensionesAula
    {
        /// <summary>
        /// Determina las dimensiones faltantes o no configuradas para un aula dada.
        /// </summary>
        public static IReadOnlyList<string> Faltantes(AgAula aula)
        {
            var faltantes = new List<string>();
            if (aula == null)
            {
                faltantes.Add("días de ventana");
                faltantes.Add("jornada");
                faltantes.Add("horas/día");
                return faltantes;
            }

            if (string.IsNullOrWhiteSpace(aula.DiasVentana))
            {
                faltantes.Add("días de ventana");
            }
            if (string.IsNullOrWhiteSpace(aula.Jornada))
            {
                faltantes.Add("jornada");
            }
            if (aula.HorasDia <= 0)
            {
                faltantes.Add("horas/día");
            }

            return faltantes;
        }

        /// <summary>
        /// Genera el mensaje descriptivo y accionable indicando las acciones correctivas
        /// que el usuario debe ejecutar en el árbol de selección y base de datos.
        /// </summary>
        public static string Construir(AgAula aula, string errorValidador)
        {
            var faltantes = Faltantes(aula);
            if (faltantes.Count > 0)
            {
                string detalle = string.Join(", ", faltantes);
                return $"Aula '{aula?.AulaRef}': sin {detalle} configurados en base de datos.\nSeleccione las dimensiones en el árbol (Jornada / Días / Horas) y ejecute 'Generar Aulas' antes de continuar.";
            }

            return $"Aula '{aula?.AulaRef}': {errorValidador}";
        }
    }
}
