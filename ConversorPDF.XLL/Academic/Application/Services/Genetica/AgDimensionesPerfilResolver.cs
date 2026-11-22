using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resuelve y aplica de forma autoritativa las dimensiones del perfil de selección
    /// sobre las entidades AgAula en memoria y detecta desalineaciones (Fase 2303, Fase 2306).
    /// </summary>
    internal static class AgDimensionesPerfilResolver
    {
        internal static IReadOnlyList<AgAula> Aplicar(
            IReadOnlyList<AgAula> aulas,
            IReadOnlyList<AgContextoSeleccionDto> combinaciones,
            out IReadOnlyList<(string AulaRef, string Perfil, string Bd)> desalineaciones)
        {
            var diffs = new List<(string, string, string)>();
            if (aulas == null || combinaciones == null)
            {
                desalineaciones = diffs;
                return aulas ?? new List<AgAula>();
            }

            foreach (var aula in aulas)
            {
                var perfil = combinaciones.FirstOrDefault(c =>
                    string.Equals(c.AulaRef, aula.AulaRef, StringComparison.OrdinalIgnoreCase));
                // Fase 2309: Validación integral de dimensiones del perfil (días, jornada no vacía y horas válidas 4 u 8)
                bool perfilValido = perfil != null
                    && AgDiasVentana.Dias(perfil.DiasVentana).Count > 0
                    && !string.IsNullOrWhiteSpace(perfil.Jornada)
                    && (perfil.HorasDia == 4 || perfil.HorasDia == 8);

                if (!perfilValido) continue;

                if (!string.Equals(aula.DiasVentana, perfil.DiasVentana, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(aula.Jornada, perfil.Jornada, StringComparison.OrdinalIgnoreCase)
                    || aula.HorasDia != perfil.HorasDia)
                {
                    diffs.Add((aula.AulaRef, $"{perfil.Jornada}/{perfil.HorasDia}h/{perfil.DiasVentana}",
                               $"{aula.Jornada}/{aula.HorasDia}h/{aula.DiasVentana}"));
                    aula.DiasVentana = perfil.DiasVentana;
                    aula.Jornada = perfil.Jornada;
                    aula.HorasDia = perfil.HorasDia;
                }
            }
            desalineaciones = diffs;
            return aulas;
        }

        internal static IReadOnlyList<(string AulaRef, string Jornada, int HorasDia, string DiasVentana, string Subunidad)> Detectar(
            IReadOnlyList<AgAula> aulas,
            IReadOnlyList<AgContextoSeleccionDto> combinaciones)
        {
            var diffs = new List<(string AulaRef, string Jornada, int HorasDia, string DiasVentana, string Subunidad)>();
            if (aulas == null || combinaciones == null) return diffs;

            foreach (var aula in aulas)
            {
                var perfil = combinaciones.FirstOrDefault(c =>
                    string.Equals(c.AulaRef, aula.AulaRef, StringComparison.OrdinalIgnoreCase));
                // Fase 2309: Validación integral de dimensiones del perfil (días, jornada no vacía y horas válidas 4 u 8)
                bool perfilValido = perfil != null
                    && AgDiasVentana.Dias(perfil.DiasVentana).Count > 0
                    && !string.IsNullOrWhiteSpace(perfil.Jornada)
                    && (perfil.HorasDia == 4 || perfil.HorasDia == 8);

                if (!perfilValido) continue;

                if (!string.Equals(aula.DiasVentana, perfil.DiasVentana, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(aula.Jornada, perfil.Jornada, StringComparison.OrdinalIgnoreCase)
                    || aula.HorasDia != perfil.HorasDia)
                {
                    diffs.Add((aula.AulaRef, perfil.Jornada, perfil.HorasDia, perfil.DiasVentana, aula.Subunidad));
                }
            }
            return diffs;
        }
    }
}
