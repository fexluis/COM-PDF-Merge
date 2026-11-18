using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio desacoplado para el cálculo de viáticos de sesiones académicas.
    /// Reglas:
    /// - Modalidad Plataforma (o legada Virtual): nunca viatica.
    /// - Misma ciudad de residencia y sede: no viatica.
    /// - Presencial con desplazamiento: siempre viatica.
    /// - Mixta (o legada Semipresencial / A Distancia) con desplazamiento: viatica solo si la fecha cae dentro de un rango presencial.
    /// </summary>
    public static class CalculadoraViaticaSesion
    {
        public static bool DeterminarViatica(
            DateTime fechaSesion,
            string modalidad,
            string sedeCiudad,
            string residenciaDocente,
            IEnumerable<(DateTime Inicio, DateTime Fin)> rangosPresenciales)
        {
            if (string.IsNullOrWhiteSpace(modalidad)) return false;

            if (ModalidadTutoriaCatalog.EsPlataforma(modalidad)) return false;

            string normSede = NormalizarCiudad(sedeCiudad);
            string normResidencia = NormalizarCiudad(residenciaDocente);

            if (!string.IsNullOrEmpty(normSede) && !string.IsNullOrEmpty(normResidencia) && normSede == normResidencia)
            {
                return false;
            }

            if (ModalidadTutoriaCatalog.EsPresencial(modalidad))
            {
                return true;
            }

            if (ModalidadTutoriaCatalog.EsMixta(modalidad))
            {
                if (rangosPresenciales == null) return false;

                var f = fechaSesion.Date;
                foreach (var r in rangosPresenciales)
                {
                    if (f >= r.Inicio.Date && f <= r.Fin.Date)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static readonly ConcurrentDictionary<string, string> _ciudadesNormalizadas =
            new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static string NormalizarCiudad(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return string.Empty;
            return _ciudadesNormalizadas.GetOrAdd(ciudad, NormalizarCiudadInterno);
        }

        private static string NormalizarCiudadInterno(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return string.Empty;

            var normalized = ciudad.ToUpperInvariant().Replace(".", "").Trim();
            var decomposed = normalized.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var c in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            string result = sb.ToString();
            while (result.Contains("  "))
            {
                result = result.Replace("  ", " ");
            }

            result = result.Trim();
            if (result.EndsWith(" DC"))
            {
                result = result.Substring(0, result.Length - 3).Trim();
            }

            return result;
        }
    }
}
