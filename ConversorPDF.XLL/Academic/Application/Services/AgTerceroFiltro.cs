using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ConversorPDF.XLL.Academic.Application.DTOs;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio puro para filtrado en memoria de docentes/terceros (Fase 1319 / 1329).
    /// </summary>
    internal static class AgTerceroFiltro
    {
        public static string NormalizarTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return string.Empty;
            string normalizado = texto.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char c in normalizado)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(c);
                if (uc != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC).Trim().ToLowerInvariant();
        }

        public static bool CoincideTerritorial(string departamentoDocente, string territorialFiltro)
        {
            if (string.IsNullOrWhiteSpace(territorialFiltro) || territorialFiltro.StartsWith("-"))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(departamentoDocente))
            {
                return false;
            }

            var partes = departamentoDocente.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(p => p.Trim());

            return partes.Any(p => string.Equals(p, territorialFiltro, StringComparison.OrdinalIgnoreCase));
        }

        public static List<AgTerceroSeleccionDto> Filtrar(
            IEnumerable<AgTerceroSeleccionDto> origen,
            string territorial,
            string busqueda,
            bool soloHabilitadas,
            bool mostrarInactivos)
        {
            if (origen == null) return new List<AgTerceroSeleccionDto>();

            string searchNorm = NormalizarTexto(busqueda);
            var consulta = origen.Where(t =>
                t.Cedula != "1111111111" &&
                !string.Equals(t.Nombre?.Trim(), "VACANTE", StringComparison.OrdinalIgnoreCase));

            if (!mostrarInactivos)
            {
                consulta = consulta.Where(t => t.Activo);
            }

            if (!string.IsNullOrWhiteSpace(territorial) && !territorial.StartsWith("-") &&
                !string.Equals(territorial, AgTerritorialCatalog.Todas, StringComparison.OrdinalIgnoreCase))
            {
                // Fase 1471: el filtro aplica sobre Departamento y Subunidad; los valores compuestos
                // ("BOLIVAR,VALLE") coinciden con cada una de sus partes.
                consulta = consulta.Where(t =>
                    CoincideTerritorial(t.Departamento, territorial) ||
                    CoincideTerritorial(t.Subunidad, territorial));
            }

            if (!string.IsNullOrEmpty(searchNorm))
            {
                consulta = consulta.Where(t =>
                    NormalizarTexto(t.Nombre).Contains(searchNorm) ||
                    NormalizarTexto(t.DisplayName).Contains(searchNorm) ||
                    NormalizarTexto(t.Departamento).Contains(searchNorm) ||
                    NormalizarTexto(t.Subunidad).Contains(searchNorm) ||
                    (t.Cedula != null && t.Cedula.Contains(searchNorm)));
            }

            if (soloHabilitadas)
            {
                // Fase 1471: "habilitadas" = habilitación en la subunidad activa (lo que realmente consume el AG).
                consulta = consulta.Where(t => t.TieneAsignaturasHabilitadas);
            }

            return consulta.ToList();
        }
    }
}
