using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Reglas de presentación y selección del selector de corridas del listado de vinculaciones en MODO AG.
    /// Lógica pura (sin WinForms ni acceso a datos) para poder probarse de forma aislada (Fases 1492/1495).
    /// </summary>
    public static class AgCorridaSelector
    {
        /// <summary>Texto del ítem comodín que representa "todas las corridas".</summary>
        public const string EtiquetaTodasLasCorridas = "-- Todas las corridas --";

        private const string Separador = " · ";

        /// <summary>
        /// Construye la etiqueta mostrada en el ComboBox:
        /// <c>#44 · 2027-01 · ADMÓN. PÚBLICA TERRITORIAL · VALLE · 244 ses</c>.
        /// Los campos vacíos se omiten para no dejar separadores huérfanos.
        /// </summary>
        public static string FormatearEtiqueta(CorridaStagingItem corrida)
        {
            if (corrida == null) return string.Empty;

            var partes = new List<string>(6)
            {
                "#" + corrida.IdCorrida.ToString(CultureInfo.InvariantCulture)
            };

            AgregarSiTieneValor(partes, corrida.PeriodoNombre);
            AgregarSiTieneValor(partes, corrida.ProgramaNombre);
            AgregarSiTieneValor(partes, corrida.Subunidad);

            partes.Add(corrida.TieneStaging
                ? corrida.TotalSesiones.ToString(CultureInfo.InvariantCulture) + " ses"
                : "(sin sesiones)");

            var sufijoEstado = DescribirEstado(corrida.EstadoCorrida);
            if (!string.IsNullOrEmpty(sufijoEstado)) partes.Add(sufijoEstado);

            return string.Join(Separador, partes);
        }

        /// <summary>
        /// Ítem comodín con IdCorrida = 0 que representa "todas las corridas".
        /// </summary>
        public static CorridaStagingItem CrearTodasLasCorridas()
        {
            return new CorridaStagingItem
            {
                IdCorrida = 0,
                IdPeriodo = 0,
                IdPrograma = 0,
                PeriodoNombre = string.Empty,
                ProgramaNombre = string.Empty,
                Subunidad = string.Empty,
                FechaInicioCorrida = string.Empty,
                EstadoCorrida = 0,
                TotalVinculaciones = 0,
                TotalSesiones = 0,
                Etiqueta = EtiquetaTodasLasCorridas
            };
        }

        /// <summary>
        /// Prepara la colección que consume el ComboBox: antepone el comodín y puebla las etiquetas.
        /// Se ejecuta en la capa Application para que Infrastructure no dependa de Application.
        /// </summary>
        public static IReadOnlyCollection<CorridaStagingItem> PrepararParaCombo(IEnumerable<CorridaStagingItem> corridas)
        {
            var lista = new List<CorridaStagingItem> { CrearTodasLasCorridas() };

            if (corridas != null)
            {
                foreach (var corrida in corridas)
                {
                    if (corrida == null) continue;
                    corrida.Etiqueta = FormatearEtiqueta(corrida);
                    lista.Add(corrida);
                }
            }

            return lista.AsReadOnly();
        }

        /// <summary>
        /// Resuelve la selección inicial del ComboBox:
        /// 1) la corrida en uso (la cargada en el Módulo R), si existe en la lista;
        /// 2) la corrida más reciente que tenga staging;
        /// 3) el comodín "todas las corridas".
        /// </summary>
        public static CorridaStagingItem ResolverSeleccionInicial(
            IEnumerable<CorridaStagingItem> corridas,
            long? idCorridaEnUso)
        {
            var items = corridas?.Where(c => c != null).ToList() ?? new List<CorridaStagingItem>();

            if (idCorridaEnUso.HasValue && idCorridaEnUso.Value > 0)
            {
                var enUso = items.FirstOrDefault(c => c.IdCorrida == idCorridaEnUso.Value);
                if (enUso != null) return enUso;
            }

            var conStaging = items.FirstOrDefault(c => c.IdCorrida > 0 && c.TieneStaging);
            return conStaging ?? CrearTodasLasCorridas();
        }

        private static void AgregarSiTieneValor(ICollection<string> partes, string valor)
        {
            if (!string.IsNullOrWhiteSpace(valor)) partes.Add(valor.Trim());
        }

        private static string DescribirEstado(int estado)
        {
            if (estado == (int)AgEstadoCorrida.Error) return "Error";
            if (estado == (int)AgEstadoCorrida.Cancelada) return "Cancelada";
            return string.Empty;
        }
    }
}
