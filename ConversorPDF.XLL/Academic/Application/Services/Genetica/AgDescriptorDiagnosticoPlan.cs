using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Resume el motivo real de un aula no verificable para el terminal y el log durable.
    /// Sólo describe códigos ya emitidos; nunca certifica origen, autoriza escritura ni inventa una causa.</summary>
    public static class AgDescriptorDiagnosticoPlan
    {
        /// <summary>Código de reserva cuando ningún diagnóstico tipado explica el bloqueo.</summary>
        public const string SinDiagnostico = "SIN_DIAGNOSTICO_CONOCIDO";

        private const string PrefijoAusente = "RELACION_AUSENTE:";

        /// <summary>Compone "RELACIONES_AUSENTES=ausentes/requeridas" seguido de hasta
        /// <paramref name="maxCodigos"/> códigos concretos. Los diagnósticos globales del lector se
        /// incluyen porque también dejan el aula en <see cref="AgEstadoValidacionPlan.PlanNoVerificable"/>.</summary>
        public static string Describir(AgValidacionAulaPlan aula, IEnumerable<string> diagnosticosGlobales = null,
            int maxCodigos = 3)
        {
            if (aula == null) throw new ArgumentNullException(nameof(aula));
            int limite = maxCodigos < 1 ? 1 : maxCodigos;
            int requeridas = aula.Relaciones.Count;
            int ausentes = aula.Relaciones.Count(r => r.Diagnosticos.Any(EsAusente));
            var codigos = aula.Relaciones.SelectMany(r => r.Diagnosticos)
                .Concat(diagnosticosGlobales ?? Enumerable.Empty<string>())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(c => c, StringComparer.Ordinal)
                .ToList();

            var partes = new List<string>();
            if (requeridas > 0)
                partes.Add("RELACIONES_AUSENTES=" + ausentes.ToString(CultureInfo.InvariantCulture) + "/" +
                    requeridas.ToString(CultureInfo.InvariantCulture));
            if (codigos.Count == 0)
            {
                partes.Add(SinDiagnostico);
                return string.Join(";", partes);
            }

            partes.AddRange(codigos.Take(limite));
            if (codigos.Count > limite)
                partes.Add("(+" + (codigos.Count - limite).ToString(CultureInfo.InvariantCulture) + ")");
            return string.Join(";", partes);
        }

        private static bool EsAusente(string codigo)
            => codigo != null && codigo.StartsWith(PrefijoAusente, StringComparison.Ordinal);
    }
}
