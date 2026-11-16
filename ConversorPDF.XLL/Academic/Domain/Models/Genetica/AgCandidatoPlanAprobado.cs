using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Autorización inmutable de una revisión candidata; nunca modifica el horario productivo.</summary>
    public sealed class AgCandidatoPlanAprobado
    {
        public Guid Operacion { get; }
        public Guid Intento { get; }
        public int IdAula { get; }
        public string Actor { get; }
        public string Motivo { get; }
        public string Alcance { get; }
        public string HuellaManifiesto { get; }
        public string HuellaOrigen { get; }
        public string HuellaPrevia { get; }
        public string HuellaEfecto { get; }
        public string Desde { get; }
        public string Hacia { get; }
        public string Consentimiento { get; }
        public DateTime Inicio { get; }
        public DateTime Fin { get; }

        /// <summary>Construye el efecto independiente de la huella previa; consentimiento son IDs de LEGADO.</summary>
        public AgCandidatoPlanAprobado(Guid operacion, Guid intento, int idAula, string actor, string motivo,
            AgManifiestoCobertura manifiesto, string huellaOrigen, string huellaPrevia, string desde, string hacia,
            IEnumerable<int> consentimiento)
        {
            if (operacion == Guid.Empty || intento == Guid.Empty || idAula <= 0 || manifiesto == null
                || string.IsNullOrWhiteSpace(actor) || string.IsNullOrWhiteSpace(motivo)
                || string.IsNullOrWhiteSpace(huellaOrigen) || string.IsNullOrWhiteSpace(huellaPrevia)
                || string.IsNullOrWhiteSpace(desde) || string.IsNullOrWhiteSpace(hacia))
                throw new ArgumentException("CANDIDATO_INCOMPLETO");
            var ids = (consentimiento ?? Enumerable.Empty<int>()).ToList();
            if (ids.Any(i => i <= 0) || ids.Count != ids.Distinct().Count()) throw new ArgumentException("CONSENTIMIENTO_INVALIDO");
            Operacion = operacion; Intento = intento; IdAula = idAula; Actor = actor.Trim(); Motivo = motivo.Trim();
            HuellaManifiesto = manifiesto.Huella; HuellaOrigen = huellaOrigen; HuellaPrevia = huellaPrevia;
            Inicio = manifiesto.Inicio; Fin = manifiesto.Fin; Desde = desde; Hacia = hacia;
            Alcance = string.Concat(manifiesto.Relaciones.Where(r => r.Clave.IdAula == idAula)
                .Select(r => AgOrigenCanonico.Campos(r.Clave.Valor)));
            Consentimiento = string.Join(",", ids.OrderBy(i => i));
            HuellaEfecto = AgOrigenCanonico.Huella(AgOrigenCanonico.Campos(Operacion, Intento, IdAula,
                Actor, Motivo, Alcance, HuellaManifiesto, HuellaOrigen, HuellaPrevia, Desde, Hacia,
                Consentimiento, Inicio, Fin));
        }
    }
}
