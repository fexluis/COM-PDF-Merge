using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Propuesta explícita y verificable de reclasificación de patrón semanal (Fase 1750).
    /// Acotada a la identidad exacta del aula (IdAula, AulaRef, Subunidad) y asignatura (IdAsignatura).
    /// Permite transicionar patrones LEGADO hacia AUTOMATICO o FIJO sin conversiones masivas no autorizadas.
    /// </summary>
    public sealed class AgReclasificacionPatron
    {
        public int IdAula { get; set; }
        public string AulaRef { get; set; } = string.Empty;
        public string Subunidad { get; set; } = string.Empty;
        public int IdAsignatura { get; set; }
        public string NombreAsignatura { get; set; } = string.Empty;
        public string OrigenAnterior { get; set; } = "LEGADO";
        public string OrigenPropuesto { get; set; } = "AUTOMATICO";
        public string DiasPatronAnterior { get; set; } = string.Empty;
        public string DiasPatronPropuesto { get; set; } = string.Empty;
        public string Motivo { get; set; } = string.Empty;
        public bool Autorizado { get; set; }

        /// <summary>Valida identidad y transición explícita desde LEGADO, conservando los patrones FIJO.</summary>
        public bool EsTransicionValida(out string error)
        {
            error = string.Empty;
            string antes = (OrigenAnterior ?? string.Empty).Trim().ToUpperInvariant();
            string despues = (OrigenPropuesto ?? string.Empty).Trim().ToUpperInvariant();

            if (antes != "LEGADO")
            {
                error = $"La asignatura {IdAsignatura} tiene origen {antes}; solo se reclasifica LEGADO.";
                return false;
            }
            if (string.IsNullOrEmpty(despues) || (despues != "AUTOMATICO" && despues != "FIJO"))
            {
                error = $"Origen propuesto '{OrigenPropuesto}' no es válido. Solo se admite AUTOMATICO o FIJO.";
                return false;
            }
            if (antes == despues)
            {
                error = $"Origen propuesto '{despues}' es idéntico al origen actual '{antes}'.";
                return false;
            }
            if (IdAsignatura <= 0 || IdAula <= 0 || string.IsNullOrWhiteSpace(AulaRef))
            {
                error = "Identidad de aula o asignatura inválida.";
                return false;
            }
            if (despues == "AUTOMATICO" && !string.IsNullOrEmpty(DiasPatronPropuesto))
            {
                error = "AUTOMATICO debe liberar el patrón semanal.";
                return false;
            }
            if (despues == "FIJO" && (string.IsNullOrWhiteSpace(DiasPatronAnterior)
                || !string.Equals(DiasPatronAnterior, DiasPatronPropuesto, StringComparison.Ordinal)))
            {
                error = "FIJO debe conservar el patrón anterior explícito.";
                return false;
            }
            return true;
        }

        /// <summary>Copia la propuesta para que el sellado no comparta objetos con la selección.</summary>
        public AgReclasificacionPatron Clone()
        {
            return new AgReclasificacionPatron
            {
                IdAula = this.IdAula,
                AulaRef = this.AulaRef,
                Subunidad = this.Subunidad,
                IdAsignatura = this.IdAsignatura,
                NombreAsignatura = this.NombreAsignatura,
                OrigenAnterior = this.OrigenAnterior,
                OrigenPropuesto = this.OrigenPropuesto,
                DiasPatronAnterior = this.DiasPatronAnterior,
                DiasPatronPropuesto = this.DiasPatronPropuesto,
                Motivo = this.Motivo,
                Autorizado = this.Autorizado
            };
        }
    }
}
