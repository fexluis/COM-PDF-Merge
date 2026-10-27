using System;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Registra una infracción o conflicto detectado en una solución del AG frente a una regla dura o blanda.
    /// Mapea a la tabla 'ag_violacion'.
    /// </summary>
    public class AgViolacion
    {
        public long Id { get; set; }
        public long IdIndividuo { get; set; }
        public long? IdGen { get; set; }
        public AgReglaViolacion Regla { get; set; }
        public bool EsDura { get; set; }
        public string Mensaje { get; set; }
        public string DetalleJson { get; set; }
        public string Subunidad { get; set; }

        public AgViolacion()
        {
            Mensaje = string.Empty;
            DetalleJson = "{}";
            Subunidad = string.Empty;
        }

        public AgViolacion(AgReglaViolacion regla, bool esDura, string mensaje, long? idGen = null, string detalleJson = "{}")
        {
            Regla = regla;
            EsDura = esDura;
            Mensaje = mensaje ?? string.Empty;
            IdGen = idGen;
            DetalleJson = detalleJson ?? "{}";
            Subunidad = string.Empty;
        }

        public override string ToString()
        {
            string tipo = EsDura ? "HARD" : "SOFT";
            return $"[{tipo}] {Regla}: {Mensaje}";
        }
    }
}
