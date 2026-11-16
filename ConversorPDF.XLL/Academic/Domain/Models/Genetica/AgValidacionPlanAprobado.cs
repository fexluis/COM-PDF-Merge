using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Diagnóstico de origen, nunca autorización para conservar ni reemplazar.</summary>
    public enum AgEstadoValidacionPlan
    {
        /// <summary>Aula no aprobada: no requiere certificación productiva.</summary>
        NoRequerida,
        /// <summary>Lectura exitosa con identidad, sesiones o cobertura no verificables.</summary>
        PlanNoVerificable,
        /// <summary>Origen reconciliado; falta ratificación auditada T4, incluso si hay déficit.</summary>
        RequiereRatificacion
    }

    /// <summary>Cobertura diagnóstica por identidad completa; no compensa déficit entre relaciones.</summary>
    public sealed class AgValidacionRelacionPlan
    {
        /// <summary>Relación requerida del manifiesto.</summary>
        public AgClaveRelacionCobertura Clave { get; }
        /// <summary>Demanda efectiva vinculada a la ventana.</summary>
        public double HorasEsperadas { get; }
        /// <summary>Horas válidas del candidato reconciliado, todavía no ratificadas.</summary>
        public double HorasValidasEnVentana { get; }
        /// <summary>Cero: esta unidad no puede autorizar conservación sin T4.</summary>
        public double HorasPreservadas => 0;
        /// <summary>Horas lectivas finitas y positivas registradas fuera de la ventana.</summary>
        public double HorasFueraVentana { get; }
        /// <summary>Déficit respecto al candidato validado, no respecto a una autorización.</summary>
        public double Deficit => Math.Max(0, HorasEsperadas - HorasValidasEnVentana);
        /// <summary>Exceso respecto a la demanda efectiva.</summary>
        public double Exceso => Math.Max(0, HorasValidasEnVentana - HorasEsperadas);
        /// <summary>Códigos exactos y referencias canónicas para investigación.</summary>
        public IReadOnlyList<string> Diagnosticos { get; }
        /// <summary>Captura métricas sin cambiar sesiones ni conceder consentimiento.</summary>
        public AgValidacionRelacionPlan(AgRelacionCobertura requerida, double validas, double fuera, IEnumerable<string> diagnosticos)
        {
            Clave = requerida.Clave; HorasEsperadas = requerida.HorasEfectivas;
            HorasValidasEnVentana = validas; HorasFueraVentana = fuera;
            Diagnosticos = diagnosticos.Distinct().OrderBy(x => x, StringComparer.Ordinal).ToList().AsReadOnly();
        }
    }

    /// <summary>Diagnóstico completo de un aula, incluyendo relaciones ausentes.</summary>
    public sealed class AgValidacionAulaPlan
    {
        /// <summary>ID explícito, no deducido de la asignatura.</summary>
        public int IdAula { get; }
        /// <summary>Estado de validación; nunca permiso de escritura.</summary>
        public AgEstadoValidacionPlan Estado { get; }
        /// <summary>Todas las relaciones requeridas del aula.</summary>
        public IReadOnlyList<AgValidacionRelacionPlan> Relaciones { get; }
        /// <summary>Demanda total sin compensación.</summary>
        public double HorasEsperadas => Relaciones.Sum(r => r.HorasEsperadas);
        /// <summary>Conservación autorizada: cero en esta unidad.</summary>
        public double HorasPreservadas => Relaciones.Sum(r => r.HorasPreservadas);
        /// <summary>Déficits sumados individualmente.</summary>
        public double Deficit => Relaciones.Sum(r => r.Deficit);
        /// <summary>Excesos sumados individualmente.</summary>
        public double Exceso => Relaciones.Sum(r => r.Exceso);
        /// <summary>Horas fuera del plazo, sin moverlas.</summary>
        public double HorasFueraVentana => Relaciones.Sum(r => r.HorasFueraVentana);
        /// <summary>Copia el diagnóstico por aula.</summary>
        public AgValidacionAulaPlan(int idAula, AgEstadoValidacionPlan estado, IEnumerable<AgValidacionRelacionPlan> relaciones)
        {
            IdAula = idAula; Estado = estado; Relaciones = relaciones.ToList().AsReadOnly();
        }
    }

    /// <summary>Evidencia reproducible ligada al snapshot y su manifiesto, sin reemplazar una revisión durable.</summary>
    public sealed class AgValidacionPlanAprobado
    {
        /// <summary>Versión del diagnóstico serializado.</summary>
        public int Version => 1;
        /// <summary>Referencia exacta al contenido leído.</summary>
        public string HuellaOrigen { get; }
        /// <summary>Resultados por todas las aulas seleccionadas.</summary>
        public IReadOnlyList<AgValidacionAulaPlan> Aulas { get; }
        /// <summary>Incertidumbres globales del lector y claves no reconciliadas.</summary>
        public IReadOnlyList<string> Diagnosticos { get; }
        /// <summary>Las aprobadas siempre esperan T4; las disponibles no se bloquean por origen canónico.</summary>
        public bool BloqueaPreparacion => Aulas.Any(a => a.Estado != AgEstadoValidacionPlan.NoRequerida);
        /// <summary>Copia evidencia de validación sin elevarla a aprobación.</summary>
        public AgValidacionPlanAprobado(string huellaOrigen, IEnumerable<AgValidacionAulaPlan> aulas, IEnumerable<string> diagnosticos)
        {
            HuellaOrigen = huellaOrigen; Aulas = aulas.ToList().AsReadOnly();
            Diagnosticos = diagnosticos.Distinct().OrderBy(x => x, StringComparer.Ordinal).ToList().AsReadOnly();
        }
    }
}
