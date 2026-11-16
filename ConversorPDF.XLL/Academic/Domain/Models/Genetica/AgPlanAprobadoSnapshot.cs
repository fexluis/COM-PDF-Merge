using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Sesión productiva inmutable; su presencia no prueba aprobación.</summary>
    public sealed class AgSesionCanonica
    {
        /// <summary>ID canónico, no ID de staging.</summary>
        public long IdSesion { get; }
        /// <summary>Vinculación canónica de origen.</summary>
        public long IdVinculacion { get; }
        /// <summary>Docente registrado en la sesión.</summary>
        public long IdTercero { get; }
        /// <summary>Fecha interpretada, null si el origen no es verificable.</summary>
        public DateTime? Fecha { get; }
        /// <summary>Inicio interpretado sin inventar horario ante valores legados.</summary>
        public TimeSpan? HoraInicio { get; }
        /// <summary>Fin interpretado del intervalo ocupado.</summary>
        public TimeSpan? HoraFin { get; }
        /// <summary>Horas lectivas independientes del intervalo ocupado.</summary>
        public double HorasLectivas { get; }
        /// <summary>Partición literal registrada en la sesión.</summary>
        public string Subunidad { get; }
        /// <summary>Contenido completo de la fila de origen, codificado de forma invariante.</summary>
        public string OrigenCanonico { get; }
        /// <summary>El esquema actual no ofrece un estado de aprobación productiva.</summary>
        public string EstadoAprobacion => null;
        /// <summary>El esquema actual no ofrece revisión de aprobación; una huella no la sustituye.</summary>
        public string RevisionAprobacion => null;
        /// <summary>Copia una sesión leída; no concede ratificación ni aprobación.</summary>
        public AgSesionCanonica(long idSesion, long idVinculacion, long idTercero, DateTime? fecha,
            TimeSpan? horaInicio, TimeSpan? horaFin, double horasLectivas, string subunidad, string origenCanonico)
        {
            IdSesion = idSesion; IdVinculacion = idVinculacion; IdTercero = idTercero; Fecha = fecha;
            HoraInicio = horaInicio; HoraFin = horaFin; HorasLectivas = horasLectivas;
            Subunidad = subunidad; OrigenCanonico = origenCanonico ?? throw new ArgumentNullException(nameof(origenCanonico));
        }
    }

    /// <summary>Vinculación productiva, incluso sin sesiones o sin correspondencia única con el manifiesto.</summary>
    public sealed class AgRelacionPlanLeida
    {
        /// <summary>ID canónico de la vinculación.</summary>
        public long IdVinculacion { get; }
        /// <summary>Claves requeridas compatibles; vacía para extras o identidad no reconciliada.</summary>
        public IReadOnlyList<AgClaveRelacionCobertura> ClavesCandidatas { get; }
        /// <summary>Todas las sesiones canónicas, sin filtro por fecha.</summary>
        public IReadOnlyList<AgSesionCanonica> Sesiones { get; }
        /// <summary>Fila completa de vinculación, incluyendo campos no proyectados al DTO.</summary>
        public string OrigenCanonico { get; }
        /// <summary>Copia la vinculación y sus sesiones sin seleccionar arbitrariamente una coincidencia.</summary>
        public AgRelacionPlanLeida(long idVinculacion, IEnumerable<AgClaveRelacionCobertura> claves,
            IEnumerable<AgSesionCanonica> sesiones, string origenCanonico)
        {
            IdVinculacion = idVinculacion;
            ClavesCandidatas = claves.OrderBy(k => k.Valor, StringComparer.Ordinal).ToList().AsReadOnly();
            Sesiones = sesiones.OrderBy(s => s.IdSesion).ToList().AsReadOnly();
            OrigenCanonico = origenCanonico ?? throw new ArgumentNullException(nameof(origenCanonico));
        }
    }

    /// <summary>Resultado de lectura, no certificado de cobertura ni permiso de escritura.</summary>
    public sealed class AgPlanAprobadoSnapshot
    {
        /// <summary>Manifiesto esperado, incluso donde producción está ausente.</summary>
        public AgManifiestoCobertura Manifiesto { get; }
        /// <summary>Vinculaciones leídas, incluidas extras y ambiguas.</summary>
        public IReadOnlyList<AgRelacionPlanLeida> Relaciones { get; }
        /// <summary>Requeridas sin ninguna vinculación reconciliada; distinto de vinculación sin sesiones.</summary>
        public IReadOnlyList<AgClaveRelacionCobertura> RelacionesAusentes { get; }
        /// <summary>Diagnósticos estructurales; el validador futuro debe bloquear extras/ambigüedad.</summary>
        public IReadOnlyList<string> Diagnosticos { get; }
        /// <summary>Huella completa de manifiesto, filas productivas y catálogo de reconciliación.</summary>
        public string HuellaOrigen { get; }
        /// <summary>Siempre exige ratificación si se reconcilia; nunca autoaprueba un origen histórico/manual.</summary>
        public string EstadoOrigen => Diagnosticos.Count > 0 || RelacionesAusentes.Count > 0
            ? "PLAN_NO_VERIFICABLE" : "REQUIERE_RATIFICACION";
        /// <summary>Construye un snapshot reproducible; la ratificación durable pertenece a T4.</summary>
        public AgPlanAprobadoSnapshot(AgManifiestoCobertura manifiesto, IEnumerable<AgRelacionPlanLeida> relaciones,
            IEnumerable<string> diagnosticos, string origenCatalogo)
        {
            Manifiesto = manifiesto ?? throw new ArgumentNullException(nameof(manifiesto));
            Relaciones = relaciones.OrderBy(r => r.IdVinculacion).ToList().AsReadOnly();
            RelacionesAusentes = manifiesto.DetectarOmisiones(Relaciones.SelectMany(r => r.ClavesCandidatas));
            Diagnosticos = diagnosticos.Distinct(StringComparer.Ordinal).OrderBy(d => d, StringComparer.Ordinal).ToList().AsReadOnly();
            HuellaOrigen = AgOrigenCanonico.Huella(AgOrigenCanonico.Campos(1, manifiesto.Huella, origenCatalogo) +
                string.Concat(Relaciones.Select(r => AgOrigenCanonico.Campos(r.IdVinculacion, r.OrigenCanonico,
                    string.Concat(r.ClavesCandidatas.Select(k => AgOrigenCanonico.Campos(k.Valor))),
                    string.Concat(r.Sesiones.Select(s => AgOrigenCanonico.Campos(s.IdSesion, s.OrigenCanonico)))))));
        }
    }
}
