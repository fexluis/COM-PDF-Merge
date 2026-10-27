using System;
using System.Globalization;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Entidad de dominio que representa la habilitación e historial de asignación
    /// de una asignatura específica para un docente/tercero dentro de un programa y semestre.
    /// Forma parte del pool docente (Capa 0) y cuenta con soporte para sincronización CRDT.
    /// </summary>
    public class TerceroAsignatura
    {
        private string _semestre = string.Empty;

        /// <summary>
        /// Clave primaria única determinista de 63 bits (calculada mediante <see cref="TerceroAsignaturaKey"/>).
        /// NOT NULL y sin AUTOINCREMENT para total compatibilidad CRDT.
        /// </summary>
        public long Id { get; set; }

        /// <summary>
        /// Identificador del tercero/docente.
        /// </summary>
        public int IdTercero { get; set; }

        /// <summary>
        /// Identificador del programa académico.
        /// </summary>
        public int IdPrograma { get; set; }

        /// <summary>
        /// Identificador de la asignatura.
        /// </summary>
        public int IdAsignatura { get; set; }

        /// <summary>
        /// Semestre lectivo canónico (I..X) o cadena vacía si es transversal.
        /// Siempre sanitizado en mayúsculas sin espacios extremos.
        /// </summary>
        public string Semestre
        {
            get => _semestre;
            set => _semestre = (value ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>
        /// 1 = habilitación activa, 0 = inactiva (columna INTEGER NOT NULL DEFAULT 1).
        /// Sigue la convención del dominio del proyecto (ej. Tercero, Sede).
        /// </summary>
        public int Activo { get; set; } = 1;

        /// <summary>
        /// Prioridad de asignación para el algoritmo genético (0 = normal, >0 = preferente).
        /// </summary>
        public int Prioridad { get; set; }

        /// <summary>
        /// Número de veces que el docente ha dictado esta asignatura históricamente.
        /// </summary>
        public int VecesAsignada { get; set; }

        /// <summary>
        /// Fecha de la última asignación en formato ISO yyyy-MM-dd.
        /// </summary>
        public string FechaUltimaAsignacion { get; set; } = string.Empty;

        /// <summary>
        /// Asigna la fecha de última asignación asegurando formato ISO yyyy-MM-dd.
        /// Si fecha es null, establece cadena vacía.
        /// </summary>
        /// <param name="fecha">Fecha a asignar o null.</param>
        public void EstablecerFechaUltimaAsignacion(DateTime? fecha)
        {
            FechaUltimaAsignacion = fecha.HasValue
                ? fecha.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        /// <summary>
        /// Intenta parsear la fecha de última asignación almacenada en formato ISO yyyy-MM-dd.
        /// </summary>
        /// <param name="fecha">Fecha parseada de salida.</param>
        /// <returns>True si la fecha es válida; false si está vacía o es inválida.</returns>
        public bool TryObtenerFechaUltimaAsignacion(out DateTime fecha)
        {
            return DateTime.TryParseExact(
                FechaUltimaAsignacion,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out fecha);
        }

        /// <summary>
        /// Origen del registro: 'MANUAL', 'HISTORICO' o 'AG_APROBADO'.
        /// </summary>
        public string Origen { get; set; } = "MANUAL";

        /// <summary>
        /// Observaciones o notas administrativas sobre la habilitación.
        /// </summary>
        public string Observaciones { get; set; } = string.Empty;

        /// <summary>
        /// Identificador de la subunidad territorial para aislamiento CRDT.
        /// </summary>
        public string Subunidad { get; set; } = string.Empty;

        /// <summary>
        /// Marca de tiempo de creación del registro.
        /// </summary>
        public string CreatedAt { get; set; } = string.Empty;

        /// <summary>
        /// Marca de tiempo de última actualización del registro.
        /// </summary>
        public string UpdatedAt { get; set; } = string.Empty;
    }
}
