namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    public sealed class AgTerceroSeleccionDto
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public string Residencia { get; set; } = string.Empty;
        public string DisplayName => !string.IsNullOrWhiteSpace(Residencia) ? $"{Nombre} ({Residencia})" : Nombre;
        public string Cedula { get; set; } = string.Empty;
        public string Tipo { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
        public string Departamento { get; set; } = string.Empty;
        public string Subunidad { get; set; } = string.Empty;
        public bool Activo { get; set; } = true;
        /// <summary>Habilitaciones activas en la subunidad activa (misma regla que el pool AG). Fase 1471.</summary>
        public int HabilitacionesLocales { get; set; }

        /// <summary>Habilitaciones activas en otras subunidades (informativo). Fase 1471.</summary>
        public int HabilitacionesOtras { get; set; }

        public int TotalHabilitaciones
        {
            get => HabilitacionesLocales + HabilitacionesOtras;
            set
            {
                HabilitacionesLocales = value;
                HabilitacionesOtras = 0;
            }
        }

        /// <summary>
        /// Semántica AG: solo cuenta la habilitación de la subunidad activa, porque el pool de generación
        /// está acotado a esa subunidad (AgFormGeneradorHorarios L912-936).
        /// </summary>
        public bool TieneAsignaturasHabilitadas
        {
            get => HabilitacionesLocales > 0;
            set => HabilitacionesLocales = value ? 1 : 0;
        }
        public bool Seleccionado { get; set; }
    }
}
