namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Proyección de una corrida del AG con el volumen real de su staging.
    /// Alimenta el selector de corridas del listado de vinculaciones en MODO AG (Fase 1488).
    /// Vive en Domain para que <c>IVinculacionQueryRepository</c> pueda retornarla sin violar la frontera de capas.
    /// </summary>
    public sealed class CorridaStagingItem
    {
        /// <summary>Identificador de la corrida (ag_corrida.id o ag_vinculaciones.id_corrida).</summary>
        public long IdCorrida { get; set; }

        /// <summary>Período académico al que pertenecen las vinculaciones de la corrida.</summary>
        public int IdPeriodo { get; set; }

        /// <summary>Nombre legible del período (p. ej. "2027-01").</summary>
        public string PeriodoNombre { get; set; }

        /// <summary>Programa académico de la corrida.</summary>
        public int IdPrograma { get; set; }

        /// <summary>Nombre legible del programa.</summary>
        public string ProgramaNombre { get; set; }

        /// <summary>Territorial (subunidad) dueña del staging.</summary>
        public string Subunidad { get; set; }

        /// <summary>Marca temporal de inicio registrada en ag_corrida.</summary>
        public string FechaInicioCorrida { get; set; }

        /// <summary>Estado de la corrida (ver <c>AgEstadoCorrida</c>).</summary>
        public int EstadoCorrida { get; set; }

        /// <summary>Cantidad de vinculaciones materializadas en staging.</summary>
        public int TotalVinculaciones { get; set; }

        /// <summary>Cantidad de sesiones materializadas en staging.</summary>
        public int TotalSesiones { get; set; }

        /// <summary>Indica si la corrida tiene sesiones visibles en staging.</summary>
        public bool TieneStaging => TotalSesiones > 0;

        /// <summary>
        /// Texto mostrado en el ComboBox. No se mapea por Dapper:
        /// la puebla <c>AgCorridaSelector</c> (ver <c>VinculacionListService</c>).
        /// </summary>
        public string Etiqueta { get; set; }

        /// <summary>Representación por defecto para enlaces de datos.</summary>
        public override string ToString() => Etiqueta ?? string.Empty;
    }
}
