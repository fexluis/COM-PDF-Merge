namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resumen de diagnóstico del staging de una corrida (Fase 1489).
    /// Separa deliberadamente lo que la corrida contiene (sin filtros de UI) de lo que
    /// devuelve la consulta con los filtros activos, para poder explicar al usuario
    /// por qué "hay sesiones en staging pero 0 coinciden con los filtros".
    /// </summary>
    public sealed class AgStagingResumen
    {
        /// <summary>Corrida consultada.</summary>
        public long IdCorrida { get; set; }

        /// <summary>Sesiones de la corrida, sin aplicar filtros de la UI.</summary>
        public int SesionesCorrida { get; set; }

        /// <summary>Sesiones de la corrida que sobreviven a los filtros activos de la UI.</summary>
        public int SesionesConFiltros { get; set; }

        /// <summary>Vinculaciones distintas de la corrida (sin filtros de la UI).</summary>
        public int VincCorrida { get; set; }

        /// <summary>Períodos reales de la corrida, separados por coma (p. ej. "2027-01").</summary>
        public string PeriodosCorrida { get; set; }

        /// <summary>Territoriales reales de la corrida, separadas por coma (p. ej. "VALLE").</summary>
        public string SubunidadesCorrida { get; set; }

        /// <summary>Programas reales de la corrida, separados por coma.</summary>
        public string ProgramasCorrida { get; set; }
    }
}
