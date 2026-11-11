namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Diagnóstico de una corrida del AG que conserva staging (<c>ag_vinculaciones</c> / <c>ag_sesiones</c>)
    /// pero perdió su cabecera en <c>ag_corrida</c> y/o su individuo en <c>ag_individuo</c> (Fase 1566).
    /// Permite reconstruir la corrida sin volver a ejecutar el algoritmo genético.
    /// </summary>
    public sealed class AgCorridaHuerfana
    {
        /// <summary>Identificador de la corrida (ag_vinculaciones.id_corrida).</summary>
        public long IdCorrida { get; set; }

        /// <summary>Período derivado del staging (fuente de verdad cuando falta la cabecera).</summary>
        public int IdPeriodo { get; set; }

        /// <summary>Programa derivado del staging.</summary>
        public int IdPrograma { get; set; }

        /// <summary>Territorial derivada del staging.</summary>
        public string Subunidad { get; set; }

        /// <summary>Individuo referenciado por el staging (ag_vinculaciones.id_individuo).</summary>
        public long IdIndividuo { get; set; }

        /// <summary>Cantidad de vinculaciones materializadas para la corrida.</summary>
        public int TotalVinculaciones { get; set; }

        /// <summary>Cantidad de sesiones materializadas para la corrida.</summary>
        public int TotalSesiones { get; set; }

        /// <summary>Indica si la corrida conserva cabecera en ag_corrida.</summary>
        public bool ExisteCabecera { get; set; }

        /// <summary>Indica si el individuo del staging conserva fila en ag_individuo.</summary>
        public bool ExisteIndividuo { get; set; }

        /// <summary>
        /// Indica si existen los datos mínimos para reconstruir la corrida:
        /// sin cabecera, con staging y con individuo/período/programa identificables.
        /// </summary>
        public bool PuedeRepararse =>
            !ExisteCabecera && TotalVinculaciones > 0 && IdIndividuo > 0 && IdPeriodo > 0 && IdPrograma > 0;

        public AgCorridaHuerfana()
        {
            Subunidad = string.Empty;
        }

        public override string ToString() =>
            $"Corrida #{IdCorrida} · período {IdPeriodo} · programa {IdPrograma} · individuo {IdIndividuo} · " +
            $"{TotalVinculaciones} vinc / {TotalSesiones} ses · cabecera={ExisteCabecera} · individuo={ExisteIndividuo}";
    }
}
