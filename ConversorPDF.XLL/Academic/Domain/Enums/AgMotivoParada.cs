namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Causa o condición por la cual finalizó el bucle evolutivo del Algoritmo Genético.
    /// </summary>
    public enum AgMotivoParada
    {
        /// <summary> La corrida no ha finalizado o no definió motivo. </summary>
        NoDefinida = 0,

        /// <summary> Se alcanzó el umbral de aptitud / fitness configurado con 0 violaciones duras. </summary>
        Umbral = 1,

        /// <summary> Se superó el número máximo de generaciones consecutivas sin mejora del fitness. </summary>
        Estancamiento = 2,

        /// <summary> Se completó el límite total de generaciones programadas sin alcanzar otro criterio de corte. </summary>
        MaxGeneraciones = 3
    }
}
