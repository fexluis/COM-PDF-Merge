namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Ciclo de vida y estado de ejecución de una corrida del Algoritmo Genético.
    /// </summary>
    public enum AgEstadoCorrida
    {
        Iniciada = 0,
        EnProgreso = 1,
        Completada = 2,
        Cancelada = 3,
        Error = 4,
        Aprobada = 5,
        /// <summary> Corrida finalizada cuya mejor solución contiene violaciones duras insatisfechas (Fase 1978). </summary>
        SinFactibilidad = 6
    }
}
