namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Decisión de continuidad del pre-vuelo ante infactibilidad de ventana (Fase 2127).
    /// </summary>
    public static class AgPreVueloEscalado
    {
        /// <summary>
        /// Sólo el rechazo explícito del escalado, seguido de "continuar", autoriza el override.
        /// </summary>
        public static bool ContinuarPeseAInfactibilidad(AgResultadoEscaladoUI resultado) =>
            resultado == AgResultadoEscaladoUI.RechazadoPorUsuario;
    }
}
