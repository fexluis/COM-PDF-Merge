namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Representa los posibles estados de conflicto entre dos sesiones para un docente.
    /// </summary>
    public enum EstadoConflicto
    {
        /// <summary>
        /// No hay conflicto de fechas ni horarios.
        /// </summary>
        Ninguno = 0,

        /// <summary>
        /// Las sesiones se solapan en hora en distintas sedes o superan la viabilidad (rojo).
        /// </summary>
        ConflictoReal = 1,

        /// <summary>
        /// Las sesiones ocurren el mismo día en sedes distintas, pero no se solapan en hora. 
        /// Requiere validación humana de tiempo de traslado (amarillo).
        /// </summary>
        RevisarViabilidad = 2,

        /// <summary>
        /// El usuario confirmó explícitamente que el tiempo de traslado es viable (verde).
        /// </summary>
        Salvado = 3
    }
}
