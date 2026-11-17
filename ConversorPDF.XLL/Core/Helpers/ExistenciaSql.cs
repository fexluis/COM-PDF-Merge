namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Resultado tri-estado de comprobación de existencia de un elemento en el esquema SQL (Fase 1889).
    /// Permite diferenciar cuando un elemento fehacientemente no existe de cuando la comprobación no pudo realizarse.
    /// </summary>
    public enum ExistenciaSql
    {
        /// <summary>
        /// El elemento existe confirmadamente en el esquema.
        /// </summary>
        Existe,

        /// <summary>
        /// El elemento no existe en el esquema.
        /// </summary>
        NoExiste,

        /// <summary>
        /// No se pudo verificar la existencia (conexión cerrada, error de I/O, identificador inválido o excepción PRAGMA).
        /// </summary>
        Indeterminada
    }
}
