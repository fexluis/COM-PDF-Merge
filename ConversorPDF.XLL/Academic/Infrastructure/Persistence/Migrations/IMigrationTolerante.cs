namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Interfaz que identifica migraciones de carácter aditivo o no destructivo.
    /// Si una migración tolerante falla, no aborta el arranque general del add-in ni bloquea las funciones del Ribbon.
    /// </summary>
    public interface IMigrationTolerante
    {
        bool Aditiva { get; }
    }
}
