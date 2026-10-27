namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Consultas canónicas del catálogo de programas para el módulo AG (Fase 1466 &amp; Fase 1467).
    /// Se alinean con <c>ProgramaLookupRepository.GetProgramasLookupAsync</c> (Vinculaciones):
    /// se listan todos los programas, sin filtrar por <c>activo</c>, porque el nivel académico
    /// (PREGRADO / POSGRADO) se resuelve en memoria con <c>AgTipoProgramaCatalog</c>.
    /// Se aplican defensas COALESCE en campos numéricos para blindar el mapeo ante valores nulos.
    /// </summary>
    internal static class AgProgramaCatalogoSql
    {
        /// <summary>
        /// Selecciona el catálogo completo de programas, deduplicable por <c>Id</c> en el llamador.
        /// Orden: activos primero (los nulos se tratan como inactivos) y luego por etiqueta compacta
        /// (abreviatura con fallback a <c>nombreprograma</c>).
        /// </summary>
        public const string SeleccionProgramas =
            @"SELECT id as Id, nombreprograma as NombrePrograma, abreviatura as Abreviatura,
                     COALESCE(pensum, 0) as Pensum, COALESCE(actual, 0) as Actual,
                     COALESCE(activo, 0) as Activo, modalidad as Modalidad
              FROM programas
              ORDER BY COALESCE(activo, 0) DESC, COALESCE(NULLIF(abreviatura, ''), nombreprograma)";
    }
}
