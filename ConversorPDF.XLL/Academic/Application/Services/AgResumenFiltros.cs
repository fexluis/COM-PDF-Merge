namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Formato de los resúmenes de filtros del generador AG (Fase 1472, continuacion de Fase 1469).
    /// </summary>
    internal static class AgResumenFiltros
    {
        public const string Todos = "Todos";

        /// <summary>
        /// Resumen de programas: "Todos" solo cuando no hay seleccion;
        /// si la modalidad oculta parte de la seleccion se informa "visibles (+ocultos otra modalidad)".
        /// </summary>
        public static string Programas(int seleccionados, int visibles)
        {
            if (seleccionados <= 0) return Todos;

            int visiblesSeguros = visibles < 0 ? 0 : visibles;
            if (visiblesSeguros >= seleccionados) return seleccionados.ToString();

            return $"{visiblesSeguros} (+{seleccionados - visiblesSeguros} otra modalidad)";
        }
    }
}
