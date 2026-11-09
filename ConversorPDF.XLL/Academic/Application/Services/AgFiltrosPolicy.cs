using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Política de aplicación de los filtros por defecto de <c>config</c> (período y territorial).
    /// En MODO AG la corrida es la fuente de verdad: los filtros ocultos de configuración
    /// no deben aplicarse, porque ocultaban sesiones de corridas de otro período (Fase 1491).
    /// </summary>
    public static class AgFiltrosPolicy
    {
        /// <summary>
        /// Indica si se debe preseleccionar <c>config.id_periodo</c> en el combo de períodos.
        /// </summary>
        public static bool AplicarPeriodoConfig(AgModoDatos modo)
        {
            return modo == null || !modo.EsModoAg;
        }

        /// <summary>
        /// Indica si se debe preseleccionar <c>config.subunidad</c> en el combo de territoriales.
        /// </summary>
        public static bool AplicarSubunidadConfig(AgModoDatos modo)
        {
            return modo == null || !modo.EsModoAg;
        }
    }
}
