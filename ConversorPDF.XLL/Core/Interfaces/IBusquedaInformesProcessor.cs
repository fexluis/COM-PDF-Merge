using System;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Procesa los informes de viáticos alojados en &lt;carpeta del libro&gt;\Informes:
    /// renombrado al estándar, unificación de parejas INF + CUM y auditoría de la
    /// columna INFORME de la hoja CONSULTAS.
    /// </summary>
    public interface IBusquedaInformesProcessor
    {
        /// <summary>
        /// Recorre la carpeta de informes (recursivo), renombra los PDF al estándar
        /// <c>{solicitud}-{cedula}-{tipo}-{nombre}</c> y, en las carpetas especiales,
        /// unifica las parejas INF + CUM en un archivo UNI.
        /// </summary>
        /// <param name="carpetaRaiz">Carpeta \Informes del libro.</param>
        /// <param name="inicio">Solicitud inicial del rango (0 = sin filtro).</param>
        /// <param name="fin">Solicitud final del rango (0 = sin filtro).</param>
        /// <param name="unificarParejas">True para unificar parejas INF + CUM.</param>
        /// <param name="simular">True para calcular sin modificar archivos.</param>
        /// <param name="progress">Callback opcional de progreso (actual, total).</param>
        /// <returns>Resultado con contadores y detalle por archivo; nunca null.</returns>
        ResultadoBusquedaInformes RenombrarYUnificar(string carpetaRaiz, int inicio, int fin,
                                                    bool unificarParejas, bool simular,
                                                    Action<int, int> progress = null);

        /// <summary>
        /// Audita los archivos presentes y escribe la columna INFORME (H) de CONSULTAS con
        /// VERIFICADO, ERROR, FALTA CUM, FALTA INF o SIN INFORME. Mueve los UNI validados a VERIFICADOS.
        /// </summary>
        /// <param name="carpetaRaiz">Carpeta \Informes del libro.</param>
        /// <param name="inicio">Solicitud inicial del rango (0 = sin filtro).</param>
        /// <param name="fin">Solicitud final del rango (0 = sin filtro).</param>
        /// <param name="simular">True para calcular sin escribir en Excel ni mover archivos.</param>
        /// <param name="progress">Callback opcional de progreso (actual, total).</param>
        /// <returns>Resultado con contadores y detalle por fila; nunca null.</returns>
        ResultadoBusquedaInformes AuditarConsultas(string carpetaRaiz, int inicio, int fin,
                                                   bool simular, Action<int, int> progress = null);
    }
}
