using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Define las operaciones de gestión documental para viáticos, incluyendo renombramiento de archivos mediante OCR y unificación de expedientes.
    /// </summary>
    public interface IViaticosGestionProcessor
    {
        /// <summary>
        /// Renombra archivos PDF de solicitudes en una carpeta basándose en los datos de la hoja de CONSULTAS y validación por OCR de la cédula, aplicando opcionalmente filtro por rango de comisiones.
        /// </summary>
        /// <param name="folder">Ruta de la carpeta que contiene los archivos PDF a renombrar.</param>
        /// <param name="inicio">Número de viático inicial del rango (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final del rango (0 para sin filtro).</param>
        /// <returns>Cantidad de archivos renombrados con éxito, o -1 en caso de error.</returns>
        int RenombrarSolicitudes(string folder, int inicio, int fin);

        /// <summary>
        /// Renombra las solicitudes devolviendo el desglose completo (carpetas, motivos y contadores).
        /// </summary>
        /// <param name="folder">Carpeta raíz del libro (se procesa también la subcarpeta Solicitudes).</param>
        /// <param name="inicio">Número de viático inicial del rango (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final del rango (0 para sin filtro).</param>
        /// <returns>Resultado con contadores y motivo por archivo; nunca null.</returns>
        ResultadoRenombrado RenombrarSolicitudesConDetalle(string folder, int inicio, int fin);

        /// <summary>
        /// Obtiene el rango (primer y último registro válido) de la columna B de la hoja CONSULTAS.
        /// </summary>
        /// <returns>Rango detectado; HayDatos = false si no hay registros legibles.</returns>
        RangoSolicitudes ObtenerRangoSolicitudes();

        /// <summary>
        /// Verifica la presencia física de cada solicitud del rango y marca la columna INFORME (H)
        /// con "SIN SOLICITUD" cuando no existe el PDF correspondiente (comisión + cédula).
        /// Protege VERIFICADO/REINTEGRO, copia estados terminales y limpia marcas erróneas.
        /// </summary>
        /// <param name="folder">Carpeta raíz del libro (se revisan también Solicitudes).</param>
        /// <param name="inicio">Número de viático inicial del rango.</param>
        /// <param name="fin">Número de viático final del rango.</param>
        /// <returns>Resultado con contadores, celdas escritas/limpiadas y estado del guardado.</returns>
        ResultadoVerificacionSolicitudes VerificarSolicitudesPresentes(string folder, int inicio, int fin);

        /// <summary>
        /// Une los diferentes componentes de un expediente de viáticos (Solicitud, Resolución, RP, CDP, Verificado) en un único archivo PDF por cada hoja del libro, aplicando opcionalmente filtro por rango de comisiones.
        /// </summary>
        /// <param name="folder">Ruta base donde se encuentran las subcarpetas de documentos (Solicitudes, RPs, CDPs, etc.).</param>
        /// <param name="inicio">Número de viático inicial del rango (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final del rango (0 para sin filtro).</param>
        /// <returns>Cantidad de expedientes unificados con éxito, o -1 en caso de error.</returns>
        int UnificarSolicitudes(string folder, int inicio, int fin);
    }
}
