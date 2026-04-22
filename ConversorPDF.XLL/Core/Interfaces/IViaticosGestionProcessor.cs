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
        /// Renombra archivos PDF de solicitudes en una carpeta basándose en los datos de la hoja de CONSULTAS y validación por OCR de la cédula.
        /// </summary>
        /// <param name="folder">Ruta de la carpeta que contiene los archivos PDF a renombrar.</param>
        void RenombrarSolicitudes(string folder);
        /// <summary>
        /// Une los diferentes componentes de un expediente de viáticos (Solicitud, Resolución, RP, CDP, Verificado) en un único archivo PDF por cada hoja del libro.
        /// </summary>
        /// <param name="folder">Ruta base donde se encuentran las subcarpetas de documentos (Solicitudes, RPs, CDPs, etc.).</param>
        /// <returns>Cantidad de expedientes unificados con éxito.</returns>
        int UnificarSolicitudes(string folder);
    }
}