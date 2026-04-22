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
    /// Procesador para la generación masiva de múltiples archivos PDF de viáticos.
    /// </summary>
    public interface IViaticosMasivoProcessor
    {
        /// <summary> Ejecuta la generación masiva de documentos para todas las filas válidas en el libro. </summary>
        int Procesar(Action<int, int> progressCallback = null);
    }
}