using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Models;
namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Procesador para la generación de viáticos de forma individual o específica.
    /// </summary>
    public interface IViaticosIndividualProcessor
    {
        /// <summary> Ejecuta el proceso de generación de PDF para los viáticos configurados. </summary>
        int Procesar(Action<int, int> progressCallback = null, DateTime? overrideInicio = null, DateTime? overrideFin = null, string overrideCdp = null);
        
        /// <summary> Ejecuta el proceso de generación únicamente para la fila seleccionada por el usuario. </summary>
        int ProcesarFilaSeleccionada(Action<int, int> progressCallback = null);
    }
}