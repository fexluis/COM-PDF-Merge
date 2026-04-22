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
    /// Procesador especializado en la generación de formatos de cumplimiento e informes de actividades.
    /// </summary>
    public interface IViaticosCumplidosProcessor
    {
        /// <summary> Procesa los cumplidos para todas las filas del libro de Excel. </summary>
        int Procesar(Action<int, int> progressCallback = null);
        
        /// <summary> Procesa el cumplido únicamente para la fila actualmente seleccionada. </summary>
        int ProcesarFilaSeleccionada(Action<int, int> progressCallback = null);
    }
}