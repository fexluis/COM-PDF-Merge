using System.Collections.Generic;
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
    /// Interfaz que define las operaciones de manipulación de archivos PDF.
    /// </summary>
    public interface IPdfService
    {
        /// <summary>
        /// Combina varios archivos PDF en un archivo de salida único.
        /// </summary>
        /// <param name="archivos">Lista de archivos PDF de entrada.</param>
        /// <param name="rutaSalida">Ruta absoluta del archivo resultante.</param>
        void CombinarPDFs(IEnumerable<string> archivos, string rutaSalida);
    }
}