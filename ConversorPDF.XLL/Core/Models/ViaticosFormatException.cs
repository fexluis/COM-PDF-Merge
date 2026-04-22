using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Excepción lanzada cuando el libro de Excel no cumple con el formato esperado 
    /// (hojas o tablas faltantes).
    /// </summary>
    public class ViaticosFormatException : Exception
    {
        /// <summary> Inicializa una nueva instancia con un mensaje de error específico. </summary>
        public ViaticosFormatException(string message) : base(message)
        {
        }

        /// <summary> Inicializa una nueva instancia con un mensaje de error y una referencia a la causa raíz. </summary>
        public ViaticosFormatException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}