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
    /// Representa una fila de datos de viáticos extraída de las hojas de programación o consultas de Excel.
    /// </summary>
    public class ViaticoRow
    {
        /// <summary> Índice de la fila en la hoja de Excel. </summary>
        public int Index { get; set; }

        /// <summary> Identificador numérico de la comisión administrativa. </summary>
        public string NumeroComision { get; set; }

        /// <summary> Fecha de emisión del documento de comisión. </summary>
        public DateTime FechaComision { get; set; }

        /// <summary> Fecha en la que inicia la comisión. </summary>
        public DateTime FechaInicio { get; set; }

        /// <summary> Fecha de finalización programada de la comisión. </summary>
        public DateTime FechaFin { get; set; }

        /// <summary> Estado actual del trámite (ej: Pagado, Pendiente). </summary>
        public string Estado { get; set; }

        /// <summary> Número o código del informe de actividades presentado. </summary>
        public string Informe { get; set; }

        /// <summary> Documento de identidad del beneficiario. </summary>
        public string Cedula { get; set; }

        /// <summary> Nombre completo del comisionado. </summary>
        public string NombreCompleto { get; set; }

        /// <summary> Cargo o rol del funcionario. </summary>
        public string Rol { get; set; }

        /// <summary> Descripción del itinerario de viaje. </summary>
        public string Itinerario { get; set; }

        /// <summary> Número de Certificado de Disponibilidad Presupuestal (CDP). </summary>
        public string CDP { get; set; }

        /// <summary> Número de Registro Presupuestal (RP). </summary>
        public string RP { get; set; }

        /// <summary> Valor total liquidado por concepto de viáticos. </summary>
        public string ValorTotal { get; set; }

        /// <summary> Texto original del objeto de la comisión (sin procesar). </summary>
        public string ObjetoRAW { get; set; }

        /// <summary> Texto del objeto de la comisión tras limpieza y formateo. </summary>
        public string ObjetoProcesado { get; set; }

        /// <summary> Número de Obligación (OB). </summary>
        public string OB { get; set; }

        /// <summary> Número de Orden de Pago (OP). </summary>
        public string OP { get; set; }
    }
}