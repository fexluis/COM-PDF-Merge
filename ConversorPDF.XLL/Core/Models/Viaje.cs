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
    /// Modelo que representa la información de un viaje o comisión de servicios.
    /// </summary>
    public class Viaje
    {
        /// <summary> Fecha de inicio del viaje. </summary>
        public DateTime FechaInicio { get; set; }

        /// <summary> Fecha de regreso del viaje. </summary>
        public DateTime FechaFin { get; set; }

        /// <summary> Tiempo total en días de la comisión. </summary>
        public int Duracion { get; set; }

        /// <summary> Municipio de origen de la comisión. </summary>
        public string CiudadOrigen { get; set; }

        /// <summary> Municipio de destino de la comisión. </summary>
        public string CiudadDestino { get; set; }

        /// <summary> Detalle día a día del recorrido. </summary>
        public string Itinerario { get; set; }

        /// <summary> Valor asignado para transporte terrestre o fluvial. </summary>
        public decimal Transporte { get; set; }

        /// <summary> Valor de tiquetes aéreos (si aplica). </summary>
        public decimal TransporteAereo { get; set; }

        /// <summary> Objeto o justificación de la comisión. </summary>
        public string Objeto { get; set; }

        /// <summary> Terminal de transporte asociada. </summary>
        public string terminal { get; set; }

        /// <summary> Asignatura académica vinculada (para docentes). </summary>
        public string Asignatura { get; set; }

        /// <summary> Unidad administrativa solicitante. </summary>
        public string Dependencia { get; set; }

        /// <summary> Documento de identidad del comisionado. </summary>
        public string cedula { get; set; }

        /// <summary> Indicador de transporte aéreo (0 = Terrestre, 1 = Aéreo). </summary>
        public int IsTerminalAereo { get; set; }

        /// <summary> ID de la terminal aérea seleccionada en tabla terminales (0 si no aplica). </summary>
        public int IdTerminales { get; set; }

        /// <summary> Nombre descriptivo de la terminal o aeropuerto de destino. </summary>
        public string TerminalAereaNombre { get; set; }

        /// <summary> Valor asignado por concepto de gasto de terminal aérea. </summary>
        public decimal GastoTerminalAerea { get; set; }
    }
}