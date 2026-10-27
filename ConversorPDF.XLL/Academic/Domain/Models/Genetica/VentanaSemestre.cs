using System;
using ConversorPDF.XLL.Academic.Domain.Enums;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa la ventana temporal canónica del semestre [FI, FF] dentro de la cual
    /// deben ocurrir todas las sesiones académicas generadas u optimizadas por el AG.
    /// </summary>
    public class VentanaSemestre
    {
        public int IdPeriodo { get; set; }
        public string PeriodoNombre { get; set; }
        public DateTime FechaInicio { get; set; }
        public DateTime FechaFin { get; set; }
        public AgOrigenVentana Origen { get; set; }
        public string Observaciones { get; set; }

        public int DuracionDias => (int)(FechaFin.Date - FechaInicio.Date).TotalDays + 1;

        public VentanaSemestre()
        {
            PeriodoNombre = string.Empty;
            Observaciones = string.Empty;
            Origen = AgOrigenVentana.Default;
        }

        /// <summary>
        /// Método de fábrica para construir y validar una ventana semestral con el patrón Result.
        /// </summary>
        public static Result<VentanaSemestre> Crear(
            int idPeriodo,
            string periodoNombre,
            DateTime fechaInicio,
            DateTime fechaFin,
            AgOrigenVentana origen,
            string observaciones = "")
        {
            if (fechaFin.Date < fechaInicio.Date)
            {
                return Result.Failure<VentanaSemestre>(
                    $"Ventana inválida: {fechaFin:yyyy-MM-dd} < {fechaInicio:yyyy-MM-dd}.");
            }

            return Result.Success(new VentanaSemestre
            {
                IdPeriodo = idPeriodo,
                PeriodoNombre = periodoNombre ?? string.Empty,
                FechaInicio = fechaInicio.Date,
                FechaFin = fechaFin.Date,
                Origen = origen,
                Observaciones = observaciones ?? string.Empty
            });
        }

        /// <summary>
        /// Comprueba si una fecha dada se encuentra estrictamente dentro del rango de la ventana.
        /// </summary>
        public bool Contiene(DateTime fecha)
        {
            var f = fecha.Date;
            return f >= FechaInicio.Date && f <= FechaFin.Date;
        }

        /// <summary>
        /// Comprueba si un rango [inicio, fin] está completamente contenido dentro de la ventana.
        /// </summary>
        public bool ContieneRango(DateTime inicio, DateTime fin)
        {
            return inicio.Date >= FechaInicio.Date && fin.Date <= FechaFin.Date && inicio.Date <= fin.Date;
        }

        public override string ToString()
        {
            return $"{PeriodoNombre}: {FechaInicio:yyyy-MM-dd} al {FechaFin:yyyy-MM-dd} ({Origen})";
        }
    }
}
