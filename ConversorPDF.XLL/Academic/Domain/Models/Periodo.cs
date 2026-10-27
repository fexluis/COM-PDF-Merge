using System;
using System.Text.RegularExpressions;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Periodo
    {
        public int IdPeriodo { get; set; }
        public int Id { get => IdPeriodo; set => IdPeriodo = value; }
        public string PeriodoNombre { get; set; }
        public string Nombre { get => PeriodoNombre; set => PeriodoNombre = value; }
        public string PeriodoInicio { get; set; }
        public string PeriodoFin { get; set; }
        public bool Activo { get; set; }

        public Periodo() { }

        /// <summary>
        /// Intenta obtener la ventana temporal [inicio, fin] del período parseando en formato ISO y variantes estándar.
        /// </summary>
        /// <param name="inicio">Fecha de inicio resuelta.</param>
        /// <param name="fin">Fecha de fin resuelta.</param>
        /// <returns>True si ambas fechas son válidas y fin >= inicio; de lo contrario false.</returns>
        public bool TryObtenerVentana(out DateTime inicio, out DateTime fin)
        {
            inicio = DateTime.MinValue;
            fin = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(PeriodoInicio) || string.IsNullOrWhiteSpace(PeriodoFin))
            {
                return false;
            }

            string[] formats = { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "yyyyMMdd" };

            bool parseIni = DateTime.TryParseExact(PeriodoInicio.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out inicio)
                            || DateTime.TryParse(PeriodoInicio.Trim(), out inicio);

            bool parseFin = DateTime.TryParseExact(PeriodoFin.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out fin)
                            || DateTime.TryParse(PeriodoFin.Trim(), out fin);

            if (parseIni && parseFin && fin.Date >= inicio.Date)
            {
                inicio = inicio.Date;
                fin = fin.Date;
                return true;
            }

            return false;
        }

        public static Result<Periodo> Create(int idPeriodo, string periodoNombre, bool activo = false, string periodoInicio = null, string periodoFin = null)
        {
            if (string.IsNullOrWhiteSpace(periodoNombre))
                return Result.Failure<Periodo>("El período no puede estar vacío.");

            if (!Regex.IsMatch(periodoNombre, @"^\d{4}-\d{2}$"))
                return Result.Failure<Periodo>("El formato del período debe ser 'YYYY-MM' (ej: 2026-01).");

            var periodo = new Periodo
            {
                IdPeriodo = idPeriodo,
                PeriodoNombre = periodoNombre,
                Activo = activo,
                PeriodoInicio = periodoInicio,
                PeriodoFin = periodoFin
            };

            if (!string.IsNullOrWhiteSpace(periodoInicio) && !string.IsNullOrWhiteSpace(periodoFin))
            {
                if (!periodo.TryObtenerVentana(out DateTime ini, out DateTime fin))
                {
                    return Result.Failure<Periodo>("El rango de fechas del período es inválido (la fecha de fin debe ser posterior o igual a la de inicio).");
                }
            }

            return Result.Success(periodo);
        }
    }
}

