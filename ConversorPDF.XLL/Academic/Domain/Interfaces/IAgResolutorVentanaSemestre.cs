using System;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de resolución de la ventana semestral (α) según la jerarquía de fuentes:
    /// Festivos(PRESENCIAL) -> Config(id_periodo) -> Manual(clamp) -> Bloqueante.
    /// </summary>
    public interface IAgResolutorVentanaSemestre
    {
        /// <summary>
        /// Resuelve de forma asíncrona la ventana temporal para el periodo indicado, aplicando cascada de fuentes.
        /// </summary>
        Task<Result<VentanaSemestre>> ResolverVentanaAsync(int idPeriodo, DateTime? fechaInicioManual, DateTime? fechaFinManual, CancellationToken ct);

        /// <summary>
        /// Resuelve de forma síncrona la ventana temporal para el periodo indicado.
        /// </summary>
        Result<VentanaSemestre> ResolverVentana(int idPeriodo, DateTime? fechaInicioManual, DateTime? fechaFinManual);
    }
}
