using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Servicio de aplicación para la orquestación de perfiles de selección de árbol académico (Fase 239).
    /// </summary>
    public interface IAgArbolService
    {
        Task<Result<long>> GuardarPerfilAsync(
            string nombre,
            string descripcion,
            int idPeriodo,
            VentanaSemestre ventana,
            bool manualVentana,
            int poblacion,
            int generaciones,
            double tasaCruce,
            double tasaMutacion,
            int elitismo,
            IReadOnlyCollection<AgContextoSeleccionDto> selecciones,
            string usuario,
            string subunidad,
            CancellationToken ct);

        Task<Result<Tuple<AgArbol, IReadOnlyCollection<AgContextoSeleccionDto>>>> CargarPerfilAsync(long idArbol, CancellationToken ct);

        Task<Result<IReadOnlyCollection<AgArbol>>> ListarPerfilesAsync(string subunidad, CancellationToken ct);

        Task<Result> EliminarPerfilAsync(long idArbol, CancellationToken ct);
    }
}
