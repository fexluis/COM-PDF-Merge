using ConversorPDF.XLL.Academic.Application.Interfaces;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de transición de usuario entre subunidades territoriales.
    /// Ruling R-20261027-T6-01: La transición física de vault se realiza en el ciclo de arranque
    /// por SyncContext.ResolveSubunidadWithDialog a partir de la subunidad persistida en usuarios y config.
    /// </summary>
    public class MoverUsuarioSubunidadService : IMoverUsuarioSubunidadService
    {
        public Result Mover(string usuarioId, string subunidadDestino)
        {
            if (string.IsNullOrWhiteSpace(usuarioId) || string.IsNullOrWhiteSpace(subunidadDestino))
            {
                return Result.Failure("Identificador de usuario y subunidad destino son requeridos.");
            }

            return Result.Success();
        }
    }
}
