using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    public class MoverUsuarioSubunidadService : IMoverUsuarioSubunidadService
    {
        private readonly IUsuarioVaultMover _vaultMover;
        private readonly IConfigRepository _configRepository;

        public MoverUsuarioSubunidadService(IUsuarioVaultMover vaultMover, IConfigRepository configRepository)
        {
            _vaultMover = vaultMover;
            _configRepository = configRepository;
        }

        public Result<MovimientoUsuarioResultado> Mover(Usuario usuario, string subunidadDestino, string usuarioActual)
        {
            if (usuario == null)
            {
                return Result.Failure<MovimientoUsuarioResultado>("El usuario es requerido.");
            }

            if (string.IsNullOrWhiteSpace(usuario.NombreUsuario))
            {
                return Result.Failure<MovimientoUsuarioResultado>("El nombre de usuario es requerido.");
            }

            if (string.IsNullOrWhiteSpace(subunidadDestino))
            {
                return Result.Failure<MovimientoUsuarioResultado>("La subunidad destino es requerida.");
            }

            string destinoNormalizado = subunidadDestino.Trim();

            var moverVaultResult = _vaultMover.Mover(usuario, destinoNormalizado);
            if (moverVaultResult.IsFailure)
            {
                return Result.Failure<MovimientoUsuarioResultado>(moverVaultResult.Error);
            }

            string resolvedUsuarioId = !string.IsNullOrWhiteSpace(usuario.Id)
                ? usuario.Id
                : "usr_" + usuario.NombreUsuario.Trim().ToLowerInvariant();

            bool esUsuarioActual = string.Equals(
                usuario.NombreUsuario.Trim(),
                usuarioActual?.Trim(),
                StringComparison.OrdinalIgnoreCase);

            if (esUsuarioActual)
            {
                var updateConfigResult = _configRepository.UpdateSubunidad(destinoNormalizado);
                if (updateConfigResult.IsFailure)
                {
                    return Result.Failure<MovimientoUsuarioResultado>(
                        $"El usuario fue movido pero falló la actualización de la configuración local: {updateConfigResult.Error}");
                }

                return Result.Success(new MovimientoUsuarioResultado
                {
                    UsuarioId = resolvedUsuarioId,
                    SubunidadDestino = destinoNormalizado,
                    ConfigActualizada = true,
                    RequiereReinicio = true
                });
            }

            return Result.Success(new MovimientoUsuarioResultado
            {
                UsuarioId = resolvedUsuarioId,
                SubunidadDestino = destinoNormalizado,
                ConfigActualizada = false,
                RequiereReinicio = false
            });
        }
    }
}
