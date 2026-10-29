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
            return Result.Failure<MovimientoUsuarioResultado>("Mover usuario no implementado");
        }
    }
}
