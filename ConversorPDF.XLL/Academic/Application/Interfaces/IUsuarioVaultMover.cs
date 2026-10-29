using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface IUsuarioVaultMover
    {
        Result<MovimientoVaultResultado> Mover(Usuario usuario, string subunidadDestino);
    }
}
