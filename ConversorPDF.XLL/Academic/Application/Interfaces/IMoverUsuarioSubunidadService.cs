using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface IMoverUsuarioSubunidadService
    {
        Result<MovimientoUsuarioResultado> Mover(Usuario usuario, string subunidadDestino, string usuarioActual);
    }
}
