using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface IMoverUsuarioSubunidadService
    {
        Result Mover(string usuarioId, string subunidadDestino);
    }
}
