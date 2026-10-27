using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Puerto de persistencia para la gestión de perfiles de selección de árbol académico (ag_arbol y ag_arbol_nodo) (Fase 237).
    /// </summary>
    public interface IAgArbolRepository
    {
        Task<Result<long>> GuardarAsync(AgArbol perfil, IReadOnlyCollection<AgArbolNodo> nodos, CancellationToken ct);
        Task<Result<AgArbol>> ObtenerAsync(long idArbol, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgArbolNodo>>> ObtenerNodosAsync(long idArbol, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgArbol>>> ListarAsync(string subunidad, CancellationToken ct);
        Task<Result<bool>> ExisteNombreAsync(string nombre, string subunidad, long excluirId, CancellationToken ct);
        Task<Result> RenombrarAsync(long idArbol, string nuevoNombre, CancellationToken ct);
        Task<Result> EliminarAsync(long idArbol, CancellationToken ct);
    }
}
