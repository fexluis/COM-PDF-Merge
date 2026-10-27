using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Servicio de sincronización y derivación automática del pool de docentes-asignaturas desde el histórico de vinculaciones.
    /// </summary>
    public interface ITerceroAsignaturaSyncService
    {
        /// <summary>
        /// Sincroniza y recomputa globalmente todas las habilitaciones históricas presentes en 'vinculaciones'.
        /// </summary>
        Task<Result<int>> SincronizarTodoAsync(string subunidad, CancellationToken ct);

        /// <summary>
        /// Sincroniza y recomputa las habilitaciones para un programa y semestre específicos.
        /// </summary>
        Task<Result<int>> SincronizarPorSemestreAsync(int idPrograma, string semestre, string subunidad, CancellationToken ct);

        /// <summary>
        /// Recomputa atómicamente el conteo y última fecha de asignación para una tupla específica tras una inserción, edición o borrado.
        /// </summary>
        Task<Result> SincronizarDocenteAsignaturaAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, string subunidad, CancellationToken ct);
    }
}
