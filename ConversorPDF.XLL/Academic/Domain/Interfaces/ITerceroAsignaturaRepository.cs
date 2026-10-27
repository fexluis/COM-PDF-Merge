using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Contrato de repositorio para el acceso y persistencia del pool docente-asignatura en la tabla CRDT 'terceros_asignaturas'.
    /// </summary>
    public interface ITerceroAsignaturaRepository
    {
        /// <summary>
        /// Obtiene todas las habilitaciones docentes registradas en el pool.
        /// </summary>
        Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetAllAsync(CancellationToken ct);

        /// <summary>
        /// Obtiene todas las asignaturas habilitadas para un docente específico.
        /// </summary>
        Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByTerceroAsync(int idTercero, CancellationToken ct);

        /// <summary>
        /// Obtiene todos los docentes habilitados para una asignatura y semestre (o todos los semestres si semestre está vacío).
        /// </summary>
        Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByAsignaturaAsync(int idAsignatura, string semestre, CancellationToken ct);

        /// <summary>
        /// Obtiene el pool de docentes habilitados para una asignatura y programa específicos.
        /// </summary>
        Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetPoolAsync(int idAsignatura, int idPrograma, string semestre, CancellationToken ct);

        /// <summary>
        /// Obtiene las habilitaciones de un programa académico y semestre.
        /// </summary>
        Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByProgramaAsync(int idPrograma, string semestre, CancellationToken ct);

        /// <summary>
        /// Obtiene una habilitación por su clave primaria física de 63 bits.
        /// </summary>
        Task<Result<TerceroAsignatura>> GetByIdAsync(long id, CancellationToken ct);

        /// <summary>
        /// Obtiene una habilitación por su clave lógica natural.
        /// </summary>
        Task<Result<TerceroAsignatura>> GetByKeyAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, CancellationToken ct);

        /// <summary>
        /// Inserta o actualiza una habilitación en la base de datos de forma idempotente con protección ante colisiones.
        /// </summary>
        Task<Result> UpsertAsync(TerceroAsignatura item, CancellationToken ct);

        /// <summary>
        /// Inserta o actualiza una habilitación en la base de datos de forma síncrona dentro de una transacción activa.
        /// </summary>
        Result Upsert(TerceroAsignatura item, System.Data.IDbTransaction transaction = null);

        /// <summary>
        /// Cuenta las vinculaciones históricas de producción para un docente, programa y asignatura (COUNT DISTINCT id).
        /// </summary>
        Task<int> ContarAsignacionesAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, CancellationToken ct);

        /// <summary>
        /// Cuenta síncronamente las vinculaciones históricas de producción para un docente, programa y asignatura.
        /// </summary>
        int ContarAsignaciones(int idTercero, int idPrograma, int idAsignatura, string semestre);

        /// <summary>
        /// Elimina físicamente una habilitación del pool por su identificador único.
        /// </summary>
        Task<Result> DeleteAsync(long id, CancellationToken ct);
    }
}
