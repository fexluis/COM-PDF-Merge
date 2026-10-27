using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Repositorio para la gestión de persistencia de las tablas del AG (aulas, bloques, corridas, individuos, genes, staging).
    /// </summary>
    public interface IAgScheduleRepository
    {
        Task<Result<IReadOnlyCollection<AgAula>>> GetAulasAsync(string subunidad, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgAula>>> GetAulasPorPeriodoAsync(int idPeriodo, int? idSede, int? idPrograma, string semestre, string grupo, string subunidad, CancellationToken ct);
        Task<Result<AgAula>> GetAulaByIdAsync(long idAula, CancellationToken ct);
        Task<Result<long>> InsertAulaAsync(AgAula aula, CancellationToken ct);
        Task<Result> InsertAulaAsignaturasAsync(long idAula, IEnumerable<int> idAsignaturas, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgBloqueHorario>>> GetBloquesHorariosAsync(string subunidad, CancellationToken ct);
        Task<Result<long>> InsertCorridaAsync(AgCorrida corrida, CancellationToken ct);
        Task<Result> UpdateCorridaAsync(AgCorrida corrida, CancellationToken ct);
        Task<Result<AgCorrida>> GetCorridaByIdAsync(long idCorrida, CancellationToken ct);
        Task<Result<AgCorrida>> GetUltimaCorridaAsync(int idPeriodo, int idPrograma, string subunidad, CancellationToken ct);
        Task<Result<long>> SaveIndividuoAsync(AgIndividuo individuo, CancellationToken ct);
        Task<Result<AgIndividuo>> GetIndividuoByIdAsync(long idIndividuo, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgIndividuo>>> GetIndividuosByCorridaAsync(long idCorrida, CancellationToken ct);
        Task<Result> SaveBitacoraAsync(AgBitacora bitacora, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgBitacora>>> GetBitacorasByCorridaAsync(long idCorrida, CancellationToken ct);
        Task<Result> GuardarStagingAsync(long idCorrida, long idIndividuo, IEnumerable<AgVinculacion> vinculaciones, IEnumerable<AgSesion> sesiones, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgVinculacion>>> GetVinculacionesStagingAsync(long idCorrida, CancellationToken ct);
        Task<Result<IReadOnlyCollection<AgSesion>>> GetSesionesStagingAsync(long idCorrida, CancellationToken ct);
        Task<Result> LimpiarStagingAsync(long idCorrida, CancellationToken ct);
        Task<Result> LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct);
    }
}
