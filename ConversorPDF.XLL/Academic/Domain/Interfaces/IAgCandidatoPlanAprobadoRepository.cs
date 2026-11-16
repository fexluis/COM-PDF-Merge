using System.Data;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>Almacén de candidato y ratificación dentro de una transacción existente.</summary>
    public interface IAgCandidatoPlanAprobadoRepository
    {
        /// <summary>Rechaza duplicados de intento/aula u operación; sólo devuelve true para una inserción completa.</summary>
        bool Guardar(IDbConnection db, IDbTransaction tx, AgCandidatoPlanAprobado candidato);
    }
}
