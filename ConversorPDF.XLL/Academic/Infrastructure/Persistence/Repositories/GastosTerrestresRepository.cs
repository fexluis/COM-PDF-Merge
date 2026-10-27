using System.Data;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class GastosTerrestresRepository : IGastosTerrestresRepository
    {
        private readonly IDbConnection _connection;

        public GastosTerrestresRepository(IDbConnection connection)
        {
            _connection = connection;
        }

        public string GetCodigoTrayectoExistente(string territorial, string trayecto)
        {
            return _connection.QueryFirstOrDefault<string>(
                "SELECT CodigoTrayecto FROM BaseGastosTerrestre WHERE Territorial = @Territorial AND Trayecto = @Trayecto LIMIT 1",
                new { Territorial = territorial, Trayecto = trayecto });
        }

        public string GetUltimoCodigoTrayecto(string prefijo)
        {
            return _connection.Query<string>(
                "SELECT CodigoTrayecto FROM BaseGastosTerrestre WHERE CodigoTrayecto LIKE @PrefijoLike ORDER BY Id DESC LIMIT 1",
                new { PrefijoLike = prefijo + "%" }).FirstOrDefault();
        }

        public void GuardarTrayecto(string territorial, string codigoTrayecto, string trayecto, decimal? valor, string subunidad)
        {
            _connection.Execute(
                "INSERT INTO BaseGastosTerrestre (Territorial, CodigoTrayecto, Trayecto, Valor, subunidad) VALUES (@Territorial, @CodigoTrayecto, @Trayecto, @Valor, @Subunidad)",
                new { Territorial = territorial, CodigoTrayecto = codigoTrayecto, Trayecto = trayecto, Valor = valor, Subunidad = subunidad });
        }

        public async System.Threading.Tasks.Task<System.Collections.Generic.IEnumerable<Domain.Models.GastoTerrestreEntity>> GetAllAsync()
        {
            return await _connection.QueryAsync<Domain.Models.GastoTerrestreEntity>("SELECT Id, Territorial, CodigoTrayecto, Trayecto, Valor, subunidad FROM BaseGastosTerrestre ORDER BY CodigoTrayecto");
        }

        public async System.Threading.Tasks.Task UpdateAsync(Domain.Models.GastoTerrestreEntity entity)
        {
            // Fase 25: Se eliminó la actualización de 'Territorial' para mantener la inmutabilidad
            // que asegura la consistencia del prefijo del CodigoTrayecto.
            await _connection.ExecuteAsync(
                "UPDATE BaseGastosTerrestre SET Trayecto = @Trayecto, Valor = @Valor WHERE Id = @Id",
                new { entity.Trayecto, entity.Valor, entity.Id });
        }

        public async System.Threading.Tasks.Task DeleteAsync(int id)
        {
            await _connection.ExecuteAsync("DELETE FROM BaseGastosTerrestre WHERE Id = @Id", new { Id = id });
        }
    }
}
