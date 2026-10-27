namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    public interface IGastosTerrestresRepository
    {
        string GetCodigoTrayectoExistente(string territorial, string trayecto);
        string GetUltimoCodigoTrayecto(string prefijo);
        void GuardarTrayecto(string territorial, string codigoTrayecto, string trayecto, decimal? valor, string subunidad);
        
        System.Threading.Tasks.Task<System.Collections.Generic.IEnumerable<Models.GastoTerrestreEntity>> GetAllAsync();
        System.Threading.Tasks.Task UpdateAsync(Models.GastoTerrestreEntity entity);
        System.Threading.Tasks.Task DeleteAsync(int id);
    }
}
