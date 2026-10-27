using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Expande los genes cromosómicos del AG en una secuencia cronológica de sesiones individuales (AgSesion)
    /// acotadas por la ventana del semestre [FI, FF], excluyendo festivos y calculando viáticos.
    /// </summary>
    public interface IAgExpansorSesiones
    {
        Result<List<AgSesion>> Expandir(AgGen gen, int idVinculacion, AgContexto contexto);
        Result<List<AgSesion>> ExpandirIndividuo(AgIndividuo individuo, Dictionary<long, int> mapaVinculaciones, AgContexto contexto);
    }
}
