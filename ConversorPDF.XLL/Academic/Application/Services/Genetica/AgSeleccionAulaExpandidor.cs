using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio que valida y expande selecciones de contexto académico, deduplicando por AulaRef
    /// y aplicando el tope de seguridad (cap N &lt;= 5000) para prevenir explosión combinatoria.
    /// </summary>
    public class AgSeleccionAulaExpandidor
    {
        public const int CapMaximoRecomendado = 5000; // Fase 536: ampliado para soporte multi-sede/multi-programa

        public Result<IReadOnlyCollection<AgContextoSeleccionDto>> ValidarYExpandir(
            IEnumerable<AgContextoSeleccionDto> combinaciones,
            int capMaximo = CapMaximoRecomendado)
        {
            if (combinaciones == null)
            {
                return Result.Failure<IReadOnlyCollection<AgContextoSeleccionDto>>("La lista de combinaciones no puede ser nula.");
            }

            var unicas = combinaciones
                .GroupBy(c => c.AulaRef)
                .Select(g => g.First())
                .ToList();

            if (unicas.Count == 0)
            {
                return Result.Failure<IReadOnlyCollection<AgContextoSeleccionDto>>("No se seleccionó ninguna combinación válida en el árbol académico.");
            }

            if (unicas.Count > capMaximo)
            {
                return Result.Failure<IReadOnlyCollection<AgContextoSeleccionDto>>(
                    $"La selección actual genera {unicas.Count} aulas/combinaciones, superando el límite de seguridad permitido ({capMaximo}). Desmarque ramas para reducir el volumen.");
            }

            return Result.Success<IReadOnlyCollection<AgContextoSeleccionDto>>(unicas.AsReadOnly());
        }
    }
}
