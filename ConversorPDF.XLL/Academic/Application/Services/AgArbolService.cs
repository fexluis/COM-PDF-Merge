using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación que implementa la lógica de negocio y validación de perfiles de selección de árbol académico (Fase 239).
    /// </summary>
    public class AgArbolService : IAgArbolService
    {
        private readonly IAgArbolRepository _repository;

        public AgArbolService(IAgArbolRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public async Task<Result<long>> GuardarPerfilAsync(
            string nombre,
            string descripcion,
            int idPeriodo,
            VentanaSemestre ventana,
            bool manualVentana,
            int poblacion,
            int generaciones,
            double tasaCruce,
            double tasaMutacion,
            int elitismo,
            IReadOnlyCollection<AgContextoSeleccionDto> selecciones,
            string usuario,
            string subunidad,
            CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(nombre))
            {
                return Result.Failure<long>("El nombre del perfil es obligatorio.");
            }

            if (idPeriodo <= 0)
            {
                return Result.Failure<long>("Debe especificar un período académico válido para el perfil.");
            }

            if (selecciones == null || selecciones.Count == 0)
            {
                return Result.Failure<long>("El perfil debe contener al menos una combinación seleccionada del árbol (>= 1 nodo).");
            }

            string sub = subunidad ?? string.Empty;

            // Fase 255 & 257: Verificación de unicidad y conectividad en el repositorio
            var dup = await _repository.ExisteNombreAsync(nombre.Trim(), sub, 0, ct);
            if (dup.IsFailure)
            {
                return Result.Failure<long>(dup.Error);
            }
            // Si dup.Value == true, el nombre ya existe en esta subunidad; el repositorio ejecutará el upsert transaccional (sobrescritura intencional)

            var perfil = new AgArbol
            {
                Id = 0,
                Nombre = nombre.Trim(),
                Descripcion = descripcion?.Trim() ?? string.Empty,
                IdPeriodo = idPeriodo,
                FechaInicio = ventana != null ? ventana.FechaInicio.ToString("yyyy-MM-dd") : string.Empty,
                FechaFin = ventana != null ? ventana.FechaFin.ToString("yyyy-MM-dd") : string.Empty,
                VentanaOrigen = ventana != null ? ventana.Origen.ToString() : "Periodos",
                ManualVentana = manualVentana ? 1 : 0,
                Poblacion = poblacion > 0 ? poblacion : 50,
                Generaciones = generaciones > 0 ? generaciones : 200,
                TasaCruce = tasaCruce > 0 ? tasaCruce : 0.85,
                TasaMutacion = tasaMutacion > 0 ? tasaMutacion : 0.10,
                Elitismo = elitismo >= 0 ? elitismo : 2,
                EstancamientoMax = 20,
                Semilla = 0,
                Usuario = usuario ?? string.Empty,
                Subunidad = sub,
                Activo = 1
            };

            // Fase 249: Deduplicar nodos por AulaRef para proteger el índice único uq_ag_arbol_nodo
            var nodos = selecciones
                .Select(s => AgArbolMapper.Mapear(s, 0, sub))
                .GroupBy(n => n.AulaRef, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList()
                .AsReadOnly();

            return await _repository.GuardarAsync(perfil, nodos, ct);
        }

        public async Task<Result<Tuple<AgArbol, IReadOnlyCollection<AgContextoSeleccionDto>>>> CargarPerfilAsync(long idArbol, CancellationToken ct)
        {
            if (idArbol <= 0)
            {
                return Result.Failure<Tuple<AgArbol, IReadOnlyCollection<AgContextoSeleccionDto>>>("Identificador de perfil inválido.");
            }

            var perfilRes = await _repository.ObtenerAsync(idArbol, ct);
            if (perfilRes.IsFailure)
            {
                return Result.Failure<Tuple<AgArbol, IReadOnlyCollection<AgContextoSeleccionDto>>>(perfilRes.Error);
            }

            var nodosRes = await _repository.ObtenerNodosAsync(idArbol, ct);
            if (nodosRes.IsFailure)
            {
                return Result.Failure<Tuple<AgArbol, IReadOnlyCollection<AgContextoSeleccionDto>>>(nodosRes.Error);
            }

            var dtos = nodosRes.Value
                .Select(AgArbolMapper.Mapear)
                .ToList()
                .AsReadOnly();

            return Result.Success(Tuple.Create(perfilRes.Value, (IReadOnlyCollection<AgContextoSeleccionDto>)dtos));
        }

        public async Task<Result<IReadOnlyCollection<AgArbol>>> ListarPerfilesAsync(string subunidad, CancellationToken ct)
        {
            return await _repository.ListarAsync(subunidad ?? string.Empty, ct);
        }

        public async Task<Result> EliminarPerfilAsync(long idArbol, CancellationToken ct)
        {
            if (idArbol <= 0)
            {
                return Result.Failure("Identificador de perfil inválido.");
            }

            return await _repository.EliminarAsync(idArbol, ct);
        }
    }
}
