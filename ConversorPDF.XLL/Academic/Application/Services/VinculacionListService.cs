using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Implementación del servicio de aplicación para la lista de vinculaciones.
    /// Sigue el patrón Vertical Slice.
    /// </summary>
    public class VinculacionListService : IVinculacionListService
    {
        private readonly IVinculacionQueryRepository _queryRepository;
        private readonly ISedeLookupRepository _sedeRepository;
        private readonly ITerceroLookupRepository _terceroRepository;
        private readonly IPeriodoLookupRepository _periodoRepository;
        private readonly IProgramaLookupRepository _programaRepository;

        public VinculacionListService(
            IVinculacionQueryRepository queryRepository,
            ISedeLookupRepository sedeRepository,
            ITerceroLookupRepository terceroRepository,
            IPeriodoLookupRepository periodoRepository,
            IProgramaLookupRepository programaRepository)
        {
            _queryRepository = queryRepository;
            _sedeRepository = sedeRepository;
            _terceroRepository = terceroRepository;
            _periodoRepository = periodoRepository;
            _programaRepository = programaRepository;
        }



        public async Task<Result<IReadOnlyCollection<SesionListItem>>> GetListaSesionesAsync(VinculacionFilterDto filtros, CancellationToken ct)
        {
            if (filtros.IdCorrida.HasValue)
            {
                return await _queryRepository.GetSesionesAsync(
                    filtros.IdSede,
                    filtros.IdTercero,
                    filtros.IdPeriodo,
                    filtros.IdPrograma,
                    filtros.ModalidadPrograma,
                    filtros.Subunidad,
                    filtros.Semestre,
                    filtros.SoloViaticas,
                    filtros.FechaDesde,
                    filtros.FechaHasta,
                    -1,
                    0,
                    ct,
                    filtros.IdCorrida);
            }

            return await _queryRepository.GetSesionesAsync(
                filtros.IdSede,
                filtros.IdTercero,
                filtros.IdPeriodo,
                filtros.IdPrograma,
                filtros.ModalidadPrograma,
                filtros.Subunidad,
                filtros.Semestre,
                filtros.SoloViaticas,
                filtros.FechaDesde,
                filtros.FechaHasta,
                -1, // Sin límite de seguridad (permite ver todas las sesiones)
                0,
                ct);
        }

        public async Task<Result<(IReadOnlyCollection<LookupItem> Sedes, IReadOnlyCollection<LookupItem> Terceros, IReadOnlyCollection<LookupItem> Periodos, IReadOnlyCollection<LookupItem> Programas)>> GetCatalogosFiltroAsync(CancellationToken ct)
        {
            try
            {
                var sedesResult = await _sedeRepository.GetSedesLookupAsync(ct);
                if (sedesResult.IsFailure) return Result.Failure<(IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>)>(sedesResult.Error);

                var tercerosResult = await _terceroRepository.GetTercerosLookupAsync(ct);
                if (tercerosResult.IsFailure) return Result.Failure<(IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>)>(tercerosResult.Error);

                var periodosResult = await _periodoRepository.GetPeriodosLookupAsync(ct);
                if (periodosResult.IsFailure) return Result.Failure<(IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>)>(periodosResult.Error);

                var programasResult = await _programaRepository.GetProgramasLookupAsync(ct);
                if (programasResult.IsFailure) return Result.Failure<(IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>)>(programasResult.Error);

                return Result.Success((sedesResult.Value, tercerosResult.Value, periodosResult.Value, programasResult.Value));
            }
            catch (Exception ex)
            {
                return Result.Failure<(IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>, IReadOnlyCollection<LookupItem>)>($"Error al cargar catálogos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene las corridas del staging y las prepara para el ComboBox: antepone el comodín
        /// y puebla las etiquetas en la capa Application (Fase 1490), manteniendo
        /// Infrastructure libre de dependencias hacia Application.
        /// </summary>
        public async Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
            int? idPeriodo,
            int? idPrograma,
            string subunidad,
            CancellationToken ct)
        {
            var resultado = await _queryRepository.GetCorridasStagingAsync(idPeriodo, idPrograma, subunidad, ct);
            if (resultado.IsFailure) return resultado;

            return Result.Success<IReadOnlyCollection<CorridaStagingItem>>(
                AgCorridaSelector.PrepararParaCombo(resultado.Value));
        }

        /// <summary>
        /// Traduce el DTO de filtros de la UI a los parámetros primitivos del repositorio
        /// y delega el resumen de diagnóstico (Fases 1490/1493).
        /// </summary>
        public Task<Result<AgStagingResumen>> GetResumenStagingAsync(
            long idCorrida,
            VinculacionFilterDto filtros,
            CancellationToken ct)
        {
            var f = filtros ?? new VinculacionFilterDto();

            return _queryRepository.GetResumenStagingAsync(
                idCorrida,
                f.IdSede,
                f.IdTercero,
                f.IdPeriodo,
                f.IdPrograma,
                f.ModalidadPrograma,
                f.Subunidad,
                f.Semestre,
                f.SoloViaticas,
                f.FechaDesde,
                f.FechaHasta,
                ct);
        }
    }
}
