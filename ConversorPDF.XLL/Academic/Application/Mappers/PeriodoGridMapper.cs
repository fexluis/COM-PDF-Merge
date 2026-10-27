using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador bidireccional entre la entidad de dominio Periodo y el DTO de grilla PeriodoGridDto.
    /// </summary>
    public static class PeriodoGridMapper
    {
        public static PeriodoGridDto MapToDto(Periodo entity)
        {
            if (entity == null) return null;

            return new PeriodoGridDto
            {
                IdPeriodo = entity.IdPeriodo,
                PeriodoNombre = entity.PeriodoNombre,
                PeriodoInicio = entity.PeriodoInicio,
                PeriodoFin = entity.PeriodoFin,
                Activo = entity.Activo ? 1 : 0
            };
        }

        public static Periodo MapToEntity(PeriodoGridDto dto)
        {
            if (dto == null) return null;

            return new Periodo
            {
                IdPeriodo = dto.IdPeriodo,
                PeriodoNombre = dto.PeriodoNombre,
                PeriodoInicio = dto.PeriodoInicio,
                PeriodoFin = dto.PeriodoFin,
                Activo = dto.Activo == 1
            };
        }

        public static IEnumerable<PeriodoGridDto> MapToDtoList(IEnumerable<Periodo> entities)
        {
            if (entities == null) return Enumerable.Empty<PeriodoGridDto>();
            return entities.Select(MapToDto).ToList();
        }
    }
}
