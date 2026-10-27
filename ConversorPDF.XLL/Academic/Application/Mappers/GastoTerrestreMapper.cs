using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    public static class GastoTerrestreMapper
    {
        public static GastoTerrestreGridDto MapToDto(GastoTerrestreEntity entity)
        {
            if (entity == null) return null;
            
            return new GastoTerrestreGridDto
            {
                Id = entity.Id,
                Territorial = entity.Territorial,
                CodigoTrayecto = entity.CodigoTrayecto,
                Trayecto = entity.Trayecto,
                Valor = entity.Valor,
                Subunidad = entity.Subunidad
            };
        }

        public static GastoTerrestreEntity MapToEntity(GastoTerrestreGridDto dto)
        {
            if (dto == null) return null;

            return new GastoTerrestreEntity
            {
                Id = dto.Id,
                Territorial = dto.Territorial,
                CodigoTrayecto = dto.CodigoTrayecto,
                Trayecto = dto.Trayecto,
                Valor = dto.Valor,
                Subunidad = dto.Subunidad
            };
        }

        public static IEnumerable<GastoTerrestreGridDto> MapToDtoList(IEnumerable<GastoTerrestreEntity> entities)
        {
            if (entities == null) return Enumerable.Empty<GastoTerrestreGridDto>();
            return entities.Select(MapToDto).ToList();
        }
    }
}
