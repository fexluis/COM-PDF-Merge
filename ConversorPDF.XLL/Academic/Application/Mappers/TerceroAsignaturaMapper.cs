using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador bidireccional entre la entidad de dominio <see cref="TerceroAsignatura"/> y el DTO <see cref="TerceroAsignaturaDto"/>.
    /// </summary>
    public static class TerceroAsignaturaMapper
    {
        /// <summary>
        /// Convierte una entidad de dominio a su correspondiente DTO.
        /// </summary>
        public static TerceroAsignaturaDto ToDto(TerceroAsignatura entity, string nombreTercero = "", string nombrePrograma = "", string nombreAsignatura = "")
        {
            if (entity == null) return null;

            return new TerceroAsignaturaDto
            {
                Id = entity.Id,
                IdTercero = entity.IdTercero,
                NombreTercero = nombreTercero ?? string.Empty,
                IdPrograma = entity.IdPrograma,
                NombrePrograma = nombrePrograma ?? string.Empty,
                IdAsignatura = entity.IdAsignatura,
                NombreAsignatura = nombreAsignatura ?? string.Empty,
                Semestre = entity.Semestre,
                Activo = entity.Activo,
                Prioridad = entity.Prioridad,
                VecesAsignada = entity.VecesAsignada,
                FechaUltimaAsignacion = entity.FechaUltimaAsignacion,
                Origen = entity.Origen,
                Observaciones = entity.Observaciones,
                Subunidad = entity.Subunidad
            };
        }

        /// <summary>
        /// Convierte un DTO a su entidad de dominio, calculando la clave si el ID es 0.
        /// </summary>
        public static TerceroAsignatura ToEntity(TerceroAsignaturaDto dto)
        {
            if (dto == null) return null;

            long id = dto.Id;
            if (id == 0)
            {
                id = TerceroAsignaturaKey.Calcular(dto.IdTercero, dto.IdPrograma, dto.IdAsignatura, dto.Semestre);
            }

            var entity = new TerceroAsignatura
            {
                Id = id,
                IdTercero = dto.IdTercero,
                IdPrograma = dto.IdPrograma,
                IdAsignatura = dto.IdAsignatura,
                Semestre = dto.Semestre,
                Activo = dto.Activo,
                Prioridad = dto.Prioridad,
                VecesAsignada = dto.VecesAsignada,
                FechaUltimaAsignacion = dto.FechaUltimaAsignacion,
                Origen = string.IsNullOrWhiteSpace(dto.Origen) ? "MANUAL" : dto.Origen,
                Observaciones = dto.Observaciones ?? string.Empty,
                Subunidad = dto.Subunidad ?? string.Empty
            };

            return entity;
        }

        /// <summary>
        /// Convierte una vinculación académica a una entidad del pool docente-asignatura.
        /// </summary>
        public static TerceroAsignatura FromVinculacion(Vinculacion vinc, int vecesAsignada = 0)
        {
            if (vinc == null) return null;

            string cleanSemestre = (vinc.Semestre ?? string.Empty).Trim().ToUpperInvariant();
            long id = TerceroAsignaturaKey.Calcular(vinc.IdTercero, vinc.IdPrograma, vinc.IdAsignatura, cleanSemestre);

            return new TerceroAsignatura
            {
                Id = id,
                IdTercero = vinc.IdTercero,
                IdPrograma = vinc.IdPrograma,
                IdAsignatura = vinc.IdAsignatura,
                Semestre = cleanSemestre,
                Activo = 1,
                Prioridad = 1,
                VecesAsignada = vecesAsignada,
                FechaUltimaAsignacion = vinc.FechaInicio.HasValue ? vinc.FechaInicio.Value.ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd"),
                Origen = "VINCULACION",
                Observaciones = string.Empty,
                Subunidad = vinc.Subunidad ?? string.Empty
            };
        }
    }
}
