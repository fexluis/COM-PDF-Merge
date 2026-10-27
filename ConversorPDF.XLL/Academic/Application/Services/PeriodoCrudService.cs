using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación para la gestión integral de períodos académicos con validación de reglas de negocio.
    /// </summary>
    public class PeriodoCrudService : IPeriodoCrudService
    {
        private readonly IAcademicRepository _repository;

        public PeriodoCrudService(IAcademicRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public Result<IEnumerable<PeriodoGridDto>> ObtenerTodos()
        {
            try
            {
                var result = _repository.GetPeriodos();
                if (!result.IsSuccess)
                {
                    return Result.Failure<IEnumerable<PeriodoGridDto>>(result.Error);
                }

                var dtos = PeriodoGridMapper.MapToDtoList(result.Value);
                return Result.Success(dtos);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[PeriodoCrudService] Error en ObtenerTodos", ex);
                return Result.Failure<IEnumerable<PeriodoGridDto>>($"Error al cargar períodos: {ex.Message}");
            }
        }

        public Result<PeriodoGridDto> ObtenerPorId(int idPeriodo)
        {
            try
            {
                var result = _repository.GetPeriodoById(idPeriodo);
                if (!result.IsSuccess || result.Value == null)
                {
                    return Result.Failure<PeriodoGridDto>(result.Error ?? "Período no encontrado.");
                }

                return Result.Success(PeriodoGridMapper.MapToDto(result.Value));
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[PeriodoCrudService] Error en ObtenerPorId {idPeriodo}", ex);
                return Result.Failure<PeriodoGridDto>($"Error al obtener período: {ex.Message}");
            }
        }

        public Result<int> Guardar(PeriodoGridDto dto)
        {
            if (dto == null) return Result.Failure<int>("Los datos del período no pueden ser nulos.");

            // RN1: Nombre de período obligatorio y formato no vacío
            if (string.IsNullOrWhiteSpace(dto.PeriodoNombre))
            {
                return Result.Failure<int>("El código o nombre del período es obligatorio.");
            }

            dto.PeriodoNombre = dto.PeriodoNombre.Trim();

            // RN2: Ventana temporal (periodo_fin >= periodo_inicio) si ambas fechas están presentes
            if (!string.IsNullOrWhiteSpace(dto.PeriodoInicio) && !string.IsNullOrWhiteSpace(dto.PeriodoFin))
            {
                var formats = new[] { "yyyy-MM-dd", "dd/MM/yyyy", "yyyy/MM/dd" };
                if (DateTime.TryParseExact(dto.PeriodoInicio, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dIni) &&
                    DateTime.TryParseExact(dto.PeriodoFin, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dFin))
                {
                    if (dFin < dIni)
                    {
                        return Result.Failure<int>("La fecha de fin del período no puede ser anterior a la fecha de inicio.");
                    }
                    dto.PeriodoInicio = dIni.ToString("yyyy-MM-dd");
                    dto.PeriodoFin = dFin.ToString("yyyy-MM-dd");
                }
            }

            try
            {
                var entity = PeriodoGridMapper.MapToEntity(dto);

                if (entity.IdPeriodo > 0)
                {
                    var exist = _repository.GetPeriodoById(entity.IdPeriodo);
                    if (exist.IsSuccess && exist.Value != null)
                    {
                        var updResult = _repository.UpdatePeriodo(entity);
                        if (!updResult.IsSuccess)
                        {
                            return Result.Failure<int>(updResult.Error);
                        }
                        return Result.Success(entity.IdPeriodo);
                    }
                }

                var insResult = _repository.InsertPeriodo(entity);
                if (!insResult.IsSuccess)
                {
                    return Result.Failure<int>(insResult.Error);
                }

                return Result.Success(insResult.Value);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[PeriodoCrudService] Error en Guardar", ex);
                return Result.Failure<int>($"Error al guardar período: {ex.Message}");
            }
        }

        public Result Eliminar(int idPeriodo)
        {
            if (idPeriodo <= 0) return Result.Failure("Identificador de período inválido.");

            try
            {
                var delResult = _repository.DeletePeriodo(idPeriodo);
                if (!delResult.IsSuccess)
                {
                    return Result.Failure(delResult.Error);
                }

                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[PeriodoCrudService] Error en Eliminar {idPeriodo}", ex);
                return Result.Failure($"Error al eliminar período: {ex.Message}");
            }
        }
    }
}
