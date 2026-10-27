using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    public class GastosTerrestresService : IGastosTerrestresService
    {
        private readonly IGastosTerrestresRepository _repository;
        private readonly ITrayectoEncoderService _encoderService;

        public GastosTerrestresService(IGastosTerrestresRepository repository, ITrayectoEncoderService encoderService)
        {
            _repository = repository;
            _encoderService = encoderService;
        }

        public async Task<Result<IEnumerable<GastoTerrestreGridDto>>> GetAllAsync()
        {
            try
            {
                var entities = await _repository.GetAllAsync();
                var dtos = GastoTerrestreMapper.MapToDtoList(entities);
                return Result.Success(dtos);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[GastosTerrestresService] Error en GetAllAsync", ex);
                return Result.Failure<IEnumerable<GastoTerrestreGridDto>>($"Error al cargar datos: {ex.Message}");
            }
        }

        public async Task<Result> UpdateAsync(GastoTerrestreGridDto dto)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dto.Trayecto) || string.IsNullOrWhiteSpace(dto.Territorial))
                    return Result.Failure("El trayecto y la territorial son obligatorios.");

                var entity = GastoTerrestreMapper.MapToEntity(dto);
                await _repository.UpdateAsync(entity);
                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[GastosTerrestresService] Error en UpdateAsync", ex);
                return Result.Failure($"Error al actualizar: {ex.Message}");
            }
        }

        public async Task<Result> DeleteAsync(int id)
        {
            try
            {
                await _repository.DeleteAsync(id);
                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[GastosTerrestresService] Error en DeleteAsync", ex);
                return Result.Failure($"Error al eliminar: {ex.Message}");
            }
        }

        public async Task<Result<string>> CrearTrayectoAsync(string territorial, string trayecto, decimal? valor)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(trayecto) || string.IsNullOrWhiteSpace(territorial))
                    return Result.Failure<string>("El trayecto y la territorial son obligatorios.");

                return await Task.Run(() =>
                {
                    string codigo = _encoderService.GetOrCreateCodigoTrayecto(territorial, trayecto, valor);
                    if (string.IsNullOrEmpty(codigo))
                        return Result.Failure<string>("No se pudo generar el código del trayecto.");

                    return Result.Success(codigo);
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[GastosTerrestresService] Error en CrearTrayectoAsync", ex);
                return Result.Failure<string>($"Error al crear trayecto: {ex.Message}");
            }
        }
    }
}
