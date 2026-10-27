using System;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    public class TrayectoEncoderService : ITrayectoEncoderService
    {
        private readonly IGastosTerrestresRepository _repository;
        private readonly IConfigRepository _configRepository;

        public TrayectoEncoderService(IGastosTerrestresRepository repository, IConfigRepository configRepository)
        {
            _repository = repository;
            _configRepository = configRepository;
        }

        public string GetOrCreateCodigoTrayecto(string territorial, string trayecto, decimal? valorTransporte)
        {
            if (string.IsNullOrWhiteSpace(territorial) || string.IsNullOrWhiteSpace(trayecto))
                return "";

            // 1. Verificar si ya existe este trayecto para esta territorial
            string codigoExistente = _repository.GetCodigoTrayectoExistente(territorial, trayecto);
            if (!string.IsNullOrEmpty(codigoExistente))
            {
                return codigoExistente;
            }

            // 2. Generar nuevo prefijo: primeras 3 letras de la territorial
            string prefijo = territorial.Length >= 3 
                ? territorial.Substring(0, 3).ToUpper() 
                : territorial.PadRight(3, 'X').ToUpper();

            string prefijoBusqueda = $"{prefijo}-GT";

            // 3. Obtener el último código generado para este prefijo
            string ultimoCodigo = _repository.GetUltimoCodigoTrayecto(prefijoBusqueda);

            int consecutivo = 1;
            if (!string.IsNullOrEmpty(ultimoCodigo) && ultimoCodigo.StartsWith(prefijoBusqueda))
            {
                string numPart = ultimoCodigo.Substring(prefijoBusqueda.Length);
                if (int.TryParse(numPart, out int ultimoConsecutivo))
                {
                    consecutivo = ultimoConsecutivo + 1;
                }
            }

            string nuevoCodigo = $"{prefijoBusqueda}{consecutivo.ToString("D3")}";

            // 4. Guardar en base de datos (con manejo de carrera)
            var configResult = _configRepository.GetConfig();
            string currentSubunidad = configResult.IsSuccess ? configResult.Value.Subunidad : territorial;
            try
            {
                _repository.GuardarTrayecto(territorial, nuevoCodigo, trayecto, valorTransporte, currentSubunidad);
                return nuevoCodigo;
            }
            catch (Exception ex) when (EsErrorDeConstraint(ex))
            {
                // Si hubo carrera, re-sincronizar con la base de datos antes de reintentar
                string ultimoCodigoActualizado = _repository.GetUltimoCodigoTrayecto(prefijoBusqueda);
                if (!string.IsNullOrEmpty(ultimoCodigoActualizado) && ultimoCodigoActualizado.StartsWith(prefijoBusqueda))
                {
                    string numPart = ultimoCodigoActualizado.Substring(prefijoBusqueda.Length);
                    if (int.TryParse(numPart, out int ultimoConsecutivo))
                    {
                        consecutivo = ultimoConsecutivo + 1;
                    }
                    else
                    {
                        consecutivo++;
                    }
                }
                else
                {
                    consecutivo++;
                }

                nuevoCodigo = $"{prefijoBusqueda}{consecutivo.ToString("D3")}";

                try
                {
                    _repository.GuardarTrayecto(territorial, nuevoCodigo, trayecto, valorTransporte, currentSubunidad);
                    return nuevoCodigo;
                }
                catch (Exception innerEx)
                {
                    // Loggear y retornar vacío para no romper la exportación
                    LoggerHelper.Error($"[TrayectoEncoderService] Fallo irreparable al guardar trayecto {nuevoCodigo}", innerEx);
                    return "";
                }
            }
        }

        private bool EsErrorDeConstraint(Exception ex)
        {
            if (ex == null || string.IsNullOrEmpty(ex.Message)) return false;
            string msg = ex.Message.ToLowerInvariant();
            return msg.Contains("unique") || 
                   msg.Contains("constraint") || 
                   msg.Contains("duplicate") || 
                   msg.Contains("restricción");
        }
    }
}
