using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador seguro de selección de nodo del árbol jerárquico a AgContextoSeleccionDto (Fase 1053).
    /// Copia las 3 dimensiones de decisión (Jornada, HorasDia, DiasVentana) y valida completitud (I1, I6).
    /// Internal para prevenir registro de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgSeleccionMapper
    {
        public static Result<AgContextoSeleccionDto> DesdeNodo(
            AgNodoArbolKey key,
            int idPeriodo,
            int idSede,
            int idPrograma,
            string semestre,
            string grupo)
        {
            if (key == null)
            {
                return Result.Failure<AgContextoSeleccionDto>("I1_KEY_NULA: Nodo de árbol sin metadatos tipados.");
            }

            var dto = new AgContextoSeleccionDto(idPeriodo, idSede, idPrograma, semestre, grupo, key.Codigo)
            {
                Jornada = (key.Jornada ?? string.Empty).Trim().ToUpperInvariant(),
                HorasDia = key.HorasDia,
                DiasVentana = (key.DiasVentana ?? string.Empty).Trim().ToUpperInvariant()
            };

            // Validar ventana de días (debe resolver al menos un día ISO válido)
            var dias = AgDiasVentana.Dias(dto.DiasVentana);
            if (dias.Count == 0)
            {
                return Result.Failure<AgContextoSeleccionDto>(
                    $"I1_VENTANA_VACIA: La modalidad '{key.Codigo}' de cohorte '{semestre}-{grupo}' no tiene ventana de días válida seleccionada (LUN-JUE..VIE-DOM).");
            }

            // Validar horas por día
            if (dto.HorasDia != 4 && dto.HorasDia != 8)
            {
                return Result.Failure<AgContextoSeleccionDto>(
                    $"I6_HORAS_DIA_INVALIDO: Horas/día ({dto.HorasDia}) inválido en modalidad '{key.Codigo}'. Debe ser 4 u 8 horas.");
            }

            // Validar jornada
            if (string.IsNullOrWhiteSpace(dto.Jornada))
            {
                return Result.Failure<AgContextoSeleccionDto>(
                    $"I1_JORNADA_VACIA: No se especificó jornada (MANANA, TARDE, NOCHE) para la modalidad '{key.Codigo}'.");
            }

            return Result.Success(dto);
        }
    }
}
