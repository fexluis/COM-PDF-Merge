using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador seguro para restaurar las 3 dimensiones de configuración académica (Jornada, HorasDia, DiasVentana)
    /// desde un <see cref="AgContextoSeleccionDto"/> hacia un <see cref="AgNodoArbolKey"/> del árbol jerárquico (Fase 1203).
    /// Garantiza simetría bidireccional completa (DTO ↔ Key) e incorpora fallback defensivo ante perfiles legacy con dimensiones vacías.
    /// Internal para prevenir el registro automático de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgNodoArbolKeyMapper
    {
        /// <summary>
        /// Aplica y normaliza las dimensiones del DTO sobre la clave de nodo indicada.
        /// Si alguna dimensión es nula o inválida, asigna el valor canónico por defecto (MANANA, 4h, MAR-VIE)
        /// y retorna una falla descriptiva para trazabilidad y diagnóstico sin abortar la carga.
        /// </summary>
        /// <param name="key">Clave de nodo de nivel 6 (Modalidad) a mutar.</param>
        /// <param name="dto">DTO fuente persistido en base de datos.</param>
        /// <returns>
        /// <see cref="Result.Success"/> si las 3 dimensiones eran válidas y completas;
        /// de lo contrario <see cref="Result.Failure"/> indicando el fallback aplicado.
        /// </returns>
        public static Result AplicarDimensiones(AgNodoArbolKey key, AgContextoSeleccionDto dto)
        {
            if (key == null)
            {
                return Result.Failure("I1_KEY_NULA: Nodo de árbol nulo.");
            }

            if (dto == null)
            {
                return Result.Failure("I1_DTO_NULO: DTO de contexto nulo.");
            }

            string jornada = (dto.Jornada ?? string.Empty).Trim().ToUpperInvariant();
            int horasDia = dto.HorasDia;
            string diasVentana = (dto.DiasVentana ?? string.Empty).Trim().ToUpperInvariant();

            bool jornadaValida = !string.IsNullOrWhiteSpace(jornada);
            bool horasValidas = horasDia == 4 || horasDia == 8;
            bool ventanaValida = AgDiasVentana.Dias(diasVentana).Count > 0;

            if (jornadaValida && horasValidas && ventanaValida)
            {
                key.Jornada = jornada;
                key.HorasDia = horasDia;
                key.DiasVentana = diasVentana;
                return Result.Success();
            }

            // Fallback defensivo para perfiles guardados previamente sin dimensiones explícitas
            key.Jornada = jornadaValida ? jornada : "MANANA";
            key.HorasDia = horasValidas ? horasDia : 4;
            key.DiasVentana = ventanaValida ? diasVentana : "MAR-VIE";

            return Result.Failure(
                $"I1_RESTAURACION_INCOMPLETA: Dimensiones inválidas o incompletas en '{dto.AulaRef}'. Se aplicó fallback (Jornada={key.Jornada}, Horas={key.HorasDia}, Ventana={key.DiasVentana}).");
        }
    }
}
