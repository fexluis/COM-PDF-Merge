using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Extrae y valida sesiones desde una representación abstracta del grid (GridRowDto).
    /// Permite testear la lógica de extracción sin dependencia de WinForms.
    /// </summary>
    public static class GridToSesionMapper
    {
        /// <summary>
        /// Extrae sesiones de una lista de DTOs de filas del grid.
        /// </summary>
        /// <param name="rows">DTOs de filas del DataGridView.</param>
        /// <param name="idTerceroMaster">IdTercero seleccionado en el combo.</param>
        /// <param name="idVinculacion">Id de vinculación actual.</param>
        /// <returns>Lista de sesiones validadas.</returns>
        /// <exception cref="InvalidOperationException">Si hay inconsistencia de integridad.</exception>
        public static List<Sesion> Extract(
            IReadOnlyList<GridRowDto> rows, 
            int idTerceroMaster, 
            int idVinculacion)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            // FIX-0.1b: Conteo visual INDEPENDIENTE del loop de extracción.
            // Esta es la "fuente de verdad visual" (Capa de validación de integridad).
            int filasVisualesMarcadas = 0;
            foreach (var row in rows)
            {
                if (!row.IsNewRow && row.IsChecked)
                    filasVisualesMarcadas++;
            }

            LoggerHelper.Info($"[GridToSesionMapper] Iniciando extracción. Filas marcadas visualmente: {filasVisualesMarcadas} de {rows.Count} totales.");

            var sesiones = new List<Sesion>();

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];

                if (row.IsNewRow)
                {
                    // LoggerHelper.Info($"[GridToSesionMapper] Fila {i}: IsNewRow, omitiendo.");
                    continue;
                }

                if (!row.IsChecked) continue;

                // Validación de Tag (Modelo asociado)
                if (row.Tag == null || !(row.Tag is Sesion tagSesion))
                {
                    string extraInfo = row.Tag == null ? "Tag es nulo" : $"Tipo inesperado: {row.Tag.GetType().Name}";
                    LoggerHelper.Warning($"[GridToSesionMapper] Fila {i + 1}: marcada pero {extraInfo}.");
                    throw new InvalidOperationException(
                        $"Fila {i + 1}: No se encontró el modelo de sesión asociado ({extraInfo}). " +
                        "Por favor, haga clic en [Generar Sesiones] nuevamente para resincronizar los datos del grid.");
                }

                // Validación de sincronización de Tercero y Datos Base
                if (tagSesion.IdTercero != idTerceroMaster)
                {
                    LoggerHelper.Warning($"[GridToSesionMapper] Fila {i+1}: Desfase de Tercero ({tagSesion.IdTercero} != {idTerceroMaster})");
                    throw new InvalidOperationException(
                        $"Fila {i + 1}: El docente de esta sesión no coincide con el seleccionado. " +
                        "Regenere para sincronizar.");
                }

                // Opción A: Sincronizar el modelo asociado (Tag) con los valores válidos editados manualmente en la fila
                tagSesion.FechaSesion = row.FechaSesion;
                tagSesion.Horas = row.Horas;
                tagSesion.HorarioSesion = row.HorarioSesion;
                tagSesion.Viatica = row.Viatica;

                sesiones.Add(new Sesion
                {
                    IdVinculacion = idVinculacion,
                    IdTercero = tagSesion.IdTercero,
                    FechaSesion = row.FechaSesion,
                    HorarioSesion = row.HorarioSesion,
                    Horas = row.Horas,
                    Viatica = row.Viatica
                });
            }

            // FIX-0.1d: Validación real de integridad por conteo independiente
            if (sesiones.Count != filasVisualesMarcadas)
            {
                LoggerHelper.Error($"[GridToSesionMapper] FALLO DE INTEGRIDAD REAL: Visual={filasVisualesMarcadas}, Extraídas={sesiones.Count}.");
                throw new InvalidOperationException(
                    $"ERROR DE INTEGRIDAD: Se detectaron {filasVisualesMarcadas} sesiones visibles pero " +
                    $"solo {sesiones.Count} se extrajeron correctamente. Por favor, regenere las sesiones.");
            }

            return sesiones;
        }
    }
}
