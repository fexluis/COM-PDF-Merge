using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación para analizar conflictos entre las sesiones de un docente,
    /// teniendo en cuenta fechas, horarios (franjas horarias), sedes y salvedades confirmadas.
    /// </summary>
    public static class SesionConflictAnalyzer
    {
        /// <summary>
        /// Analiza la lista de sesiones de un docente y determina el estado de conflicto de cada una.
        /// Retorna un diccionario donde la clave es el IdSesion y el valor es su EstadoConflicto más crítico.
        /// </summary>
        public static Dictionary<int, EstadoConflicto> Analizar(
            List<SesionListItem> sesionesDelDocente,
            IReadOnlyCollection<SalvedadConflicto> salvedades)
        {
            var resultados = new Dictionary<int, EstadoConflicto>();

            // Ignorar modalidad Plataforma (o legada Virtual) y sin fecha
            var conFecha = sesionesDelDocente
                .Where(s => s.FechaSesion.HasValue && !ModalidadTutoriaCatalog.EsPlataforma(s.Modalidad))
                .OrderBy(s => s.FechaSesion)
                .ToList();

            // Inicializar todas las procesadas en Ninguno
            foreach (var s in conFecha)
            {
                if (!resultados.ContainsKey(s.IdSesion))
                    resultados[s.IdSesion] = EstadoConflicto.Ninguno;
            }

            for (int i = 0; i < conFecha.Count; i++)
            {
                for (int j = i + 1; j < conFecha.Count; j++)
                {
                    var a = conFecha[i];
                    var b = conFecha[j];

                    // Si no son la misma fecha, no hay conflicto.
                    if (a.FechaSesion.Value.Date != b.FechaSesion.Value.Date)
                        continue;

                    // Si son la misma sede, no hay conflicto (un docente puede dar dos clases en la misma sede seguidas).
                    if (string.Equals(a.SedeCiudad ?? "", b.SedeCiudad ?? "", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Llegados a este punto: Misma fecha, distinta sede. 
                    // Revisar si existe salvedad:
                    if (ExisteSalvedad(a.IdSesion, b.IdSesion, salvedades))
                    {
                        RegistrarEstado(resultados, a.IdSesion, EstadoConflicto.Salvado);
                        RegistrarEstado(resultados, b.IdSesion, EstadoConflicto.Salvado);
                        continue;
                    }

                    // Determinar si hay traslape horario
                    var franjaA = HorarioSesionParser.TryParse(a.HorarioSesion);
                    var franjaB = HorarioSesionParser.TryParse(b.HorarioSesion);

                    bool hayTraslape = true; // Por defecto asumimos peor caso si no se puede parsear

                    if (franjaA.IsSuccess && franjaB.IsSuccess)
                    {
                        hayTraslape = franjaA.Value.SeSolapaCon(franjaB.Value);
                    }

                    if (hayTraslape)
                    {
                        RegistrarEstado(resultados, a.IdSesion, EstadoConflicto.ConflictoReal);
                        RegistrarEstado(resultados, b.IdSesion, EstadoConflicto.ConflictoReal);
                    }
                    else
                    {
                        RegistrarEstado(resultados, a.IdSesion, EstadoConflicto.RevisarViabilidad);
                        RegistrarEstado(resultados, b.IdSesion, EstadoConflicto.RevisarViabilidad);
                    }
                }
            }

            return resultados;
        }

        private static bool ExisteSalvedad(int idSesionA, int idSesionB, IReadOnlyCollection<SalvedadConflicto> salvedades)
        {
            if (salvedades == null || !salvedades.Any())
                return false;

            int minId = Math.Min(idSesionA, idSesionB);
            int maxId = Math.Max(idSesionA, idSesionB);

            return salvedades.Any(s => s.IdSesionA == minId && s.IdSesionB == maxId);
        }

        private static void RegistrarEstado(Dictionary<int, EstadoConflicto> dict, int idSesion, EstadoConflicto nuevoEstado)
        {
            if (!dict.ContainsKey(idSesion))
            {
                dict[idSesion] = nuevoEstado;
                return;
            }

            var estadoActual = dict[idSesion];

            // Jerarquía de estados: Ninguno < Salvado < RevisarViabilidad < ConflictoReal
            // Queremos quedarnos con el estado más crítico.
            if (estadoActual == EstadoConflicto.ConflictoReal)
                return; // Ya está en el nivel máximo, no lo degradamos

            if (nuevoEstado == EstadoConflicto.ConflictoReal)
            {
                dict[idSesion] = nuevoEstado;
                return;
            }

            if (estadoActual == EstadoConflicto.RevisarViabilidad)
                return;

            if (nuevoEstado == EstadoConflicto.RevisarViabilidad)
            {
                dict[idSesion] = nuevoEstado;
                return;
            }

            if (estadoActual == EstadoConflicto.Salvado)
                return;

            if (nuevoEstado == EstadoConflicto.Salvado)
            {
                dict[idSesion] = nuevoEstado;
            }
        }
    }
}
