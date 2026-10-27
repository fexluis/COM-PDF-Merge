using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio de verificación temprana de esquema relacional para SQLite (Fase 278).
    /// Detecta anomalías como columnas faltantes o tablas inexistentes mediante PRAGMA table_info
    /// antes de ejecutar sentencias de negocio, evitando excepciones y abortos transaccionales silenciosos.
    /// </summary>
    internal static class AgEsquemaDiagnostico
    {
        /// <summary>
        /// Inspecciona la existencia de columnas requeridas en una tabla de SQLite.
        /// </summary>
        /// <param name="db">Conexión activa a la base de datos.</param>
        /// <param name="traza">Instancia de traza diagnóstica opcional.</param>
        /// <param name="tabla">Nombre de la tabla a verificar.</param>
        /// <param name="requeridas">Lista de nombres de columnas obligatorias.</param>
        /// <returns>Lista de columnas ausentes (vacía si el esquema es 100% conforme).</returns>
        public static List<string> Verificar(IDbConnection db, AgDiagnosticoTraza traza, string tabla, params string[] requeridas)
        {
            if (db == null) throw new ArgumentNullException(nameof(db));
            if (string.IsNullOrWhiteSpace(tabla)) throw new ArgumentException("El nombre de la tabla no puede ser nulo o vacío.", nameof(tabla));

            try
            {
                var columnasDb = db.Query<string>($"SELECT name FROM pragma_table_info('{tabla}');")
                                   .Select(c => c.ToLowerInvariant())
                                   .ToList();

                var columnasSet = new HashSet<string>(columnasDb, StringComparer.OrdinalIgnoreCase);

                if (columnasSet.Count == 0)
                {
                    string msgInexistente = $"Tabla '{tabla}' no existe o no tiene columnas definidas en la base de datos.";
                    traza?.Resultado(false, $"Esquema '{tabla}': TABLA INEXISTENTE"); // Fase 609
                    traza?.Dato(tabla, $"TABLA INEXISTENTE [FALTA]");
                    traza?.Fallo("VerificarEsquema", new InvalidOperationException(msgInexistente), AgCausaFallo.EsquemaInvalido);
                    LoggerHelper.Warning($"[AgEsquemaDiagnostico] {msgInexistente}");
                    return requeridas != null ? requeridas.ToList() : new List<string>();
                }

                var faltantes = new List<string>();
                if (requeridas != null)
                {
                    foreach (var req in requeridas)
                    {
                        if (!string.IsNullOrWhiteSpace(req) && !columnasSet.Contains(req.Trim()))
                        {
                            faltantes.Add(req.Trim());
                        }
                    }
                }

                if (faltantes.Count == 0)
                {
                    traza?.Dato(tabla, $"OK ({columnasSet.Count} columnas verificadas)"); // Fase 611: detalle técnico, resumen agrupado en el invocador
                }
                else
                {
                    string listaFaltantes = string.Join(", ", faltantes);
                    traza?.Resultado(false, $"Esquema '{tabla}': FALTAN COLUMNAS {listaFaltantes}"); // Fase 609
                    traza?.Dato(tabla, $"FALTAN COLUMNAS: {listaFaltantes}");
                    traza?.Fallo("VerificarEsquema", new MissingFieldException(tabla, listaFaltantes), AgCausaFallo.EsquemaInvalido);
                    LoggerHelper.Warning($"[AgEsquemaDiagnostico] [ALERTA DE ESQUEMA] Tabla '{tabla}' carece de columnas requeridas: {listaFaltantes}");
                }

                return faltantes;
            }
            catch (Exception ex)
            {
                traza?.Fallo($"VerificarEsquema({tabla})", ex, AgCausaFallo.EsquemaInvalido);
                LoggerHelper.Error($"[AgEsquemaDiagnostico] Error consultando pragma_table_info para tabla '{tabla}'", ex);
                return requeridas != null ? requeridas.ToList() : new List<string>();
            }
        }
    }
}
