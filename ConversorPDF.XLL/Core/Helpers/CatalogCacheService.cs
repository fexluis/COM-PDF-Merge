using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Caché in-memory para los catálogos académicos inmutables o de rarísimo cambio.
    /// Reduce significativamente los tiempos de I/O de la DB SQLite Encriptada 
    /// al reabrir el Add-In de forma reiterada.
    /// </summary>
    public static class CatalogCacheService
    {
        private static bool _isLoaded = false;
        private static readonly object _lock = new object();

        public static List<Sede> Sedes { get; private set; } = new List<Sede>();
        public static List<Periodo> Periodos { get; private set; } = new List<Periodo>();
        public static List<Programa> Programas { get; private set; } = new List<Programa>();
        public static List<Asignatura> Asignaturas { get; private set; } = new List<Asignatura>();
        public static List<Tercero> Terceros { get; private set; } = new List<Tercero>();

        /// <summary>
        /// Asegura que los catálogos estén cargados en memoria usando una conexión externa. 
        /// Útil para optimizar I/O cuando el llamador ya tiene un contexto abierto.
        /// </summary>
        public static void EnsureLoaded(IDbConnection db)
        {
            if (_isLoaded) return;

            lock (_lock)
            {
                if (_isLoaded) return;

                LoggerHelper.Info("[CatalogCacheService] Inicializando caché de catálogos v1.2.50.0 (Shared Connection Mandatory)...");
                try
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    bool tieneVentana = PeriodosTieneVentana(db);
                    string queryPeriodos = tieneVentana
                        ? "SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, periodo_inicio as PeriodoInicio, periodo_fin as PeriodoFin, activo as Activo FROM periodos;"
                        : "SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, NULL as PeriodoInicio, NULL as PeriodoFin, activo as Activo FROM periodos;";

                    string query = $@"
                        {queryPeriodos}
                        SELECT * FROM sedes WHERE activo = 1;
                        SELECT * FROM programas;
                        SELECT *, 
                               entidad_bancaria as EntidadBancaria, 
                               cuenta_bancaria as CuentaBancaria, 
                               tipo_cuenta_bancaria as TipoCuentaBancaria,
                               correo_personal as CorreoPersonal,
                               fecha_nacimiento as FechaDeNacimiento 
                        FROM terceros;";

                    List<Periodo> periodos;
                    List<Sede> sedes;
                    List<Programa> programas;
                    List<Tercero> terceros;

                    using (var multi = db.QueryMultiple(query))
                    {
                        periodos = multi.Read<Periodo>().ToList();
                        sedes = multi.Read<Sede>().ToList();
                        programas = multi.Read<Programa>().ToList();
                        terceros = multi.Read<Tercero>().ToList();
                    }

                    Periodos = periodos;
                    Sedes = sedes;
                    Programas = programas;
                    Terceros = terceros;

                    sw.Stop();
                    LoggerHelper.Info($"[PERF] [CatalogCacheService] Carga de catálogos finalizada en {sw.ElapsedMilliseconds}ms.");

                    _isLoaded = true;
                    LoggerHelper.Info($"[CatalogCacheService] Caché lista. Sedes:{Sedes.Count}, Programas:{Programas.Count}, Terceros:{Terceros.Count}");
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[CatalogCacheService] Falló la carga del caché de catálogos (Shared)", ex);
                }
            }
        }

        public static void Invalidate()
        {
            lock (_lock)
            {
                _isLoaded = false;
                Periodos.Clear();
                Sedes.Clear();
                Programas.Clear();
                Asignaturas.Clear();
                Terceros.Clear();
            }
        }

        /// <summary>
        /// Actualiza o agrega un período en el caché en memoria de forma no destructiva (Fase 2080).
        /// </summary>
        public static void ActualizarPeriodo(Periodo actualizado)
        {
            if (actualizado == null) return;
            lock (_lock)
            {
                int i = Periodos.FindIndex(p => p.IdPeriodo == actualizado.IdPeriodo);
                if (i >= 0)
                {
                    Periodos[i].PeriodoNombre = actualizado.PeriodoNombre;
                    Periodos[i].PeriodoInicio = actualizado.PeriodoInicio;
                    Periodos[i].PeriodoFin = actualizado.PeriodoFin;
                    Periodos[i].Activo = actualizado.Activo;
                }
                else
                {
                    Periodos.Add(actualizado);
                }
            }
        }

        /// <summary>
        /// Elimina un período del caché en memoria sin invalidar el resto de catálogos (Fase 2080).
        /// </summary>
        public static void EliminarPeriodo(int idPeriodo)
        {
            lock (_lock)
            {
                var p = Periodos.FirstOrDefault(x => x.IdPeriodo == idPeriodo);
                if (p != null)
                {
                    Periodos.Remove(p);
                }
            }
        }

        private static bool PeriodosTieneVentana(IDbConnection db)
        {
            try
            {
                var cols = db.Query<string>("SELECT name FROM pragma_table_info('periodos');")
                             .Select(x => x.ToLowerInvariant())
                             .ToList();
                return cols.Contains("periodo_inicio") && cols.Contains("periodo_fin");
            }
            catch
            {
                return false;
            }
        }
    }
}
