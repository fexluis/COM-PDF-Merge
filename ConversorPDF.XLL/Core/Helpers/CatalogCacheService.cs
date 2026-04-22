using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models.Academic;
using ConversorPDF.XLL.Core.Repositories;
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
                    string query = @"
                        SELECT * FROM sedes;
                        SELECT * FROM programas;
                        SELECT id, departamento, nombre, cedula, tipo, salario, rol, entidad_bancaria as EntidadBancaria, cuenta_bancaria as CuentaBancaria, telefono, correo, fecha_nacimiento as FechaDeNacimiento, contrato FROM terceros;";

                    using (var multi = db.QueryMultiple(query))
                    {
                        Sedes = multi.Read<Sede>().ToList();
                        Programas = multi.Read<Programa>().ToList();
                        Terceros = multi.Read<Tercero>().ToList();
                    }
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
                Sedes.Clear();
                Programas.Clear();
                Asignaturas.Clear();
                Terceros.Clear();
            }
        }
    }
}
