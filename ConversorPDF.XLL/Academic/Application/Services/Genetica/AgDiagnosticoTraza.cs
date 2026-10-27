using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Gestiona la traza diagnóstica secuencial y estructurada (P01..Pnn) para las operaciones
    /// del Algoritmo Genético en Módulo S (Fase 277).
    /// Proporciona correlación, temporización, inspección de SQL y parámetros, conteo de filas
    /// y tipificación de causas raíz sin afectar el rendimiento en producción.
    /// </summary>
    public sealed class AgDiagnosticoTraza
    {
        private readonly Action<string> _ui;
        private readonly Stopwatch _reloj = Stopwatch.StartNew();
        private int _paso;

        /// <summary> Indica si la emisión de trazas detalladas está habilitada (Fase 596: opt-in). </summary>
        public bool Habilitado { get; set; } = false;

        /// <summary> Nombre representativo de la operación en curso. </summary>
        public string Operacion { get; }

        /// <summary> Identificador de correlación unívoco para la sesión de diagnóstico. </summary>
        public string Correlacion { get; }

        private readonly Dictionary<string, int> _conteos = new Dictionary<string, int>();

        /// <summary> Inicializa una nueva instancia de la traza diagnóstica. </summary>
        /// <param name="operacion">Nombre de la operación.</param>
        /// <param name="uiLog">Delegado opcional para visualización en el componente de UI.</param>
        /// <param name="correlacion">Identificador de correlación opcional.</param>
        public AgDiagnosticoTraza(string operacion, Action<string> uiLog = null, string correlacion = null)
        {
            Operacion = string.IsNullOrWhiteSpace(operacion) ? "ModuloS" : operacion.Trim();
            _ui = uiLog;
            Correlacion = string.IsNullOrWhiteSpace(correlacion)
                ? $"S-{DateTime.Now:HHmmss}-{Guid.NewGuid().ToString("N").Substring(0, 4)}"
                : correlacion.Trim();
        }

        /// <summary> Emite el encabezado de inicio de la secuencia de diagnóstico. </summary>
        public void Inicio(string detalle = null)
        {
            string sufijo = string.IsNullOrWhiteSpace(detalle) ? string.Empty : $" | {detalle}";
            Emit("INFO", $"=== INICIO {Operacion} [{Correlacion}]{sufijo} ===");
        }

        /// <summary> Registra un nuevo paso numerado (P01..Pnn) con su tiempo transcurrido en milisegundos. </summary>
        public void Paso(string nombre, string detalle = null)
        {
            _paso++;
            string sufijo = string.IsNullOrWhiteSpace(detalle) ? string.Empty : $" - {detalle}";
            Emit("DEBUG", $"[P{_paso:00}] {nombre}{sufijo} (t={_reloj.ElapsedMilliseconds}ms)");
        }

        /// <summary> Registra una clave y valor de datos para inspección contextual. </summary>
        public void Dato(string clave, object valor)
        {
            Emit("DEBUG", $"        · {clave} = {(valor ?? "<null>")}");
        }

        /// <summary> Registra una sentencia SQL normalizada y los parámetros asociados. </summary>
        public void Sql(string nombre, string sql, object parametros = null)
        {
            string sqlNorm = NormalizarSql(sql);
            string paramStr = Serializar(parametros);
            Emit("DEBUG", $"        · SQL[{nombre}]: {sqlNorm} | params: {paramStr}");
        }

        /// <summary> Registra el conteo de registros obtenidos evaluando umbral mínimo (Fase 596). </summary>
        public void Conteo(string etiqueta, int obtenidos, int minimo = 1)
        {
            _conteos[etiqueta] = obtenidos;
            bool cumple = obtenidos >= minimo;
            if (!cumple)
            {
                Emit("WARN", $"· {etiqueta}: {obtenidos} ALERTA (<{minimo})");
            }
            else if (Habilitado)
            {
                Emit("DEBUG", $"· {etiqueta}: {obtenidos} OK");
            }
        }

        /// <summary> Registra un mensaje de advertencia contextual. </summary>
        public void Advertencia(string clave, string mensaje)
        {
            Emit("WARN", $"        · [ALERTA] {clave}: {mensaje}");
        }

        /// <summary> Realiza un volcado estructurado del estado de <see cref="AgContexto"/>. </summary>
        public void Contexto(AgContexto c)
        {
            if (c == null)
            {
                Dato("AgContexto", "<null>");
                return;
            }

            Dato("IdPeriodo/IdPrograma/IdAulaCorrida", $"{c.IdPeriodo}/{c.IdPrograma}/{c.IdAulaCorrida}");
            Dato("Subunidad", string.IsNullOrEmpty(c.Subunidad) ? "<global>" : c.Subunidad);
            Dato("Ventana", c.Ventana == null
                ? "<null>"
                : $"{c.Ventana.Origen} {c.Ventana.FechaInicio:yyyy-MM-dd} -> {c.Ventana.FechaFin:yyyy-MM-dd} ({c.Ventana.DuracionDias}d)");

            Conteo("Aulas", c.Aulas?.Count ?? 0);
            Conteo("Bloques", c.Bloques?.Count ?? 0);
            Conteo("PoolDocentes", c.PoolDocentes?.Count ?? 0);
            Conteo("Docentes", c.Docentes?.Count ?? 0);
            Conteo("Asignaturas", c.Asignaturas?.Count ?? 0);
            Conteo("AulaAsignaturas", c.AulaAsignaturas?.Count ?? 0);
            Conteo("Festivos", c.Festivos?.Count ?? 0);
            Conteo("Presenciales", c.Presenciales?.Count ?? 0);
            Dato("PoolPorAsignatura(claves)", c.PoolPorAsignatura?.Count ?? 0);
        }

        /// <summary> Última causa raíz registrada durante el diagnóstico. </summary>
        public AgCausaFallo UltimaCausa { get; private set; } = AgCausaFallo.Ninguna;

        /// <summary> Registra el resultado general de una fase o validación. </summary>
        public void Resultado(bool ok, string detalle)
        {
            string nivel = ok ? "INFO" : "WARN";
            Emit(nivel, $"        => {(ok ? "[APROBADO]" : "[BLOQUEADO]")} {detalle}");
        }

        /// <summary> Registra el código de causa raíz sin excepción asociada (Fase 284). </summary>
        public void Causa(AgCausaFallo causa, string detalle = null)
        {
            if (causa == AgCausaFallo.Ninguna) return;
            UltimaCausa = causa;
            string sufijo = string.IsNullOrWhiteSpace(detalle) ? string.Empty : $": {detalle}";
            Emit("WARNING", $"        · causa={causa} (CAUSA: {causa}){sufijo}");
        }

        /// <summary> Registra una falla o excepción vinculando la causa raíz formal. </summary>
        public void Fallo(string paso, Exception ex, AgCausaFallo causa = AgCausaFallo.Ninguna)
        {
            if (causa != AgCausaFallo.Ninguna)
            {
                UltimaCausa = causa;
            }
            string causaStr = causa != AgCausaFallo.Ninguna ? $" [causa={causa}]" : string.Empty;
            string msg = $"[FALLO] En paso '{paso}'{causaStr}: {(ex != null ? ex.Message : "Operación no completada")}";
            Emit("ERROR", msg);
            // Fase 616: caller correlacionado y detalle de excepción sólo en modo verboso
            if (ex != null && LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Error($"[AgDiagnosticoTraza] Detalle de excepción en {paso}", ex, $"Ag:{Correlacion}");
            }
        }

        /// <summary> Emite el pie de cierre de la traza con el tiempo total transcurrido y resumen de contadores (Fase 596). </summary>
        public void Fin(bool ok, string detalle = null)
        {
            string estado = ok ? "EXITOSO" : "CON FALLAS";
            string conteos = _conteos.Count == 0 ? null
                : string.Join(", ", System.Linq.Enumerable.Select(System.Linq.Enumerable.OrderBy(_conteos, k => k.Key), kv => $"{kv.Key}={kv.Value}"));
            string cuerpo = string.Join(" | ", System.Linq.Enumerable.Where(new[] { detalle, conteos }, s => !string.IsNullOrWhiteSpace(s)));
            string sufijo = string.IsNullOrWhiteSpace(cuerpo) ? string.Empty : $" | {cuerpo}";
            Emit("INFO", $"=== FIN {Operacion} [{Correlacion}] estado={estado} total={_reloj.ElapsedMilliseconds}ms{sufijo} ===");
        }

        /// <summary> Normaliza el formato de una cadena SQL colapsando saltos de línea y espacios redundantes (Fase 482). </summary>
        public static string NormalizarSql(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return string.Empty;
            return Regex.Replace(sql.Trim(), @"\s+", " ");
        }

        /// <summary> Serializa de forma legible y segura un objeto de parámetros (anónimo, diccionario o POCO). </summary>
        public static string Serializar(object parametros)
        {
            if (parametros == null) return "<ninguno>";

            if (parametros is IDictionary dict)
            {
                var partesDict = new List<string>();
                foreach (DictionaryEntry de in dict)
                {
                    partesDict.Add($"{de.Key}={de.Value ?? "null"}");
                }
                return string.Join(", ", partesDict);
            }

            var props = parametros.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            if (props.Length == 0) return parametros.ToString();

            var partes = new List<string>();
            foreach (var p in props)
            {
                try
                {
                    object val = p.GetValue(parametros, null);
                    partes.Add($"{p.Name}={(val ?? "null")}");
                }
                catch
                {
                    partes.Add($"{p.Name}=<err>");
                }
            }
            return string.Join(", ", partes);
        }

        private void Emit(string nivel, string msg)
        {
            // Fase 596: marca temporal/nivel única; el correlativo va como caller (elimina el prefijo [Emit])
            string caller = $"Ag:{Correlacion}";
            bool esDetalle = string.Equals(nivel, "DEBUG", StringComparison.OrdinalIgnoreCase);
            switch (nivel.ToUpperInvariant())
            {
                case "DEBUG":
                    if (Habilitado) LoggerHelper.Debug($"        · {msg}", caller);
                    break;
                case "INFO":
                    LoggerHelper.Info(msg, caller);
                    break;
                case "WARN":
                case "WARNING":
                    LoggerHelper.Warning(msg, caller);
                    break;
                case "ERROR":
                    LoggerHelper.Error(msg, null, caller);
                    break;
                default:
                    LoggerHelper.Info(msg, caller);
                    break;
            }

            // Fase 604: la UI refleja siempre las etapas (INFO/WARN/ERROR); el detalle DEBUG es opt-in
            if (_ui != null && (!esDetalle || Habilitado))
            {
                try
                {
                    _ui($"[{Correlacion}] {msg}");
                }
                catch
                {
                    // Evitar que excepciones de visualización interrumpan la ejecución
                }
            }
        }
    }
}
