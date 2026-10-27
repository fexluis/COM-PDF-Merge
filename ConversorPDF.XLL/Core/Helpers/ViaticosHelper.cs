using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;

namespace ConversorPDF.XLL.Core.Helpers
{
    public static class ViaticosHelper
    {
        public const string SIN_TERMINAL = "NO APLICA";
        public const string TERMINAL_TERRESTRE = "Terrestre";

        private static readonly object _terminalesLock = new object();
        private static List<TerminalCatalogoItem> _cacheTerminales;
        private static DateTime _cacheTerminalesExpira = DateTime.MinValue;

        /// <summary>
        /// Determina si una cadena de texto representa una terminal aérea o aeropuerto real,
        /// descartando etiquetas sintéticas o predeterminadas ("NO APLICA", "Terrestre", "Aéreo", "Aereo", vacíos).
        /// </summary>
        /// <param name="terminal">Texto representativo de la terminal o medio de transporte.</param>
        /// <returns>True si es una terminal aérea específica del catálogo.</returns>
        public static bool EsTerminalAerea(string terminal)
        {
            if (string.IsNullOrWhiteSpace(terminal)) return false;
            string t = terminal.Trim();
            return !string.Equals(t, SIN_TERMINAL, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, TERMINAL_TERRESTRE, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Aéreo", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Aereo", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Invalida la memoria caché del catálogo de terminales aéreas en memoria.
        /// </summary>
        public static void InvalidarCacheTerminales()
        {
            lock (_terminalesLock)
            {
                _cacheTerminales = null;
                _cacheTerminalesExpira = DateTime.MinValue;
            }
        }

        /// <summary>
        /// Obtiene el catálogo de terminales aéreas con sus topes canónicos (maxvalor),
        /// cargándolos desde SQLite una sola vez con almacenamiento en caché thread-safe.
        /// En caso de contingencia o fallo de base de datos, provee el catálogo canónico base en memoria.
        /// </summary>
        /// <returns>Lista de sólo lectura con los elementos del catálogo de terminales.</returns>
        public static IReadOnlyList<TerminalCatalogoItem> ObtenerCatalogoTerminales()
        {
            lock (_terminalesLock)
            {
                if (_cacheTerminales != null && DateTime.Now < _cacheTerminalesExpira)
                {
                    return _cacheTerminales;
                }

                var lista = new List<TerminalCatalogoItem>();
                try
                {
                    using (var ctx = SqliteHelper.GetContext())
                    {
                        var db = ctx.Open();
                        var items = Dapper.SqlMapper.Query<TerminalCatalogoItem>(db,
                            "SELECT id AS Id, terminal AS Terminal, COALESCE(maxvalor, 0) AS MaxValor FROM terminales ORDER BY id ASC");
                        if (items != null)
                        {
                            lista.AddRange(items);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[ViaticosHelper] Error al cargar catálogo de terminales desde SQLite: {ex.Message}");
                }

                if (lista.Count == 0)
                {
                    lista.Add(new TerminalCatalogoItem { Id = 1, Terminal = "NO APLICA", MaxValor = 0m });
                    lista.Add(new TerminalCatalogoItem { Id = 2, Terminal = "ANTIOQUIA (Rionegro)", MaxValor = 162634m });
                    lista.Add(new TerminalCatalogoItem { Id = 3, Terminal = "ARAUCA (Arauca)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 4, Terminal = "ATLÁNTICO (Barranquilla)", MaxValor = 162634m });
                    lista.Add(new TerminalCatalogoItem { Id = 5, Terminal = "BOLÍVAR (Cartagena)", MaxValor = 162634m });
                    lista.Add(new TerminalCatalogoItem { Id = 6, Terminal = "BOYACÁ (Paipa)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 7, Terminal = "CALDAS (Manizales)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 8, Terminal = "CAQUETÁ (Florencia)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 9, Terminal = "CASANARE (Yopal)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 10, Terminal = "CAUCA (Popayán)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 11, Terminal = "CESAR (Valledupar)", MaxValor = 50689m });
                    lista.Add(new TerminalCatalogoItem { Id = 12, Terminal = "CÓRDOBA (Montería)", MaxValor = 162634m });
                    lista.Add(new TerminalCatalogoItem { Id = 13, Terminal = "OTRO", MaxValor = 50689m });
                }

                _cacheTerminales = lista;
                _cacheTerminalesExpira = DateTime.Now.AddMinutes(10);
                return _cacheTerminales;
            }
        }

        /// <summary>
        /// Resuelve el tope monetario (maxvalor) e identificador de una terminal aérea en base al catálogo en memoria.
        /// Prioriza: 1) Id > 1, 2) Coincidencia exacta de nombre normalizado, 3) Contención de texto con mayor tope, 4) Fallback canónico 'OTRO'.
        /// Emite trazas de auditoría estructuradas con LoggerHelper.
        /// </summary>
        /// <param name="idTerminal">Id de la terminal (si es > 1).</param>
        /// <param name="nombreTerminal">Nombre o texto de la terminal aérea.</param>
        /// <param name="contextoLog">Etiqueta de contexto para trazabilidad de log (ej. 'Exporter', 'Processor').</param>
        /// <param name="idResuelto">Id del catálogo resuelto.</param>
        /// <param name="nombreResuelto">Nombre oficial de la terminal resuelta.</param>
        /// <param name="maxValor">Tope máximo oficial autorizado.</param>
        /// <returns>True si se resolvió una terminal aérea con tope positivo.</returns>
        public static bool ResolverTopeTerminal(
            int idTerminal,
            string nombreTerminal,
            string contextoLog,
            out int idResuelto,
            out string nombreResuelto,
            out decimal maxValor)
        {
            idResuelto = 1;
            nombreResuelto = SIN_TERMINAL;
            maxValor = 0m;

            var catalogo = ObtenerCatalogoTerminales();

            // 1. Intento por ID específico (> 1)
            if (idTerminal > 1)
            {
                var matchId = catalogo.FirstOrDefault(t => t.Id == idTerminal);
                if (matchId != null)
                {
                    idResuelto = matchId.Id;
                    nombreResuelto = matchId.Terminal;
                    maxValor = matchId.MaxValor;
                    LoggerHelper.Info($"[{contextoLog}] Terminal resuelta por Id={idTerminal}: '{nombreResuelto}', MaxValor={maxValor:N0}");
                    return true;
                }
            }

            string nombreBuscado = NormalizarNombre(nombreTerminal);
            if (string.IsNullOrWhiteSpace(nombreBuscado) || !EsTerminalAerea(nombreTerminal))
            {
                return false;
            }

            // 2. Intento por coincidencia exacta de nombre normalizado
            var matchExacto = catalogo.FirstOrDefault(t => t.Id > 1 && NormalizarNombre(t.Terminal) == nombreBuscado);
            if (matchExacto != null)
            {
                idResuelto = matchExacto.Id;
                nombreResuelto = matchExacto.Terminal;
                maxValor = matchExacto.MaxValor;
                LoggerHelper.Info($"[{contextoLog}] Terminal resuelta por coincidencia exacta '{nombreTerminal}' -> '{nombreResuelto}', MaxValor={maxValor:N0}");
                return true;
            }

            // 3. Intento por contención (Contains) priorizando mayor maxvalor
            var matchContiene = catalogo
                .Where(t => t.Id > 1 && (NormalizarNombre(t.Terminal).Contains(nombreBuscado) || nombreBuscado.Contains(NormalizarNombre(t.Terminal))))
                .OrderByDescending(t => t.MaxValor)
                .ThenBy(t => t.Id)
                .FirstOrDefault();

            if (matchContiene != null)
            {
                idResuelto = matchContiene.Id;
                nombreResuelto = matchContiene.Terminal;
                maxValor = matchContiene.MaxValor;
                LoggerHelper.Info($"[{contextoLog}] Terminal resuelta por contención '{nombreTerminal}' -> '{nombreResuelto}', MaxValor={maxValor:N0}");
                return true;
            }

            // 4. Fallback a 'OTRO' con advertencia en log
            var matchOtro = catalogo.FirstOrDefault(t => t.Terminal.Equals("OTRO", StringComparison.OrdinalIgnoreCase))
                            ?? catalogo.FirstOrDefault(t => t.Id == 13);

            if (matchOtro != null)
            {
                idResuelto = matchOtro.Id;
                nombreResuelto = matchOtro.Terminal;
                maxValor = matchOtro.MaxValor;
                LoggerHelper.Warning($"[{contextoLog}] No se encontró coincidencia directa para terminal aérea '{nombreTerminal}'. Aplicando fallback canónico '{nombreResuelto}' (Id={idResuelto}), MaxValor={maxValor:N0}");
                return true;
            }

            return false;
        }

        /// <summary>
        /// Resuelve los textos canónicos de transporte aéreo (horas de ida/regreso, requerimiento de tiquete y nombre de terminal/aeropuerto).
        /// Limpia etiquetas sintéticas residuales de entrada ("NO APLICA", "Terrestre", "Aéreo", "Aereo") para evitar nombres ambiguos,
        /// preservando terminales canónicas del catálogo como "OTRO".
        /// </summary>
        /// <param name="isTerminalAereo">Indicador de transporte aéreo (1 = Aéreo, 0 = Terrestre).</param>
        /// <param name="terminalAereaNombre">Nombre de la terminal aérea o aeropuerto.</param>
        /// <param name="docente">True si el formato corresponde a Docente de Cátedra.</param>
        /// <returns>Tupla con (HoraIda, HoraRegreso, Tiquete, Terminal).</returns>
        public static (string HoraIda, string HoraRegreso, string Tiquete, string Terminal) ResolverAereo(
            int isTerminalAereo, string terminalAereaNombre, bool docente)
        {
            bool aereo = isTerminalAereo == 1;
            string nombreLimpio = (terminalAereaNombre ?? "").Trim();
            if (string.Equals(nombreLimpio, SIN_TERMINAL, StringComparison.OrdinalIgnoreCase) 
                || string.Equals(nombreLimpio, TERMINAL_TERRESTRE, StringComparison.OrdinalIgnoreCase)
                || string.Equals(nombreLimpio, "Aéreo", StringComparison.OrdinalIgnoreCase)
                || string.Equals(nombreLimpio, "Aereo", StringComparison.OrdinalIgnoreCase))
            {
                nombreLimpio = "";
            }

            string terminal = aereo
                ? (string.IsNullOrWhiteSpace(nombreLimpio) ? (docente ? "Otros" : "Aéreo") : nombreLimpio)
                : (docente ? "Otros" : TERMINAL_TERRESTRE);

            string horaIda = aereo ? "7:00" : "";
            string horaRegreso = aereo ? "18:00" : "";
            string tiquete = aereo ? "SI" : "NO";

            return (horaIda, horaRegreso, tiquete, terminal);
        }

        /// <summary>
        /// Intenta parsear defensivamente un objeto o celda COM a decimal.
        /// Distingue explícitamente entre celdas vacías/nulas y valores numéricos válidos (incluido 0).
        /// Elimina el riesgo de NumberStyles.Any con AllowThousands que multiplicaba decimales ambiguos por 10 o 1000.
        /// </summary>
        /// <param name="val">Valor crudo extraído de Excel o base de datos.</param>
        /// <param name="resultado">Valor decimal obtenido si fue parseado exitosamente; 0m en caso contrario.</param>
        /// <param name="esVacioOInvalido">True si el valor de entrada era nulo, vacío o ilegible.</param>
        /// <returns>True si se obtuvo un valor decimal válido.</returns>
        public static bool TryParseDecimal(object val, out decimal resultado, out bool esVacioOInvalido)
        {
            resultado = 0m;
            esVacioOInvalido = true;
            if (val == null) return false;

            if (val is decimal dec)
            {
                resultado = dec;
                esVacioOInvalido = false;
                return true;
            }
            if (val is int i)
            {
                resultado = i;
                esVacioOInvalido = false;
                return true;
            }
            if (val is long l)
            {
                resultado = l;
                esVacioOInvalido = false;
                return true;
            }
            if (val is short s)
            {
                resultado = s;
                esVacioOInvalido = false;
                return true;
            }
            if (val is byte b)
            {
                resultado = b;
                esVacioOInvalido = false;
                return true;
            }
            if (val is double d)
            {
                if (double.IsNaN(d) || double.IsInfinity(d)) return false;
                try
                {
                    resultado = (decimal)d;
                    esVacioOInvalido = false;
                    return true;
                }
                catch (OverflowException) { return false; }
            }
            if (val is float f)
            {
                if (float.IsNaN(f) || float.IsInfinity(f)) return false;
                try
                {
                    resultado = (decimal)f;
                    esVacioOInvalido = false;
                    return true;
                }
                catch (OverflowException) { return false; }
            }

            string str = val.ToString().Trim();
            if (string.IsNullOrWhiteSpace(str)) return false;

            // Normalizar símbolos monetarios, espacios no separables (\u00A0) y espacios
            str = str.Replace("$", "")
                     .Replace("COP", "")
                     .Replace("PES", "")
                     .Replace("cop", "")
                     .Replace("\u00A0", "")
                     .Replace(" ", "")
                     .Trim();

            if (string.IsNullOrWhiteSpace(str)) return false;

            bool negativo = false;
            if (str.StartsWith("(") && str.EndsWith(")"))
            {
                negativo = true;
                str = str.Substring(1, str.Length - 2).Trim();
            }
            else if (str.StartsWith("-"))
            {
                negativo = true;
                str = str.Substring(1).Trim();
            }

            // Detección estricta de separadores sin AllowThousands
            int lastDot = str.LastIndexOf('.');
            int lastComma = str.LastIndexOf(',');

            if (lastDot >= 0 && lastComma >= 0)
            {
                if (lastComma > lastDot)
                {
                    // Formato estándar es-CO: 1.234,56 -> quitar puntos de miles y sustituir coma decimal por punto
                    str = str.Replace(".", "").Replace(",", ".");
                }
                else
                {
                    // Formato en-US: 1,234.56 -> quitar comas de miles
                    str = str.Replace(",", "");
                }
            }
            else if (lastComma >= 0)
            {
                int commaCount = 0;
                for (int idx = 0; idx < str.Length; idx++) if (str[idx] == ',') commaCount++;
                if (commaCount > 1)
                {
                    str = str.Replace(",", "");
                }
                else
                {
                    str = str.Replace(",", ".");
                }
            }
            else if (lastDot >= 0)
            {
                int dotCount = 0;
                for (int idx = 0; idx < str.Length; idx++) if (str[idx] == '.') dotCount++;
                if (dotCount > 1)
                {
                    str = str.Replace(".", "");
                }
                // Si hay un solo punto, se respeta como punto decimal invariante
            }

            if (decimal.TryParse(str, System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out decimal parsed))
            {
                resultado = negativo ? -parsed : parsed;
                esVacioOInvalido = false;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Parsea defensivamente un objeto o celda COM a decimal, tolerando tipos numéricos directos,
        /// formatos de moneda, separadores regionales (puntos/comas) y espacios.
        /// Retorna 0m en caso de valor nulo o inválido.
        /// </summary>
        /// <param name="val">Valor crudo extraído de Excel o base de datos.</param>
        /// <returns>Valor decimal parseado o 0 en caso de fallo.</returns>
        public static decimal ParseDecimal(object val)
        {
            TryParseDecimal(val, out decimal res, out _);
            return res;
        }

        public static DateTime ToDateTime(object val)
        {
            if (val == null) return DateTime.MinValue;
            if (val is DateTime dt) return dt;
            if (val is double d) return DateTime.FromOADate(d);
            if (double.TryParse(val?.ToString(), out double d2)) return DateTime.FromOADate(d2);
            if (DateTime.TryParse(val?.ToString(), out DateTime dt2)) return dt2;
            return DateTime.MinValue;
        }

        public static string SinTildes(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return texto;
            const string conTilde = "ÁÉÍÓÚáéíóú";
            const string sinTilde = "AEIOUaeiou";
            for (int i = 0; i < conTilde.Length; i++)
                texto = texto.Replace(conTilde[i], sinTilde[i]);
            return texto;
        }

        public static string LimpiarTexto(string t) => t?.Replace("\"", "").Replace(":", "").Replace("/", "-");

        /// <summary>
        /// Sanitiza la cédula removiendo todos los caracteres no dígitos (puntos, guiones, espacios).
        /// Permite comparaciones numéricas limpias entre cédulas con o sin formato (ej. "1.018.412.345" vs "1018412345").
        /// </summary>
        /// <param name="cedula">Cadena original con la cédula a sanitizar.</param>
        /// <returns>Cadena con los caracteres puramente numéricos de la cédula o vacía si el parámetro es nulo o espacios.</returns>
        public static string SanitizarCedula(string cedula)
        {
            if (string.IsNullOrWhiteSpace(cedula)) return string.Empty;
            return System.Text.RegularExpressions.Regex.Replace(cedula, @"\D", "");
        }

        /// <summary>
        /// Crea y retorna la ruta de la carpeta destino para los viáticos generados.
        /// Cambio (Fase 15): Se agregó HHmmss al timestamp para evitar colisión de archivos 
        /// al exportar viáticos múltiples veces en un mismo día.
        /// </summary>
        public static string CrearCarpetaDestino()
        {
            string folderName = "Viaticos-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), folderName);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }

        private static bool _avisoCedulaVacia;

        public static string NormalizarNombre(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string t = SinTildes(s).ToUpperInvariant();
            return string.Join(" ", t.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
        }

        /// <summary>
        /// Fase 233 y 234: Resuelve el aprobador efectivo. Aprobador habitual = cfg.Gerente/cfg.Cargo
        /// (dt_nombre/dt_cargo) con cédula cfg.CedulaGerente (corporativo_cedula = "Dir. Terr. Cédula").
        /// Si el comisionado es el aprobador habitual (por cédula o por nombre normalizado), se usa cfg.GerenteSustituto/cfg.CargoSustituto.
        /// </summary>
        public static void ResolverAprobador(ConversorPDF.XLL.Core.Models.ViaticosConfig cfg, string cedulaEmpleado,
            string nombreEmpleado, string etiquetaLog, out string aprobador, out string cargoAprobador)
        {
            aprobador = cfg?.Gerente ?? "";
            cargoAprobador = cfg?.Cargo ?? "";

            if (cfg == null) return;

            string cedEmp = SanitizarCedula(cedulaEmpleado);
            string cedApr = SanitizarCedula(cfg.CedulaGerente);

            bool coincideCedula = !string.IsNullOrWhiteSpace(cedEmp) && !string.IsNullOrWhiteSpace(cedApr)
                               && string.Equals(cedEmp, cedApr, StringComparison.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(cedApr) && !_avisoCedulaVacia)
            {
                _avisoCedulaVacia = true;
                LoggerHelper.Warning($"[{etiquetaLog}] Config.CedulaGerente vacía: la detección de auto-aprobación usará el nombre del aprobador habitual.");
            }

            bool coincideNombre = string.IsNullOrWhiteSpace(cedApr)
                               && NormalizarNombre(nombreEmpleado).Length > 0
                               && NormalizarNombre(nombreEmpleado) == NormalizarNombre(cfg.Gerente);

            bool esMismoAprobador = coincideCedula || coincideNombre;

            LoggerHelper.Debug($"[{etiquetaLog}] AprobadorHabitual -> Empleado Cedula='{cedEmp}', Nombre='{NormalizarNombre(nombreEmpleado)}'"
                + $" | Aprobador Cedula='{cedApr}', Nombre='{NormalizarNombre(cfg.Gerente)}'"
                + $" | SustitutoNombre='{cfg.GerenteSustituto}', SustitutoCargo='{cfg.CargoSustituto}'"
                + $" | PorCedula={coincideCedula}, PorNombre={coincideNombre}");

            if (!esMismoAprobador) return;

            if (!string.IsNullOrWhiteSpace(cfg.GerenteSustituto))
            {
                aprobador = cfg.GerenteSustituto;
                cargoAprobador = cfg.CargoSustituto ?? "";
                LoggerHelper.Info($"[{etiquetaLog}] Sustitución activada (cédula {cedEmp}, nombre {NormalizarNombre(nombreEmpleado)}). Nuevo aprobador: '{aprobador}' ({cargoAprobador}).");
            }
            else
            {
                aprobador = "";
                cargoAprobador = "";
                LoggerHelper.Warning($"[{etiquetaLog}] Comisionado = aprobador habitual (cédula {cedEmp}) y sustituto ausente/incompleto; aprobador vacío para evitar auto-aprobación.");
            }
        }
    }
}