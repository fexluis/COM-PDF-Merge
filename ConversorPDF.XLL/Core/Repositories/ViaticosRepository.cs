using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using Dapper;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Services;
namespace ConversorPDF.XLL.Core.Repositories
{
    public class ViaticosRepository : IViaticosRepository
    {
        private readonly IExcelRepository _excelRepository;
        private readonly Dictionary<string, Empleado> _cacheEmp = new Dictionary<string, Empleado>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _cacheCiu = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _cacheDep = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _cacheFullDep = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _cacheFullDepLock = new object();
        private bool _empIndexed;

        /// <summary>
        /// Fase 27.2: Invalida el caché de FullDependencia tras escribir en cdpDependencias.
        /// </summary>
        public static void InvalidarCacheDependenciaFull()
        {
            lock (_cacheFullDepLock) { _cacheFullDep.Clear(); }
        }

        private const string SqlCdpPorCdp = "SELECT * FROM cdpDependencias WHERE cdpnum = @Cdp";
        private const string SqlCdpOrden =
            @"ORDER BY IFNULL(activo, 0) DESC,
                      CASE WHEN idperiodo = (SELECT id_periodo FROM config WHERE id = 1) THEN 1 ELSE 0 END DESC,
                      id DESC
              LIMIT 1";

        public ViaticosRepository(IExcelRepository excelRepository)
        {
            _excelRepository = excelRepository;
        }

        public ViaticosConfig GetConfig(DateTime? start = null, DateTime? end = null, string cdpOverride = null)
        {
            var config = new ViaticosConfig();

            try
            {
                var db = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext().Open();
                string subunidadConfig = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
                if (string.IsNullOrWhiteSpace(subunidadConfig))
                    subunidadConfig = Dapper.SqlMapper.QueryFirstOrDefault<string>(db, "SELECT subunidad FROM config LIMIT 1");

                var dbCfg = Dapper.SqlMapper.QuerySingleOrDefault<dynamic>(db,
                    "SELECT * FROM config WHERE id = 1 AND (subunidad IS NULL OR subunidad = '' OR subunidad = @Subunidad)",
                    new { Subunidad = subunidadConfig });

                if (dbCfg == null)
                {
                    dbCfg = Dapper.SqlMapper.QuerySingleOrDefault<dynamic>(db, "SELECT * FROM config WHERE id = 1");
                    string subCfg = dbCfg?.subunidad?.ToString();
                    if (!string.IsNullOrWhiteSpace(subCfg) && !string.IsNullOrWhiteSpace(subunidadConfig)
                        && !string.Equals(subCfg.Trim(), subunidadConfig.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        LoggerHelper.Warning(
                            $"[ViaticosRepository] config.subunidad='{subCfg}' no coincide con la sede activa '{subunidadConfig}'. Se usó config por fallback.");
                    }
                }

                if (dbCfg == null) return null;

                var rowCfg = dbCfg as System.Collections.Generic.IDictionary<string, object>;
                if (rowCfg == null)
                {
                    LoggerHelper.Warning("[ViaticosRepository] La fila de config no expone IDictionary<string, object>; se aplicarán valores por defecto.");
                }

                Func<string, object> raw = k => (rowCfg != null && rowCfg.ContainsKey(k)) ? rowCfg[k] : null;
                Func<string, string> col = k => raw(k)?.ToString() ?? "";

                if (start.HasValue) config.FechaInicio = start.Value;
                else if (DateTime.TryParse(col("fecha_inicio"), out DateTime fi)) config.FechaInicio = fi;

                if (end.HasValue) config.FechaFin = end.Value;
                else if (DateTime.TryParse(col("fecha_fin"), out DateTime ff)) config.FechaFin = ff;

                config.Cdp = (!string.IsNullOrWhiteSpace(cdpOverride) ? cdpOverride : col("cdp")).Trim();
                config.SalarioDecreto = raw("salario_decreto_613_2025");
                config.ViaticoA = raw("viatico_diario_tarifa_a");
                config.ViaticoB = raw("viatico_diario_tarifa_b");
                config.PorcA = raw("porcentaje_decreto_a");
                config.PorcB = raw("porcentaje_decreto_b");

                config.NombreDT = col("dt_nombre");
                config.CargoDT = col("dt_cargo");
                config.Gerente = !string.IsNullOrWhiteSpace(config.NombreDT) ? config.NombreDT : col("corporativo_nombre");
                config.Cargo = !string.IsNullOrWhiteSpace(config.CargoDT) ? config.CargoDT : col("corporativo_cargo");
                if (string.IsNullOrWhiteSpace(config.NombreDT))
                {
                    LoggerHelper.Warning($"[ViaticosRepository] config.dt_nombre vacío para la sede '{subunidadConfig}'; se usó corporativo_nombre como aprobador.");
                }
                config.CedulaGerente = col("corporativo_cedula");
                config.GerenteSustituto = col("corporativo_nombre");
                config.CargoSustituto = col("corporativo_cargo");
                config.DeptoBase = col("departamento_base");
                config.CiudadBase = col("ciudad_base");
                config.AeropuertoBase = col("aeropuerto_base");

                config.CarpetaDestino = ViaticosHelper.CrearCarpetaDestino();
                LoadDepInfo(config, db, subunidadConfig);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[ViaticosRepository] Error al obtener config de la base de datos.", ex);
                return null;
            }

            return config;
        }
        /// <summary>
        /// Carga información complementaria de la dependencia asociada al CDP.
        /// Fase 1.3/2.1/2.3: Filtrado por subunidad null-safe, estado activo y período vigente.
        /// Fase 3.1/4.1: Fallback de búsqueda global validando pertenencia de subunidad para evitar firmas cruzadas entre sedes.
        /// Fase 3.3/5.3: Reutilización directa del parámetro subunidad obtenido en GetConfig.
        /// Fase 4.7/5.1: Saneamiento de cadenas con Trim() de entrada y salida para evitar espacios espurios.
        /// Fase 5.4: Advertencia de trazabilidad cuando la sede activa es indeterminada.
        /// Fase 6.2/22.4: Resolución de FullDependencia exclusivamente desde SQLite (cdpDependencias.dependencia_full).
        /// Fase 10.2/13.2: Log de advertencia diagnóstica detallado con la sede cuando config.cdp viene nulo o vacío.
        /// </summary>
        private void LoadDepInfo(ViaticosConfig cfg, System.Data.IDbConnection db, string subunidad)
        {
            if (string.IsNullOrWhiteSpace(cfg.Cdp))
            {
                LoggerHelper.Warning($"[ViaticosRepository] config.cdp vacío o nulo para la sede '{subunidad}'; no se cargarán dependencia, rubro ni responsable del CDP.");
                return;
            }
            try
            {
                var cdpDep = Dapper.SqlMapper.QueryFirstOrDefault<dynamic>(db,
                    $@"{SqlCdpPorCdp}
                        AND (@Subunidad IS NULL OR @Subunidad = ''
                             OR subunidad = @Subunidad OR IFNULL(subunidad, '') = '')
                      {SqlCdpOrden}",
                    new { Cdp = cfg.Cdp, Subunidad = subunidad });

                if (cdpDep == null && !string.IsNullOrWhiteSpace(subunidad))
                {
                    var filaGlobal = Dapper.SqlMapper.QueryFirstOrDefault<dynamic>(db,
                        $@"{SqlCdpPorCdp}
                          {SqlCdpOrden}",
                        new { Cdp = cfg.Cdp });

                    string subFila = filaGlobal?.subunidad?.ToString();
                    if (filaGlobal != null && (string.IsNullOrWhiteSpace(subFila)
                        || string.Equals(subFila.Trim(), subunidad.Trim(), StringComparison.OrdinalIgnoreCase)))
                    {
                        cdpDep = filaGlobal;
                        LoggerHelper.Warning(
                            $"[ViaticosRepository] CDP '{cfg.Cdp}' sin registro para la sede '{subunidad}'; se usó fila global.");
                    }
                    else if (filaGlobal != null)
                    {
                        LoggerHelper.Warning(
                            $"[ViaticosRepository] CDP '{cfg.Cdp}' solo existe en la sede '{subFila}'. No se usa para la sede '{subunidad}'.");
                        return;
                    }
                }

                if (cdpDep == null)
                {
                    LoggerHelper.Warning($"[ViaticosRepository] CDP '{cfg.Cdp}' sin registro en cdpDependencias.");
                    return;
                }

                string subFilaBase = cdpDep.subunidad?.ToString();
                if (string.IsNullOrWhiteSpace(subunidad) && !string.IsNullOrWhiteSpace(subFilaBase))
                {
                    LoggerHelper.Warning(
                        $"[ViaticosRepository] Sede activa indeterminada; se usó la fila de la sede '{subFilaBase}' para el CDP '{cfg.Cdp}'.");
                }

                string dep = cdpDep.dependencia?.ToString()?.Trim() ?? "";
                string depFull = cdpDep.dependencia_full?.ToString()?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(dep))
                {
                    LoggerHelper.Warning($"[ViaticosRepository] CDP '{cfg.Cdp}' sin 'dependencia' en cdpDependencias; el párrafo presupuestal quedará incompleto.");
                }

                cfg.CortaDependencia = dep;
                cfg.FullDependencia = !string.IsNullOrWhiteSpace(depFull)
                    ? depFull
                    : ResolverFullDependencia(db, dep, subunidad);
                cfg.Fcdp = cdpDep.fechacdp?.ToString()?.Trim() ?? "";
                cfg.Rubro = cdpDep.rubro?.ToString()?.Trim() ?? "";
                cfg.ResponsableCDP = cdpDep.responsable?.ToString()?.Trim() ?? "";
                cfg.CargoResponsableCDP = cdpDep.cargo?.ToString()?.Trim() ?? "";
                cfg.Elaboro = cdpDep.elaboro?.ToString()?.Trim() ?? "";
                cfg.Reviso = cdpDep.reviso?.ToString()?.Trim() ?? "";
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[ViaticosRepository] Error al cargar LoadDepInfo", ex);
            }
        }

        /// <summary>
        /// Resuelve el nombre completo de la dependencia (FullDependencia) exclusivamente desde SQLite.
        /// Cadena: cdpDependencias.dependencia_full (sede/período vigente) -> nombre corto.
        /// No consulta la hoja Dependencias (tblDependencias) ni ninguna otra fuente de Excel.
        /// </summary>
        /// <param name="db">Conexión activa a SQLite.</param>
        /// <param name="corta">Nombre corto de la dependencia.</param>
        /// <param name="subunidad">Subunidad o sede activa.</param>
        /// <returns>dependencia_full registrada o el nombre corto si no existe.</returns>
        private string ResolverFullDependencia(System.Data.IDbConnection db, string corta, string subunidad)
        {
            if (string.IsNullOrWhiteSpace(corta)) return "";

            string cacheKey = ((subunidad ?? "").Trim().ToUpperInvariant()) + "|" + ((corta ?? "").Trim().ToUpperInvariant());
            lock (_cacheFullDepLock)
            {
                if (_cacheFullDep.TryGetValue(cacheKey, out var cached))
                    return string.IsNullOrWhiteSpace(cached) ? corta : cached;
            }

            string result = corta;
            bool resuelto = false;

            // Única fuente: SQLite (cdpDependencias.dependencia_full)
            try
            {
                if (db != null)
                {
                    string fullFromDb = Dapper.SqlMapper.QueryFirstOrDefault<string>(db,
                        @"SELECT dependencia_full FROM cdpDependencias
                          WHERE UPPER(TRIM(dependencia)) = UPPER(TRIM(@Corta))
                            AND dependencia_full IS NOT NULL AND TRIM(dependencia_full) <> ''
                            AND (@Subunidad IS NULL OR @Subunidad = '' OR subunidad = @Subunidad OR IFNULL(subunidad,'') = '')
                          ORDER BY IFNULL(activo, 0) DESC,
                                   CASE WHEN idperiodo = (SELECT id_periodo FROM config WHERE id = 1) THEN 1 ELSE 0 END DESC,
                                   id DESC
                          LIMIT 1",
                        new { Corta = corta, Subunidad = subunidad });

                    if (!string.IsNullOrWhiteSpace(fullFromDb))
                    {
                        result = fullFromDb.Trim();
                        resuelto = true;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning("[ViaticosRepository] Error al consultar dependencia_full en SQLite: " + ex.Message);
            }

            lock (_cacheFullDepLock)
            {
                _cacheFullDep[cacheKey] = resuelto ? result : null;
            }

            if (!resuelto)
            {
                LoggerHelper.Warning($"[ViaticosRepository] Dependencia '{corta}' sin 'dependencia_full' en cdpDependencias (sede '{subunidad}'); se usará el nombre corto. Registre el valor en 'Gestión de CDP Dependencias' (FormCdp).");
            }

            return result;
        }

        public Empleado GetEmpleado(string cc)
        {
            if (string.IsNullOrEmpty(cc)) return null;
            if (_cacheEmp.TryGetValue(cc, out var e)) return e;

            var wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return null;

            if (!_empIndexed) Indexar(wb);
            _cacheEmp.TryGetValue(cc, out e);
            return e;
        }

        private void Indexar(dynamic wb)
        {
            try
            {
                dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_DATOS];
                int last = (int)ws.Cells[ws.Rows.Count, "B"].End[-4162].Row;
                object[,] data = (object[,])ws.Range[ws.Cells[1, 1], ws.Cells[last, 14]].Value2;
                int total = data.GetLength(0);
                for (int i = 1; i <= total; i++)
                {
                    string cc = data[i, ConstantsHelper.DATOS_CEDULA]?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(cc) || _cacheEmp.ContainsKey(cc)) continue;

                    _cacheEmp[cc] = new Empleado
                    {
                        Cedula = cc, 
                        NombreCompleto = data[i, ConstantsHelper.DATOS_NOMBRE]?.ToString() ?? "",
                        Salario = data[i, ConstantsHelper.DATOS_SALARIO], 
                        Rol = data[i, ConstantsHelper.DATOS_ROL]?.ToString() ?? "",
                        Categoria = data[i, ConstantsHelper.DATOS_CATEGORIA]?.ToString() ?? "", 
                        EntidadBancaria = data[i, ConstantsHelper.DATOS_ENTIDAD]?.ToString() ?? "",
                        NumeroCuenta = data[i, ConstantsHelper.DATOS_CUENTA_NUM]?.ToString() ?? "", 
                        TipoCuenta = data[i, ConstantsHelper.DATOS_CUENTA_TIPO]?.ToString() ?? "",
                        Telefono = data[i, ConstantsHelper.DATOS_TEL]?.ToString() ?? "", 
                        Correo = data[i, ConstantsHelper.DATOS_CORREO]?.ToString() ?? "",
                        FechaNacimiento = data[i, ConstantsHelper.DATOS_NACIMIENTO], 
                        Contrato = data[i, ConstantsHelper.DATOS_CONTRATO]?.ToString() ?? "", 
                        Dependencia = data[i, ConstantsHelper.DATOS_DEPENDENCIA]?.ToString() ?? ""
                    };
                }
                _empIndexed = true;
            }
            catch (Exception ex)
            {
                _empIndexed = true;
                LoggerHelper.Warning($"[ViaticosRepository] Error al indexar empleados de la hoja DATOS: {ex.Message}");
            }
        }

        /// <summary>
        /// Construye y mapea una entidad <see cref="Viaje"/> a partir de una fila del rango Excel de Programación.
        /// Fase 246: Parseo defensivo estricto, distinción de celda vacía vs 0 explícito y resolución de topes
        /// de terminales aéreas desde el catálogo en memoria con trazabilidad completa.
        /// </summary>
        /// <param name="data">Matriz bidimensional COM con los datos leídos de la hoja Excel.</param>
        /// <param name="r">Índice de la fila a procesar (1-based).</param>
        /// <returns>Instancia de <see cref="Viaje"/> totalmente configurada.</returns>
        public Viaje GetViajeFromRow(object[,] data, int r)
        {
            object rawTransporte = data[r, ConstantsHelper.PROG_TRANSPORTE];
            object rawAereo = data[r, ConstantsHelper.PROG_TRANSPORTEAEREO];

            ViaticosHelper.TryParseDecimal(rawAereo, out decimal aereoVal, out bool aereoBlank);
            ViaticosHelper.TryParseDecimal(rawTransporte, out decimal transpVal, out _);

            var v = new Viaje
            {
                cedula = data[r, ConstantsHelper.PROG_CEDULA]?.ToString() ?? "",
                FechaInicio = ViaticosHelper.ToDateTime(data[r, ConstantsHelper.PROG_FECHAVIAJE]),
                FechaFin = ViaticosHelper.ToDateTime(data[r, ConstantsHelper.PROG_FECHAFIN]),
                Duracion = int.TryParse(data[r, ConstantsHelper.PROG_DURACION]?.ToString(), out int dur) ? dur : 1,
                terminal = (data[r, ConstantsHelper.PROG_TERMINAL]?.ToString() ?? "TERRESTRE").Trim(),
                Transporte = transpVal,
                TransporteAereo = aereoVal,
                Dependencia = (data[r, ConstantsHelper.PROG_DEPENDENCIA]?.ToString() ?? "PREGRADO").Trim()
            };

            // Fase 246: Detección centralizada de transporte aéreo
            bool esAereo = ViaticosHelper.EsTerminalAerea(v.terminal);

            if (esAereo)
            {
                v.IsTerminalAereo = 1;
                v.TerminalAereaNombre = v.terminal;

                // Si la celda de transporte aéreo en Excel estaba vacía o no provista, autocompletar con el catálogo canónico
                if (aereoBlank)
                {
                    if (ViaticosHelper.ResolverTopeTerminal(0, v.terminal, "ViaticosRepository.GetViajeFromRow",
                        out int idRes, out string nomRes, out decimal maxVal))
                    {
                        v.IdTerminales = idRes;
                        v.TransporteAereo = maxVal;
                        v.GastoTerminalAerea = maxVal;
                    }
                }
                else
                {
                    // Si el usuario digitó explícitamente un valor numérico (incluido 0), se respeta
                    v.GastoTerminalAerea = v.TransporteAereo;
                    if (ViaticosHelper.ResolverTopeTerminal(0, v.terminal, "ViaticosRepository.GetViajeFromRow",
                        out int idRes, out _, out _))
                    {
                        v.IdTerminales = idRes;
                    }
                }
            }

            string iti = (data[r, ConstantsHelper.PROG_ITINERARIO]?.ToString() ?? "").Trim().ToUpper();
            v.Itinerario = iti;
            var pts = iti.Split('-');
            if (pts.Length >= 2)
            {
                v.CiudadOrigen = NormalizarCiudad(pts[0].Trim());
                v.CiudadDestino = NormalizarCiudad(pts[1].Trim());
            }

            string objBase = data[r, ConstantsHelper.PROG_OBJETO]?.ToString() ?? "************ ADICIONAR OBJETO DEL VIATICO ************";
            string fechasObjeto = v.Duracion > 1
                ? $" DEL {v.FechaInicio:dd/MM/yyyy} AL {v.FechaFin:dd/MM/yyyy}"
                : $" EL {v.FechaInicio:dd/MM/yyyy}";
            v.Objeto = ViaticosHelper.LimpiarTexto($"PAGO VIATICOS PARA {objBase} EN CETAP {v.CiudadDestino}{fechasObjeto}");
            v.Asignatura = $"{v.FechaInicio:dd/MM/yyyy} - {objBase} - {data[r, ConstantsHelper.PROG_NOMBRE]?.ToString() ?? ""}";
            return v;
        }

        /// <summary>
        /// Normaliza el nombre de una ciudad según la tabla 'tblDestinos' de la hoja Dependencias.
        /// Fase 13.1: Guard defensivo (wb == null).
        /// Fase 15.1: Bloque try-catch para proteger llamadas COM a la hoja de dependencias y lista de destinos.
        /// </summary>
        /// <param name="ciudad">Nombre original de la ciudad.</param>
        /// <returns>Nombre normalizado de la ciudad o el nombre original si no se encuentra o no hay libro activo.</returns>
        public string NormalizarCiudad(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return ciudad;
            if (_cacheCiu.TryGetValue(ciudad, out var n)) return n;

            var wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return ciudad;

            try
            {
                dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_DEPENDENCIAS];
                dynamic tbl = ws.ListObjects["tblDestinos"];
                object[,] data = (object[,])tbl.DataBodyRange.Value2;
                int total = data.GetLength(0);
                for (int i = 1; i <= total; i++)
                {
                    if (string.Equals(data[i, ConstantsHelper.DEST_CIUDAD]?.ToString()?.Trim(), ciudad, StringComparison.OrdinalIgnoreCase))
                    {
                        string res = data[i, ConstantsHelper.DEST_NORMALIZADA]?.ToString()?.Trim() ?? ciudad;
                        _cacheCiu[ciudad] = res; return res;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[ViaticosRepository] Error al normalizar ciudad '{ciudad}': {ex.Message}");
            }

            _cacheCiu[ciudad] = ciudad; return ciudad;
        }

        /// <summary>
        /// Obtiene el departamento correspondiente a una ciudad según 'tblDestinos' o la configuración base.
        /// Fase 14.2: Bloque try-catch para proteger las lecturas COM de la hoja de dependencias y de la hoja de configuración.
        /// </summary>
        /// <param name="ciudad">Nombre de la ciudad.</param>
        /// <returns>Nombre del departamento o el valor por defecto configurado.</returns>
        public string ObtenerDepartamento(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return "";
            if (_cacheDep.TryGetValue(ciudad, out var d)) return d;

            var wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return "";

            try
            {
                dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_DEPENDENCIAS];
                dynamic tbl = ws.ListObjects["tblDestinos"];
                object[,] data = (object[,])tbl.DataBodyRange.Value2;
                int total = data.GetLength(0);
                for (int i = 1; i <= total; i++)
                {
                    if (string.Equals(data[i, ConstantsHelper.DEST_NORMALIZADA]?.ToString()?.Trim(), ciudad, StringComparison.OrdinalIgnoreCase))
                    {
                        string res = data[i, ConstantsHelper.DEST_DEP]?.ToString()?.Trim() ?? "";
                        _cacheDep[ciudad] = res; return res;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[ViaticosRepository] Error al buscar departamento en tblDestinos para '{ciudad}': {ex.Message}");
            }

            // Fallback: CONFIG_DEPTO_BASE del Config
            try
            {
                dynamic wsCfg = wb.Worksheets[ConstantsHelper.SHEET_CONFIG];
                string def = wsCfg.Cells[ConstantsHelper.CONFIG_DEPTO_BASE, ConstantsHelper.CONFIG_COL_VALOR].Value?.ToString()?.Trim() ?? "";
                _cacheDep[ciudad] = def;
                return def;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[ViaticosRepository] Error al leer departamento base de la hoja Config: {ex.Message}");
                _cacheDep[ciudad] = "";
                return "";
            }
        }

        public bool ValidarFormatoComisionado(dynamic sheet)
        {
            if (sheet == null) return false;
            try
            {
                object val = sheet.Range["A1"].Value;
                if (val == null) return false;
                string s = val.ToString().Trim().ToUpperInvariant();
                return s.StartsWith("IDENTIFICACIÓN") || s.StartsWith("IDENTIFICACIN") || s.StartsWith("IDENTIFICACION");
            }
            catch { return false; }
        }

        /// <summary>
        /// Obtiene la lista de números de CDP registrados exclusivamente desde SQLite.
        /// Fase 25.2/26.2/27.3: Desacoplamiento total de Excel COM con fallback escalonado por subunidad (específica -> global vacía -> todas).
        /// </summary>
        public List<string> GetCdpList()
        {
            try
            {
                var db = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext().Open();
                string subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;

                var lista = Dapper.SqlMapper.Query<string>(db,
                    @"SELECT DISTINCT cdpnum FROM cdpDependencias
                      WHERE cdpnum IS NOT NULL AND TRIM(cdpnum) <> ''
                        AND (@Subunidad IS NULL OR @Subunidad = '' OR subunidad = @Subunidad OR IFNULL(subunidad,'') = '')
                      ORDER BY cdpnum;",
                    new { Subunidad = subunidad }).AsList();

                if (lista.Count == 0)
                {
                    lista = Dapper.SqlMapper.Query<string>(db,
                        @"SELECT DISTINCT cdpnum FROM cdpDependencias
                          WHERE cdpnum IS NOT NULL AND TRIM(cdpnum) <> ''
                            AND IFNULL(subunidad, '') = ''
                          ORDER BY cdpnum;").AsList();

                    if (lista.Count == 0)
                    {
                        lista = Dapper.SqlMapper.Query<string>(db,
                            @"SELECT DISTINCT cdpnum FROM cdpDependencias
                              WHERE cdpnum IS NOT NULL AND TRIM(cdpnum) <> ''
                              ORDER BY cdpnum;").AsList();
                        if (lista.Count > 0)
                            LoggerHelper.Warning($"[ViaticosRepository] Sin CDPs para la sede '{subunidad}'; se usó la lista global de cdpDependencias.");
                    }
                }

                return lista;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning("[ViaticosRepository] Error al obtener lista de CDPs desde SQLite: " + ex.Message);
                return new List<string>();
            }
        }

        /// <summary>
        /// Obtiene el CDP asociado a una dependencia corta exclusivamente desde SQLite.
        /// Cadena: cdpDependencias.cdpnum (sede/período vigente) -> vacío.
        /// Fase 25.2: Desacoplamiento total de Excel COM.
        /// </summary>
        /// <param name="dep">Nombre corto de la dependencia.</param>
        /// <returns>Número de CDP o cadena vacía si no existe.</returns>
        public string GetDependenciaCDP(string dep)
        {
            if (string.IsNullOrWhiteSpace(dep)) return "";

            try
            {
                var db = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext().Open();
                string subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;

                string cdp = Dapper.SqlMapper.QueryFirstOrDefault<string>(db,
                    @"SELECT cdpnum FROM cdpDependencias
                      WHERE UPPER(TRIM(dependencia)) = UPPER(TRIM(@Dep))
                        AND cdpnum IS NOT NULL AND TRIM(cdpnum) <> ''
                        AND (@Subunidad IS NULL OR @Subunidad = '' OR subunidad = @Subunidad OR IFNULL(subunidad,'') = '')
                      ORDER BY IFNULL(activo, 0) DESC,
                               CASE WHEN idperiodo = (SELECT id_periodo FROM config WHERE id = 1) THEN 1 ELSE 0 END DESC,
                               id DESC
                      LIMIT 1",
                    new { Dep = dep.Trim(), Subunidad = subunidad });

                if (!string.IsNullOrWhiteSpace(cdp)) return cdp.Trim();
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[ViaticosRepository] Error al resolver CDP por dependencia '{dep}' en SQLite: {ex.Message}");
            }

            LoggerHelper.Warning($"[ViaticosRepository] Dependencia '{dep}' sin CDP en cdpDependencias (sede '{ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad}').");
            return "";
        }

        public string GetGlobalVariable(string varName) => _excelRepository.GetGlobalVariable(varName);

        private void ValidarEstructura(dynamic wb)
        {
            if (!Exists(wb, ConstantsHelper.SHEET_DATOS))
                throw new ViaticosFormatException($"No se encontró la hoja necesaria: '{ConstantsHelper.SHEET_DATOS}'");
        }

        private bool Exists(dynamic wb, string name)
        {
            foreach (dynamic sh in wb.Worksheets)
            {
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }
    }
}