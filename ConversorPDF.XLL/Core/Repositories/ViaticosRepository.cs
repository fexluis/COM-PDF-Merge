using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
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
        private bool _empIndexed;

        public ViaticosRepository(IExcelRepository excelRepository)
        {
            _excelRepository = excelRepository;
        }

        public ViaticosConfig GetConfig(DateTime? start = null, DateTime? end = null, string cdpOverride = null)
        {
            var wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return null;

            ValidarEstructura(wb);

            var config = new ViaticosConfig();
            dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_CONFIG];
            object[,] data = (object[,])ws.Range["A1:B" + ConstantsHelper.CONFIG_ROW_MAX].Value2;
            int v = ConstantsHelper.CONFIG_COL_VALOR;

            config.FechaInicio      = start       ?? ViaticosHelper.ToDateTime(data[ConstantsHelper.CONFIG_FECHA_INICIO,      v]);
            config.FechaFin         = end         ?? ViaticosHelper.ToDateTime(data[ConstantsHelper.CONFIG_FECHA_FIN,         v]);
            config.Cdp              = cdpOverride ?? data[ConstantsHelper.CONFIG_CDP,              v]?.ToString() ?? "";
            config.SalarioDecreto   =                data[ConstantsHelper.CONFIG_SALARIO_DECRETO,  v];
            config.ViaticoA         =                data[ConstantsHelper.CONFIG_VIATICO_A,        v];
            config.ViaticoB         =                data[ConstantsHelper.CONFIG_VIATICO_B,        v];
            config.PorcA            =                data[ConstantsHelper.CONFIG_PORC_A,           v];
            config.PorcB            =                data[ConstantsHelper.CONFIG_PORC_B,           v];
            config.NombreDT         =                data[ConstantsHelper.CONFIG_NOMBRE_DT,        v]?.ToString() ?? "";
            config.CargoDT          =                data[ConstantsHelper.CONFIG_CARGO_DT,         v]?.ToString() ?? "";
            // Alias: por defecto, el firmante es el DT
            config.Gerente          = config.NombreDT;
            config.Cargo            = config.CargoDT;
            config.CedulaGerente    =                data[ConstantsHelper.CONFIG_CEDULA_GERENTE,   v]?.ToString() ?? "";
            config.GerenteSustituto =                data[ConstantsHelper.CONFIG_GERENTE_SUSTITUTO,v]?.ToString() ?? "";
            config.CargoSustituto   =                data[ConstantsHelper.CONFIG_CARGO_SUSTITUTO,  v]?.ToString() ?? "";
            config.DeptoBase        =                data[ConstantsHelper.CONFIG_DEPTO_BASE,       v]?.ToString() ?? "";
            config.CiudadBase       =                data[ConstantsHelper.CONFIG_CIUDAD_BASE,      v]?.ToString() ?? "";
            config.AeropuertoBase   =                data[ConstantsHelper.CONFIG_AEROPUERTO_BASE,  v]?.ToString() ?? "";

            config.CarpetaDestino = ViaticosHelper.CrearCarpetaDestino();
            LoadDepInfo(wb, config);
            return config;
        }

        private void LoadDepInfo(dynamic wb, ViaticosConfig cfg)
        {
            dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_DEPENDENCIAS];
            dynamic tbl = null;
            try { tbl = ws.ListObjects["tblDependencias"]; } catch { }

            if (tbl == null || tbl.DataBodyRange == null)
                throw new ViaticosFormatException($"No se encontró la tabla 'tblDependencias' en la hoja '{ConstantsHelper.SHEET_DEPENDENCIAS}' o está vacía.");

            object[,] data = (object[,])tbl.DataBodyRange.Value2;
            int rows = data.GetLength(0);
            for (int i = 1; i <= rows; i++)
            {
                if (data[i, ConstantsHelper.DEP_CDP]?.ToString() == cfg.Cdp)
                {
                    cfg.CortaDependencia = data[i, ConstantsHelper.DEP_CORTA]?.ToString() ?? "";
                    cfg.FullDependencia = data[i, ConstantsHelper.DEP_FULL]?.ToString() ?? "";
                    cfg.Fcdp = data[i, ConstantsHelper.DEP_FCDP]?.ToString() ?? "";
                    cfg.Rubro = data[i, ConstantsHelper.DEP_RUBRO]?.ToString() ?? "";
                    cfg.Elaboro = data[i, ConstantsHelper.DEP_ELABORO]?.ToString() ?? "";
                    cfg.Reviso = data[i, ConstantsHelper.DEP_REVISO]?.ToString() ?? "";
                    break;
                }
            }
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
            catch { }
        }

        public Viaje GetViajeFromRow(object[,] data, int r)
        {
            var v = new Viaje
            {
                cedula = data[r, ConstantsHelper.PROG_CEDULA]?.ToString() ?? "",
                FechaInicio = ViaticosHelper.ToDateTime(data[r, ConstantsHelper.PROG_FECHAVIAJE]),
                FechaFin = ViaticosHelper.ToDateTime(data[r, ConstantsHelper.PROG_FECHAFIN]),
                Duracion = int.TryParse(data[r, ConstantsHelper.PROG_DURACION]?.ToString(), out int dur) ? dur : 1,
                terminal = (data[r, ConstantsHelper.PROG_TERMINAL]?.ToString() ?? "TERRESTRE").Trim(),
                Transporte = decimal.TryParse(data[r, ConstantsHelper.PROG_TRANSPORTE]?.ToString(), out decimal t) ? t : 0,
                TransporteAereo = decimal.TryParse(data[r, ConstantsHelper.PROG_TRANSPORTEAEREO]?.ToString(), out decimal ta) ? ta : 0,
                Dependencia = (data[r, ConstantsHelper.PROG_DEPENDENCIA]?.ToString() ?? "PREGRADO").Trim()
            };

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

        public string NormalizarCiudad(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return ciudad;
            if (_cacheCiu.TryGetValue(ciudad, out var n)) return n;

            var wb = _excelRepository.GetActiveWorkbook();
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
            _cacheCiu[ciudad] = ciudad; return ciudad;
        }

        public string ObtenerDepartamento(string ciudad)
        {
            if (string.IsNullOrWhiteSpace(ciudad)) return "";
            if (_cacheDep.TryGetValue(ciudad, out var d)) return d;

            var wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return "";

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

            // Fallback: CONFIG_DEPTO_BASE del Config
            dynamic wsCfg = wb.Worksheets[ConstantsHelper.SHEET_CONFIG];
            string def = wsCfg.Cells[ConstantsHelper.CONFIG_DEPTO_BASE, ConstantsHelper.CONFIG_COL_VALOR].Value?.ToString() ?? "";
            _cacheDep[ciudad] = def;
            return def;
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

        public List<string> GetCdpList() => _excelRepository.GetTableColumnValues("tblDependencias", "CDP");

        public string GetDependenciaCDP(string dep)
        {
            if (string.IsNullOrEmpty(dep)) return "";
            var wb = _excelRepository.GetActiveWorkbook();
            dynamic ws = wb.Worksheets[ConstantsHelper.SHEET_DEPENDENCIAS];
            dynamic tbl = ws.ListObjects["tblDependencias"];
            object[,] data = (object[,])tbl.DataBodyRange.Value2;
            int rows = data.GetLength(0);
            for (int i = 1; i <= rows; i++)
            {
                if (string.Equals(data[i, ConstantsHelper.DEP_CORTA]?.ToString(), dep, StringComparison.OrdinalIgnoreCase))
                {
                    return data[i, ConstantsHelper.DEP_CDP]?.ToString() ?? "";
                }
            }
            return "";
        }

        public string GetGlobalVariable(string varName) => _excelRepository.GetGlobalVariable(varName);

        private void ValidarEstructura(dynamic wb)
        {
            if (!Exists(wb, ConstantsHelper.SHEET_CONFIG))
                throw new ViaticosFormatException($"No se encontró la hoja necesaria: '{ConstantsHelper.SHEET_CONFIG}'");

            if (!Exists(wb, ConstantsHelper.SHEET_DEPENDENCIAS))
                throw new ViaticosFormatException($"No se encontró la hoja necesaria: '{ConstantsHelper.SHEET_DEPENDENCIAS}'");

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