using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.IO;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Repositories;

namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio fachada que coordina todas las operaciones relacionadas con viáticos.
    /// Inyecta y utiliza repositorios y procesadores especializados para cada tarea.
    /// </summary>
    public class ViaticosService : IViaticosService
    {
        private readonly IExcelRepository _excelRepository;
        private readonly IUserService _userService;
        private readonly IViaticosRepository _viaticosRepository;
        private readonly IViaticosIndividualProcessor _individualProcessor;
        private readonly IViaticosMasivoProcessor _masivoProcessor;
        private readonly IViaticosCumplidosProcessor _cumplidosProcessor;
        private readonly IViaticosGestionProcessor _gestionProcessor;
        private readonly IBusquedaInformesProcessor _busquedaInformesProcessor;

        /// <summary> constructor con inyección de dependencias. </summary>
        public ViaticosService(
            IExcelRepository excelRepository,
            IUserService userService,
            IViaticosRepository viaticosRepository,
            IViaticosIndividualProcessor individualProcessor,
            IViaticosMasivoProcessor masivoProcessor,
            IViaticosCumplidosProcessor cumplidosProcessor,
            IViaticosGestionProcessor gestionProcessor,
            IBusquedaInformesProcessor busquedaInformesProcessor)
        {
            _excelRepository = excelRepository;
            _userService = userService;
            _viaticosRepository = viaticosRepository;
            _individualProcessor = individualProcessor;
            _masivoProcessor = masivoProcessor;
            _cumplidosProcessor = cumplidosProcessor;
            _gestionProcessor = gestionProcessor;
            _busquedaInformesProcessor = busquedaInformesProcessor;
        }

        /// <inheritdoc/>
        public void ClearCache() { }

        /// <inheritdoc/>
        public ViaticosConfig GetConfig(DateTime? start = null, DateTime? end = null, string cdp = null) 
            => _viaticosRepository.GetConfig(start, end, cdp);

        /// <inheritdoc/>
        public Empleado GetEmpleado(string cc) => _viaticosRepository.GetEmpleado(cc);

        /// <inheritdoc/>
        public Viaje GetViajeFromRow(object[,] data, int r) => _viaticosRepository.GetViajeFromRow(data, r);

        /// <inheritdoc/>
        public int ProcesarIndividual(Action<int, int> progressCallback = null, DateTime? inicio = null, DateTime? fin = null, string cdp = null)
            => _individualProcessor.Procesar(progressCallback, inicio, fin, cdp);

        /// <inheritdoc/>
        public int ProcesarMasivo(Action<int, int> progressCallback = null)
            => _masivoProcessor.Procesar(progressCallback);

        /// <inheritdoc/>
        public int ProcesarCumplidos(Action<int, int> progressCallback = null, Action onComplete = null)
            => _cumplidosProcessor.Procesar(progressCallback);

        /// <summary>
        /// Determina el tipo de procesamiento basado en la hoja activa de Excel y ejecuta el proceso para la fila seleccionada.
        /// Identifica si el usuario está en la hoja de CONSULTAS (cumplidos) o Programación (individual).
        /// </summary>
        public int ProcesarFilaSeleccionada(Action<int, int> progressCallback = null)
        {
            dynamic wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return 0;
            
            dynamic app = wb.Application;
            dynamic activeSheet = app.ActiveSheet;
            string sheetName = activeSheet.Name;
            int row = app.ActiveCell.Row;

            if (sheetName.Equals(ConstantsHelper.SHEET_CONSULTAS, StringComparison.OrdinalIgnoreCase))
            {
                object fecha = activeSheet.Cells[row, ConstantsHelper.CONSULTAS_FECHA_INICIO].Value;
                object cdp = activeSheet.Cells[row, ConstantsHelper.CONSULTAS_CDP].Value;
                
                FilaSeleccionadaSaveConfig(fecha, fecha, cdp?.ToString());
                
                return _cumplidosProcessor.ProcesarFilaSeleccionada(progressCallback);
            }
            else
            {
                object fecha = activeSheet.Cells[row, ConstantsHelper.PROG_FECHAVIAJE].Value;
                string dep = activeSheet.Cells[row, ConstantsHelper.PROG_DEPENDENCIA].Value?.ToString();
                string cdp = _viaticosRepository.GetDependenciaCDP(dep);
                
                FilaSeleccionadaSaveConfig(fecha, fecha, string.IsNullOrWhiteSpace(cdp) ? null : cdp);
                
                return _individualProcessor.ProcesarFilaSeleccionada(progressCallback);
            }
        }

        /// <inheritdoc/>
        public int RenombrarSolicitudes(int inicio, int fin)
            => _gestionProcessor.RenombrarSolicitudes(_excelRepository.GetOneDriveFolderPath(), inicio, fin);

        /// <inheritdoc/>
        public ResultadoRenombrado RenombrarSolicitudesConDetalle(int inicio, int fin)
            => _gestionProcessor.RenombrarSolicitudesConDetalle(_excelRepository.GetOneDriveFolderPath(), inicio, fin);

        /// <inheritdoc/>
        public RangoSolicitudes ObtenerRangoSolicitudes()
            => _gestionProcessor.ObtenerRangoSolicitudes();

        /// <inheritdoc/>
        public ResultadoVerificacionSolicitudes VerificarSolicitudesPresentes(int inicio, int fin)
            => _gestionProcessor.VerificarSolicitudesPresentes(_excelRepository.GetOneDriveFolderPath(), inicio, fin);

        /// <inheritdoc/>
        public int UnificarSolicitudes(int inicio, int fin)
            => _gestionProcessor.UnificarSolicitudes(_excelRepository.GetOneDriveFolderPath(), inicio, fin);

        /// <inheritdoc/>
        public bool ValidarFormatoComisionado(dynamic sheet) => _viaticosRepository.ValidarFormatoComisionado(sheet);

        /// <inheritdoc/>
        public List<string> GetCdpList() => _viaticosRepository.GetCdpList();

        /// <inheritdoc/>
        public string[] GetSelectionConfig()
        {
            var cfg = _viaticosRepository.GetConfig();
            return new string[] { cfg.FechaInicio.ToString("dd/MM/yyyy"), cfg.FechaFin.ToString("dd/MM/yyyy"), cfg.Cdp };
        }

        /// <summary>
        /// Guarda temporalmente la configuración de fechas y CDP en la hoja 'Config' para que los procesadores la utilicen.
        /// </summary>
        public void FilaSeleccionadaSaveConfig(object inicio, object fin, string cdp = null)
        {
            dynamic ws = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_CONFIG);
            if (ws == null) return;
            if (inicio != null) _excelRepository.SetCellValue(ws, "B1", inicio);
            if (fin != null) _excelRepository.SetCellValue(ws, "B2", fin);
            if (!string.IsNullOrWhiteSpace(cdp)) _excelRepository.SetCellValue(ws, "B3", cdp);
        }

        /// <inheritdoc/>
        public string GetGlobalVariable(string varName) => _viaticosRepository.GetGlobalVariable(varName);

        /// <inheritdoc/>
        public object[,] GetAdvancedConfigRows()
        {
            dynamic ws = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_CONFIG);
            return _excelRepository.GetRangeValues(ws, 4, 2, 16, 2);
        }

        /// <inheritdoc/>
        public void SaveAdvancedConfigRows(object[,] data)
        {
            dynamic ws = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_CONFIG);
            for (int i = 0; i < data.GetLength(0); i++)
                _excelRepository.SetCellValue(ws, "B" + (4 + i), data[i, 0]);
        }

        /// <inheritdoc/>
        public object[,] GetTableData(string tableName) => _excelRepository.GetTableData(tableName);

        /// <inheritdoc/>
        public void SaveTableData(string tableName, object[,] data) => _excelRepository.SetTableData(tableName, data);

        /// <inheritdoc/>
        public string ObtenerCarpetaInformes()
        {
            string carpetaLibro = _excelRepository.GetOneDriveFolderPath();
            if (string.IsNullOrWhiteSpace(carpetaLibro)) return string.Empty;

            return Path.Combine(carpetaLibro, App.FolderNames.INFORMES);
        }

        /// <inheritdoc/>
        public ResultadoBusquedaInformes ProcesarInformes(int inicio, int fin, bool unificarParejas,
                                                           bool auditarConsultas, bool simular,
                                                           string carpetaRaiz = null,
                                                           Action<int, int> progress = null)
        {
            string carpeta = string.IsNullOrWhiteSpace(carpetaRaiz) ? ObtenerCarpetaInformes() : carpetaRaiz;

            // El progreso se normaliza a 0-100: 50 % para el renombrado y 50 % para la auditoría.
            Action<int, int> progresoRenombrado = null;
            Action<int, int> progresoAuditoria = null;

            if (progress != null)
            {
                double techoRenombrado = auditarConsultas ? 50.0 : 100.0;

                progresoRenombrado = (actual, total) =>
                {
                    double porcentaje = total <= 0 ? 0 : techoRenombrado * actual / total;
                    progress((int)porcentaje, 100);
                };

                progresoAuditoria = (actual, total) =>
                {
                    double porcentaje = total <= 0
                        ? techoRenombrado
                        : techoRenombrado + (100.0 - techoRenombrado) * actual / total;
                    progress((int)porcentaje, 100);
                };
            }

            var resultado = _busquedaInformesProcessor.RenombrarYUnificar(
                carpeta, inicio, fin, unificarParejas, simular, progresoRenombrado);

            if (!resultado.HayError && auditarConsultas)
            {
                var auditoria = _busquedaInformesProcessor.AuditarConsultas(
                    carpeta, inicio, fin, simular, progresoAuditoria);
                resultado.Absorber(auditoria);
            }

            return resultado;
        }
    }
}
