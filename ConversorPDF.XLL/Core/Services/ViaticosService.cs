using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
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

        /// <summary> constructor con inyección de dependencias. </summary>
        public ViaticosService(
            IExcelRepository excelRepository,
            IUserService userService,
            IViaticosRepository viaticosRepository,
            IViaticosIndividualProcessor individualProcessor,
            IViaticosMasivoProcessor masivoProcessor,
            IViaticosCumplidosProcessor cumplidosProcessor,
            IViaticosGestionProcessor gestionProcessor)
        {
            _excelRepository = excelRepository;
            _userService = userService;
            _viaticosRepository = viaticosRepository;
            _individualProcessor = individualProcessor;
            _masivoProcessor = masivoProcessor;
            _cumplidosProcessor = cumplidosProcessor;
            _gestionProcessor = gestionProcessor;
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
        public void RenombrarSolicitudes(int inicio, int fin)
            => _gestionProcessor.RenombrarSolicitudes(_excelRepository.GetOneDriveFolderPath());

        /// <inheritdoc/>
        public int UnificarSolicitudes(int inicio, int fin)
            => _gestionProcessor.UnificarSolicitudes(_excelRepository.GetOneDriveFolderPath());

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
    }
}