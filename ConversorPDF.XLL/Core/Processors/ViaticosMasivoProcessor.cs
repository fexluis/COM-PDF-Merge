using System;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.IO;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Repositories;
namespace ConversorPDF.XLL.Core.Processors
{
    /// <summary>
    /// Procesador encargado de la generación masiva de registros de viáticos en una única hoja de cálculo consolidada.
    /// Utiliza la plantilla 'Masivo' y aplica reglas de cálculo de tarifas automáticas.
    /// </summary>
    internal class ViaticosMasivoProcessor : IViaticosMasivoProcessor
    {
        private const int FILA_INICIAL = 4;
        private readonly IExcelRepository _excelRepository;
        private readonly IViaticosRepository _viaticosRepository;

        /// <summary>
        /// Inicializa una nueva instancia de la clase <see cref="ViaticosMasivoProcessor"/>.
        /// </summary>
        /// <param name="excelRepository">Acceso a Excel.</param>
        /// <param name="viaticosRepository">Lógica de negocio de viáticos.</param>
        public ViaticosMasivoProcessor(IExcelRepository excelRepository, IViaticosRepository viaticosRepository)
        {
            _excelRepository = excelRepository;
            _viaticosRepository = viaticosRepository;
        }

        /// <summary>
        /// Ejecuta el proceso de consolidación masiva para todos los registros que coincidan con la configuración activa.
        /// </summary>
        /// <param name="progressCallback">Opcional. Callback para notificar el progreso (actual, total).</param>
        /// <returns>Número de registros consolidados exitosamente.</returns>
        public int Procesar(Action<int, int> progressCallback = null)
        {
            dynamic wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return 0;

            var config = _viaticosRepository.GetConfig();
            if (config == null) return 0;
            if (string.IsNullOrWhiteSpace(config.Cdp))
            {
                MessageBox.Show("No hay CDP configurado. Regístrelo en 'Gestión de CDP Dependencias' antes de generar documentos.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            dynamic wsProg = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_PROG);
            int last = (int)wsProg.Cells[wsProg.Rows.Count, "A"].End[-4162 /*xlUp*/].Row;
            if (last < FILA_INICIAL)
            {
                LoggerHelper.Warning("No hay datos suficientes para el proceso masivo.");
                return 0;
            }

            LoggerHelper.Info($"Iniciando proceso masivo. Analizando {last - FILA_INICIAL + 1} filas.");

            string name = "Masivo " + DateTime.Now.ToString("yyyyMMdd");
            dynamic wsMas = CrearHojaMasiva(wb, name);
            var cat1 = GetCat1();

            int count = 0;
            int destRow = 1;

            using (var scope = new ExcelAppStateHelper(wb.Application))
            {
                scope.ApplyOptimized();
                object[,] data = _excelRepository.GetRangeValues(wsProg, FILA_INICIAL, 1, last, ConstantsHelper.PROG_COL_MAX);
                int total = data.GetLength(0);

                for (int i = 1; i <= total; i++)
                {
                    progressCallback?.Invoke(i, total);
                    if (ProcesarFila(wb, data, i, wsMas, config, FILA_INICIAL + i - 1, ref destRow, cat1))
                    {
                        count++;
                    }
                }

                if (count > 0)
                {
                    GuardarExcel(wb, wsMas, config.CarpetaDestino, name);
                    LoggerHelper.Success($"Proceso masivo completado. Consolidados {count} registros en {name}.xlsx");
                }
                else
                {
                    LoggerHelper.Warning("No se encontraron registros que cumplan los criterios para el masivo.");
                }
            }

            EliminarHoja(wb, name);

            try { wsProg.Select(); } catch { }

            return count;
        }

        /// <summary>
        /// Obtiene la lista de ciudades categorizadas como 'Categoría 1' para el cálculo de viáticos.
        /// </summary>
        /// <returns>Hashset de nombres de ciudades para búsqueda rápida.</returns>
        internal HashSet<string> GetCat1()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var data = _excelRepository.GetTableData("tblCategoria1");
            if (data != null)
            {
                int rows = data.GetLength(0);
                for (int i = 1; i <= rows; i++)
                {
                    string c = data[i, 1]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(c)) set.Add(c);
                }
            }
            return set;
        }

        /// <summary>
        /// Procesa una única fila de la hoja de programación y, si cumple los criterios, la añade a la hoja de consolidación masiva.
        /// </summary>
        private bool ProcesarFila(
            dynamic wb,
            object[,] data,
            int r,
            dynamic wsMas,
            ViaticosConfig cfg,
            int filaOrigen,
            ref int dest,
            HashSet<string> cat1)
        {
            try
            {
                // --- Fecha del viaje (columna H = PROG_FECHAVIAJE) ---
                object fechaRaw = data[r, ConstantsHelper.PROG_FECHAVIAJE];
                if (fechaRaw == null) return false;

                DateTime fechaViaje;
                if (fechaRaw is double d) fechaViaje = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(fechaRaw.ToString(), out fechaViaje)) return false;

                if (fechaViaje < cfg.FechaInicio || fechaViaje > cfg.FechaFin) return false;

                // --- Filtro de Dependencia ---
                string dependenciaRaw = data[r, ConstantsHelper.PROG_DEPENDENCIA]?.ToString()?.Trim() ?? "";
                string dependenciaCfg = cfg.CortaDependencia?.Trim() ?? "";
                if (!dependenciaRaw.Equals(dependenciaCfg, StringComparison.OrdinalIgnoreCase))
                {
                    LoggerHelper.Debug($"Fila {filaOrigen}: dependencia '{dependenciaRaw}' distinta de la configurada '{dependenciaCfg}'.");
                    return false;
                }

                // --- Validar cédula numérica ---
                string cedula = data[r, ConstantsHelper.PROG_CEDULA]?.ToString() ?? "";
                if (!double.TryParse(cedula, out double cedNum) || cedNum <= 0) return false;

                // --- Obtener empleado ---
                var emp = _viaticosRepository.GetEmpleado(cedula);
                if (emp == null)
                {
                    LoggerHelper.Debug($"Fila {filaOrigen}: Empleado con cédula {cedula} no encontrado en Datos.");
                    return false;
                }

                // --- Construir datos del viaje para el masivo ---
                var viaje = BuildViajeMasivo(wb, data, r);

                Llenar(wsMas, cfg, dest, emp, viaje, cat1);
                LoggerHelper.Debug($"Fila {filaOrigen}: Agregado {emp.NombreCompleto} al masivo.");
                dest++;
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error procesando fila masiva {filaOrigen}", ex);
                return false;
            }
        }

        /// <summary>
        /// Construye el objeto Viaje con la lógica del original ViaticosMasivos.ObtenerDatosViajeMasivo:
        /// - Objeto = "[CortaDependencia] EN CETAP [CiudadDestino] DEL ... AL ..." / "EL ..."
        /// - Transporte leído desde columna PROG_TRANSPORTE limpiando $, comas, .00
        /// </summary>
        /// <summary>
        /// Construye el objeto <see cref="Viaje"/> siguiendo las reglas específicas del proceso masivo.
        /// Genera el 'Objeto' de la comisión dinámicamente según la dependencia y destino.
        /// </summary>
        private Viaje BuildViajeMasivo(dynamic wb, object[,] data, int r)
        {
            var v = new Viaje();

            // Fechas
            object fi = data[r, ConstantsHelper.PROG_FECHAVIAJE];
            object ff = data[r, ConstantsHelper.PROG_FECHAFIN];
            v.FechaInicio = fi is double d1 ? DateTime.FromOADate(d1) : DateTime.TryParse(fi?.ToString(), out DateTime dt1) ? dt1 : DateTime.MinValue;
            v.FechaFin    = ff is double d2 ? DateTime.FromOADate(d2) : DateTime.TryParse(ff?.ToString(), out DateTime dt2) ? dt2 : DateTime.MinValue;

            // Duración
            v.Duracion = int.TryParse(data[r, ConstantsHelper.PROG_DURACION]?.ToString(), out int dur) ? dur : 1;

            // Itinerario → ciudades
            string itinerario = (data[r, ConstantsHelper.PROG_ITINERARIO]?.ToString() ?? "").Trim().ToUpperInvariant();
            v.Itinerario = itinerario;
            var partes = itinerario.Split('-');
            if (partes.Length >= 2)
            {
                v.CiudadOrigen  = _viaticosRepository.NormalizarCiudad(partes[0].Trim());
                v.CiudadDestino = _viaticosRepository.NormalizarCiudad(partes[1].Trim());
            }

            // Transporte: limpiar $, comas, .00
            string transporteStr = (data[r, ConstantsHelper.PROG_TRANSPORTE]?.ToString() ?? "")
                                    .Replace("$", "").Replace(",", "").Replace(".00", "").Trim();
            v.Transporte = decimal.TryParse(transporteStr, out decimal trans) ? trans : 0;

            // Objeto: construido igual que en el VBA original
            // Usa el objeto detallado (columna PROG_OBJETO) si está disponible, de lo contrario la dependencia
            string depTexto = (data[r, ConstantsHelper.PROG_DEPENDENCIA]?.ToString() ?? "").Trim();
            string objBase = (data[r, ConstantsHelper.PROG_OBJETO]?.ToString() ?? "").Trim();
            string baseTexto = !string.IsNullOrWhiteSpace(objBase) ? "PARA " + objBase : depTexto;

            string fechasObjeto = v.Duracion > 1
                ? $" DEL {v.FechaInicio:dd/MM/yyyy} AL {v.FechaFin:dd/MM/yyyy}"
                : $" EL {v.FechaInicio:dd/MM/yyyy}";
            v.Objeto = ViaticosHelper.LimpiarTexto($"{baseTexto} EN CETAP {v.CiudadDestino}{fechasObjeto}");

            // Dependencia (para el filtro en ProcesarFila)
            v.Dependencia = depTexto;

            // Cédula
            v.cedula = data[r, ConstantsHelper.PROG_CEDULA]?.ToString() ?? "";

            return v;
        }

        /// <summary>
        /// Llenado físico de las celdas en la hoja masiva consolidada.
        /// Aplica la lógica de porcentajes y tarifas diarias basadas en la categoría de ciudad y salario del empleado.
        /// </summary>
        private void Llenar(dynamic wsMas, ViaticosConfig cfg, int f, Empleado emp, Viaje v, HashSet<string> cat1)
        {
            // 1. Calcular tarifas (lógica idéntica al original)
            double sal = 0;
            bool hasSal = emp.Salario != null && double.TryParse(emp.Salario.ToString(), out sal) && sal > 0;

            double porc;
            double vd;
            double baseSal;

            if (hasSal)
            {
                porc    = 100;
                vd      = (double)(cfg.ViaticoA ?? 0);
                baseSal = sal;
            }
            else
            {
                porc    = (double)(cfg.PorcB ?? 0);
                vd      = (double)(cfg.ViaticoB ?? 0);
                baseSal = (double)(cfg.SalarioDecreto ?? 0);

                // Ciudad Categoría 1 → aplica tarifa A
                if (cat1 != null && cat1.Count > 0 && cat1.Contains(v.CiudadDestino))
                {
                    porc = (double)(cfg.PorcA ?? 0);
                    vd   = (double)(cfg.ViaticoA ?? 0);
                }
            }

            // 2. Formatear números como strings sin separadores (igual que el original)
            string transporteStr  = ((long)v.Transporte).ToString();
            string valorDiarioStr = ((long)vd).ToString();

            // 3. Escritura de celdas A–T (mapeo idéntico al original)
            wsMas.Range["A" + f].Value = f;                                                                         // A = Secuencia
            wsMas.Range["B" + f].Value = "03";                                                                      // B = Fijo "03"
            wsMas.Range["C" + f].Value = emp.Cedula;                                                                // C = Cédula
            wsMas.Range["D" + f].Value = emp.NumeroCuenta;                                                          // D = Cuenta
            wsMas.Range["E" + f].Value = emp.Categoria;                                                             // E = Categoría
            wsMas.Range["F" + f].Value = baseSal;                                                                   // F = Salario Base
            wsMas.Range["G" + f].Value = porc;                                                                      // G = %
            wsMas.Range["H" + f].Value = valorDiarioStr;                                                            // H = Tarifa Diaria (string)
            wsMas.Range["I" + f].Value = transporteStr;                                                             // I = Transporte (string)
            wsMas.Range["J" + f].Value = _viaticosRepository.ObtenerDepartamento(v.CiudadOrigen);                  // J = Dpto Origen
            wsMas.Range["K" + f].Value = ViaticosHelper.SinTildes(v.CiudadOrigen);                                 // K = Ciudad Origen
            wsMas.Range["L" + f].Value = _viaticosRepository.ObtenerDepartamento(v.CiudadDestino);                 // L = Dpto Destino
            wsMas.Range["M" + f].Value = ViaticosHelper.SinTildes(v.CiudadDestino);                                // M = Ciudad Destino
            wsMas.Range["N" + f].Value = v.FechaInicio.ToString("yyyy/MM/dd");                                      // N = Fecha Inicio
            wsMas.Range["O" + f].Value = v.FechaFin.ToString("yyyy/MM/dd");                                         // O = Fecha Fin
            wsMas.Range["P" + f].Value = "NO";                                                                      // P = Fijo "NO"
            wsMas.Range["Q" + f].Value = "SI";                                                                      // Q = Fijo "SI"
            wsMas.Range["R" + f].Value = 50;                                                                        // R = Fijo 50
            wsMas.Range["S" + f].Value = "PAGO VIATICOS " + v.Objeto?.ToUpperInvariant();                          // S = Concepto
            wsMas.Range["T" + f].Value = "SI";                                                                      // T = Fijo "SI"

            wsMas.Calculate();
        }

        private dynamic CrearHojaMasiva(dynamic wb, string name)
        {
            try
            {
                dynamic excelApp = wb.Application;

                // Eliminar hoja preexistente de forma robusta (evita DISP_E_BADINDEX)
                dynamic wsExistente = null;
                foreach (dynamic sh in wb.Worksheets)
                {
                    if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        wsExistente = sh;
                        break;
                    }
                }

                if (wsExistente != null)
                {
                    excelApp.DisplayAlerts = false;
                    wsExistente.Delete();
                    excelApp.DisplayAlerts = true;
                }

                // Copiar plantilla "Masivo"
                dynamic wsTemplate = wb.Worksheets["Masivo"];
                wsTemplate.Visible = -1; // xlSheetVisible
                wsTemplate.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);

                dynamic wsNueva = excelApp.ActiveSheet;
                wsNueva.Name = name;

                try { wsNueva.Unprotect(); } catch { }

                wsTemplate.Visible = 2; // xlSheetHidden

                // Limpiar contenido existente
                int last = (int)wsNueva.Cells[wsNueva.Rows.Count, "A"].End[-4162].Row;
                if (last >= 1) wsNueva.Range[$"A1:T{last}"].ClearContents();

                return wsNueva;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error creando hoja Masivo", ex);
                throw new Exception("Error creando hoja Masivo: " + ex.Message);
            }
        }

        /// <summary>
        /// Exporta la hoja masiva a un nuevo libro Excel independiente y lo guarda en la carpeta destino.
        /// </summary>
        private void GuardarExcel(dynamic wb, dynamic ws, string folder, string name)
        {
            string path = Path.Combine(folder, name + ".xlsx");

            // Cerrar si ya está abierto
            try
            {
                dynamic excelApp = wb.Application;
                foreach (dynamic wbOpen in excelApp.Workbooks)
                {
                    if (string.Equals(wbOpen.Name, Path.GetFileName(path), StringComparison.OrdinalIgnoreCase))
                    {
                        wbOpen.Close(false);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Detalle al verificar libro masivo abierto: {ex.Message}");
            }

            // Eliminar si existe
            if (File.Exists(path))
            {
                try { File.Delete(path); } catch { }
            }

            // Copiar hoja a nuevo libro y guardar
            ws.Copy();
            dynamic wbNuevo = ((dynamic)ws.Application).ActiveWorkbook;
            try
            {
                dynamic wsN = wbNuevo.Worksheets[1];
                try { wsN.Unprotect(); } catch { }
                wsN.UsedRange.Value = wsN.UsedRange.Value;

                if (File.Exists(ConstantsHelper.THEME_PATH))
                {
                    try { wbNuevo.Theme.ThemeColorScheme.Load(ConstantsHelper.THEME_PATH); }
                    catch (Exception ex) { LoggerHelper.Warning("No se pudo cargar el tema: " + ex.Message); }
                }

                wbNuevo.SaveAs(path, 51);
            }
            finally
            {
                try { wbNuevo.Close(false); } catch { }
            }
        }

        private void EliminarHoja(dynamic wb, string name)
        {
            wb.Application.DisplayAlerts = false;
            try { wb.Worksheets[name].Delete(); } catch { }
            wb.Application.DisplayAlerts = true;
        }
    }
}