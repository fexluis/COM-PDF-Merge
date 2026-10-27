using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.IO;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Repositories;
namespace ConversorPDF.XLL.Core.Processors
{
    /// <summary>
    /// Implementación del procesador de viáticos para generación individual o por selección.
    /// Maneja la lógica de copiado de plantillas Excel, llenado de datos y exportación a PDF con metadatos.
    /// </summary>
    internal class ViaticosIndividualProcessor : IViaticosIndividualProcessor
    {
        private const int FILA_INICIAL = 4;
        private readonly IExcelRepository _excelRepository;
        private readonly IViaticosRepository _viaticosRepository;
        private readonly IUserService _userService;
        private readonly IPdfMetadataService _metadataService;

        /// <summary>
        /// Inicializa una nueva instancia de la clase <see cref="ViaticosIndividualProcessor"/>.
        /// </summary>
        /// <param name="excelRepository">Repositorio de acceso a datos de Excel.</param>
        /// <param name="viaticosRepository">Repositorio de negocio para viáticos.</param>
        /// <param name="userService">Servicio de información de usuario.</param>
        /// <param name="metadataService">Servicio de gestión de metadatos PDF.</param>
        public ViaticosIndividualProcessor(
            IExcelRepository excelRepository, 
            IViaticosRepository viaticosRepository,
            IUserService userService,
            IPdfMetadataService metadataService)
        {
            _excelRepository = excelRepository;
            _viaticosRepository = viaticosRepository;
            _userService = userService;
            _metadataService = metadataService;
        }

        /// <summary>
        /// Procesa todos los registros válidos de la hoja 'Programación' dentro del rango de fechas configurado.
        /// </summary>
        /// <param name="progressCallback">Opcional. Callback para notificar el progreso (actual, total).</param>
        /// <param name="overrideInicio">Opcional. Fecha de inicio manual.</param>
        /// <param name="overrideFin">Opcional. Fecha de fin manual.</param>
        /// <param name="overrideCdp">Opcional. CDP manual para filtrar.</param>
        /// <returns>Número de archivos generados exitosamente.</returns>
        public int Procesar(Action<int, int> progressCallback = null, DateTime? overrideInicio = null, DateTime? overrideFin = null, string overrideCdp = null)
        {
            dynamic wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return 0;

            dynamic excelApp = wb.Application;
            dynamic activeSheet = excelApp.ActiveSheet;

            if (!_viaticosRepository.ValidarFormatoComisionado(activeSheet))
            {
                MessageBox.Show("La hoja activa no tiene el formato de Comisión esperado.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            var config = _viaticosRepository.GetConfig(overrideInicio, overrideFin, overrideCdp);
            if (config == null) return 0;
            if (string.IsNullOrWhiteSpace(config.Cdp))
            {
                MessageBox.Show("No hay CDP configurado. Regístrelo en 'Gestión de CDP Dependencias' antes de generar documentos.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            if (config.FechaInicio > config.FechaFin)
            {
                LoggerHelper.Warning("La fecha de inicio no puede ser mayor que la fecha final.");
                return 0;
            }

            dynamic wsProg = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_PROG);
            int last = (int)wsProg.Cells[wsProg.Rows.Count, "A"].End[-4162 /*xlUp*/].Row;
            if (last < FILA_INICIAL)
            {
                LoggerHelper.Warning($"No hay datos suficientes en la hoja {ConstantsHelper.SHEET_PROG}. Fila final: {last}");
                return 0;
            }

            LoggerHelper.Info($"Iniciando proceso individual para {last - FILA_INICIAL + 1} registros.");
            int contador = 0;
            using (var scope = new ExcelAppStateHelper(excelApp))
            {
                scope.ApplyOptimized();
                object[,] data = _excelRepository.GetRangeValues(wsProg, FILA_INICIAL, 1, last, ConstantsHelper.PROG_COL_MAX);
                int total = data.GetLength(0);

                for (int i = 1; i <= total; i++)
                {
                    progressCallback?.Invoke(i, total);
                    if (ProcesarRegistro(wb, data, i, config, FILA_INICIAL + i - 1, out string err)) contador++;
                    else if (!string.IsNullOrEmpty(err))
                        LoggerHelper.Debug($"Fila {FILA_INICIAL + i - 1} omitida: {err}");
                }
            }

            try { wsProg.Select(); } catch { }

            if (contador > 0)
            {
                LoggerHelper.Success($"Proceso individual finalizado exitosamente. Generados: {contador} archivos.");
                System.Diagnostics.Process.Start("explorer.exe", config.CarpetaDestino);
            }
            else
            {
                LoggerHelper.Info("Proceso individual finalizado sin generar archivos.");
            }
            return contador;
        }

        /// <summary>
        /// Procesa únicamente el registro de la fila donde se encuentra el cursor del usuario.
        /// </summary>
        /// <param name="progressCallback">Callback para progreso (1/1).</param>
        /// <returns>1 si el registro se generó con éxito, 0 en caso contrario.</returns>
        public int ProcesarFilaSeleccionada(Action<int, int> progressCallback = null)
        {
            dynamic wb = _excelRepository.GetActiveWorkbook();
            if (wb == null) return 0;
            dynamic excelApp = wb.Application;
            int rowIdx = excelApp.ActiveCell.Row;

            if (rowIdx < FILA_INICIAL)
            {
                MessageBox.Show("Seleccione una fila válida.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            var config = _viaticosRepository.GetConfig();
            if (config == null) return 0;
            dynamic wsProg = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_PROG);
            object[,] data = _excelRepository.GetRangeValues(wsProg, rowIdx, 1, rowIdx, ConstantsHelper.PROG_COL_MAX);

            if (ProcesarRegistro(wb, data, 1, config, rowIdx, out string err, true))
            {
                System.Diagnostics.Process.Start("explorer.exe", config.CarpetaDestino);
                return 1;
            }
            else
            {
                MessageBox.Show(err, App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }

        /// <summary>
        /// Realiza el procesamiento atómico de un único registro de viáticos.
        /// Cubre validación de fechas, dependencias, carga de modelo y generación de documentos.
        /// Fase 15.2: Verificación de nulo tras reasignación de cfg en isUno para evitar NullReferenceException.
        /// </summary>
        /// <param name="wb">Instancia del libro Excel activo.</param>
        /// <param name="data">Matriz completa de datos de la hoja de programación.</param>
        /// <param name="r">Índice de la fila en la matriz de datos.</param>
        /// <param name="cfg">Configuración de viáticos cargada.</param>
        /// <param name="excelRow">Número físico de la fila en Excel.</param>
        /// <param name="error">Parámetro de salida con descripción de error si el proceso falla.</param>
        /// <param name="isUno">Flag que indica si es un proceso individual (permite cargar configuración específica).</param>
        /// <returns>True si el registro se procesó correctamente.</returns>
        private bool ProcesarRegistro(dynamic wb, object[,] data, int r, ViaticosConfig cfg, int excelRow, out string error, bool isUno = false)
        {
            error = "";
            try
            {
                string dep = data[r, ConstantsHelper.PROG_DEPENDENCIA]?.ToString()?.Trim() ?? "";
                if (!dep.Equals(cfg.CortaDependencia?.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    if (isUno)
                    {
                        string cdpDep = _viaticosRepository.GetDependenciaCDP(dep);
                        if (string.IsNullOrWhiteSpace(cdpDep))
                        {
                            error = $"Dependencia '{dep}' sin CDP asociado en cdpDependencias";
                            return false;
                        }

                        cfg = _viaticosRepository.GetConfig(null, null, cdpDep);
                        if (cfg == null) { error = $"No se pudo cargar configuración CDP para dependencia '{dep}'"; return false; }
                    }
                    else { error = $"Dependencia '{dep}' no coincide con '{cfg.CortaDependencia}'"; return false; }
                }

                // Validar fecha antes de construir el viaje (igual que en el original)
                object fechaRaw = data[r, ConstantsHelper.PROG_FECHAVIAJE];
                if (fechaRaw == null) { error = "Fecha de viaje nula"; return false; }
                DateTime fechaViaje;
                if (fechaRaw is double dv) fechaViaje = DateTime.FromOADate(dv);
                else if (!DateTime.TryParse(fechaRaw.ToString(), out fechaViaje)) { error = $"Fecha inválida: {fechaRaw}"; return false; }
                if (fechaViaje < cfg.FechaInicio || fechaViaje > cfg.FechaFin) { error = $"Fecha {fechaViaje:dd/MM/yyyy} fuera de rango"; return false; }

                var viaje = _viaticosRepository.GetViajeFromRow(data, r);

                // Validar cédula
                if (!double.TryParse(viaje.cedula, out double cedNum) || cedNum <= 0)
                { error = $"Cédula inválida: {viaje.cedula}"; return false; }

                var emp = _viaticosRepository.GetEmpleado(viaje.cedula);
                if (emp == null) { error = $"Empleado no encontrado: {viaje.cedula}"; return false; }

                // Tipo de formato basado en PROG_TIPOFORMATO (columna 1) — igual que el original
                string tipoFormatoStr = data[r, ConstantsHelper.PROG_TIPOFORMATO]?.ToString() ?? "";
                if (!MapearTipoFormato(tipoFormatoStr, out var tipo))
                { error = $"Tipo de formato desconocido: '{tipoFormatoStr}'"; return false; }

                string sheetName = $"{tipoFormatoStr}-{excelRow}";
                dynamic ws = CopiarPlantilla(wb, sheetName, ConstantsHelper.GetTemplate(tipo));
                LlenarFormato(ws, cfg, emp, viaje, tipo);

                string name = $"xV-FILA-{excelRow:000}";
                string nomPdf = NombreArchivoPdf(emp.NombreCompleto, viaje.FechaInicio, excelRow);
                Guardar(wb.Application, ws, name, nomPdf, cfg.CarpetaDestino);
                
                // Guardar metadatos en el PDF generado
                string fullPdfPath = Path.Combine(cfg.CarpetaDestino, nomPdf + ".pdf");
                string fullUser = _userService.GetUsuarioWindows();
                var parts = fullUser.Split('\\');
                string cleanUser = parts.Length > 1 ? parts[1].ToUpper() : fullUser.ToUpper();

                _metadataService.GuardarMetadatos(fullPdfPath, new Metadata
                {
                    Titulo = nomPdf,
                    Autor = cleanUser,
                    Asunto = viaje.Objeto,
                    PalabrasClave = $"{emp.Cedula}, {emp.NombreCompleto}, {dep}",
                    Fecha = viaje.FechaInicio,
                    MetaDocumento = TipoMetadata.VIA_SOLICITUD,
                    CDP = cfg.Cdp,
                    Tipo = "Solicitud",
                    Comision = "0"
                });

                EliminarHoja(wb, sheetName);

                LoggerHelper.Debug($"Fila {excelRow}: Generado {name}.pdf");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                LoggerHelper.Error($"Error procesando registro fila {excelRow}", ex);
                return false;
            }
        }

        /// <summary>
        /// Mapea la cadena de PROG_TIPOFORMATO (col 1) al enum, igual que en ViaticosAutorizados original.
        /// Valores válidos: "contratistas", "servidores", "docentes catedra" / "docentes cátedra".
        /// </summary>
        private bool MapearTipoFormato(string texto, out TipoFormato tipo)
        {
            tipo = TipoFormato.Contratista;
            string t = (texto ?? "").ToLowerInvariant().Trim();
            if (t == "contratistas") { tipo = TipoFormato.Contratista; return true; }
            if (t == "servidores") { tipo = TipoFormato.Servidor; return true; }
            if (t.Contains("docentes catedra") || t.Contains("docentes cátedra")) { tipo = TipoFormato.DocenteCatedra; return true; }
            return false;
        }

        /// <summary>
        /// Duplica la hoja de plantilla especificada y la prepara para el llenado.
        /// </summary>
        /// <param name="wb">Libro activo.</param>
        /// <param name="name">Nombre único para la nueva hoja.</param>
        /// <param name="template">Nombre de la hoja plantilla original.</param>
        /// <returns>Referencia a la nueva hoja de cálculo creada.</returns>
        private dynamic CopiarPlantilla(dynamic wb, string name, string template)
        {
            dynamic wsT = null;
            try
            {
                wsT = wb.Worksheets[template];
            }
            catch
            {
                LoggerHelper.Warning($"No se encontró la plantilla '{template}'. Intentando fallback a 'Contratistas'.");
                try
                {
                    wsT = wb.Worksheets["Contratistas"];
                }
                catch
                {
                    throw new Exception($"Error crítico: No se encontró la plantilla '{template}' ni el formato base 'Contratistas'. Verifique los nombres de las hojas.");
                }
            }

            wsT.Visible = -1; // xlSheetVisible
            wsT.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            
            dynamic wsN = wb.Application.ActiveSheet;
            
            try { wsN.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }

            wsN.Name = name;
            wsT.Visible = 2; // xlSheetVeryHidden
            
            return wsN;
        }

        /// <summary>
        /// Mapea los datos del empleado y el viaje a las celdas físicas de la hoja de plantilla.
        /// Maneja tres variaciones de formato (Contratista, Servidor, Docente) con mapeos específicos.
        /// </summary>
        private void LlenarFormato(dynamic ws, ViaticosConfig cfg, Empleado emp, Viaje v, TipoFormato t)
        {
            // ── Datos comunes (celdas idénticas para todos los tipos de formato) ──
            ws.Range["C7"].Value = emp.NombreCompleto?.ToUpperInvariant();
            ws.Range["J7"].Value = emp.FechaNacimiento?.ToString()?.ToUpperInvariant() ?? "";
            ws.Range["C8"].Value = emp.Cedula;
            ws.Range["C9"].Value = emp.Telefono;
            ws.Range["G10"].Value = emp.TipoCuenta?.ToUpperInvariant() ?? "";
            ws.Range["G11"].Value = emp.NumeroCuenta;
            ws.Range["E13"].Value = v.FechaInicio;
            ws.Range["E14"].Value = v.FechaFin;
            ws.Range["H14"].Value = v.Itinerario?.ToUpperInvariant();

            // ── Cadenas reutilizables ──
            string depLarga = string.IsNullOrWhiteSpace(cfg.FullDependencia) ? cfg.CortaDependencia : cfg.FullDependencia;
            string cdpFull = $"El pago de la presente comisión de servicios / autorización de desplazamiento, se hará con cargo a {depLarga}, del Rubro {cfg.Rubro}, según Certificado de Disponibilidad Presupuestal No. {cfg.Cdp} del {cfg.Fcdp}.";
            string cdpResponsable = cfg.ResponsableCDP ?? "";
            string cdpCargo = cfg.CargoResponsableCDP ?? "";
            string elaboro = string.IsNullOrWhiteSpace(cfg.Elaboro) ? "" : "Elaboró: " + cfg.Elaboro;
            string reviso = string.IsNullOrWhiteSpace(cfg.Reviso) ? "" : "Revisó: " + cfg.Reviso;

            string gerente;
            string cargo;
            ViaticosHelper.ResolverAprobador(cfg, emp.Cedula, emp.NombreCompleto, "Processor", out gerente, out cargo);
            string aprobo = string.IsNullOrWhiteSpace(gerente) ? "" : "Aprobó: " + gerente + (string.IsNullOrWhiteSpace(cargo) ? "" : ", " + cargo);

            // ── Cálculo canónico de transporte aéreo (soporta tanto ruta BD como ruta Excel Programador) ──
            bool esAereoExcel = ViaticosHelper.EsTerminalAerea(v.terminal);

            int isTerminalAereo = (v.IsTerminalAereo == 1 || esAereoExcel) ? 1 : 0;
            string terminalNombre = !string.IsNullOrWhiteSpace(v.TerminalAereaNombre)
                ? v.TerminalAereaNombre
                : (esAereoExcel ? App.CapitalizarPalabras(v.terminal) : v.terminal);

            var (aereoHoraIda, aereoHoraRegreso, aereoRequiereTiquete, aeroTerminal) = 
                ViaticosHelper.ResolverAereo(isTerminalAereo, terminalNombre, docente: t == TipoFormato.DocenteCatedra);

            // Fase 246: Resolver valor canónico de transporte aéreo para el formato
            decimal valorTransporteAereo = isTerminalAereo == 1
                ? (v.TransporteAereo > 0 ? v.TransporteAereo : v.GastoTerminalAerea)
                : 0;

            LoggerHelper.Debug($"[Processor] [Formato] Cedula='{emp.Cedula}', Nombre='{emp.NombreCompleto}' | aeroTerminal='{aeroTerminal}', valorTransporteAereo={valorTransporteAereo:N0}, Transporte={v.Transporte:N0} | Formato={t}");

            // ── Datos específicos por tipo de formato ──
            switch (t)
            {
                case TipoFormato.Contratista:
                    ws.Range["F8"].Value = emp.Contrato?.ToUpperInvariant();
                    ws.Range["F9"].Value = emp.Correo;
                    ws.Range["C10"].Value = emp.EntidadBancaria?.ToUpperInvariant();
                    ws.Range["C11"].Value = emp.Salario?.ToString();
                    ws.Range["G13"].Value = aereoHoraIda;
                    ws.Range["G14"].Value = aereoHoraRegreso;
                    ws.Range["C15"].Value = v.Objeto?.ToUpperInvariant();
                    ws.Range["H16"].Value = v.CiudadOrigen?.ToUpperInvariant();
                    ws.Range["H17"].Value = aeroTerminal;
                    ws.Range["E26"].Value = valorTransporteAereo;
                    ws.Range["E27"].Value = v.Transporte;
                    ws.Range["B31"].Value = cdpFull;
                    // Contratista: C36 = Responsable CDP, B37 = Cargo del responsable del CDP
                    ws.Range["C36"].Value = cdpResponsable;
                    ws.Range["B37"].Value = cdpCargo;
                    ws.Range["G36"].Value = gerente;
                    ws.Range["F37"].Value = cargo;
                    ws.Range["B38"].Value = elaboro;
                    ws.Range["B39"].Value = reviso;
                    ws.Range["B40"].Value = aprobo;
                    break;

                case TipoFormato.Servidor:
                    ws.Range["F8"].Value = emp.Rol?.ToUpperInvariant();
                    ws.Range["F9"].Value = emp.Correo;
                    ws.Range["C10"].Value = emp.EntidadBancaria?.ToUpperInvariant();
                    ws.Range["C11"].Value = emp.Salario?.ToString();
                    ws.Range["G13"].Value = aereoHoraIda;
                    ws.Range["G14"].Value = aereoHoraRegreso;
                    ws.Range["C15"].Value = v.Objeto?.ToUpperInvariant();
                    ws.Range["H16"].Value = v.CiudadOrigen?.ToUpperInvariant();
                    ws.Range["H17"].Value = aeroTerminal;
                    ws.Range["E26"].Value = valorTransporteAereo;
                    ws.Range["E27"].Value = v.Transporte;
                    ws.Range["B31"].Value = cdpFull;
                    // Servidor: C36 = Responsable CDP, B37 = Cargo del responsable del CDP
                    ws.Range["C36"].Value = cdpResponsable;
                    ws.Range["B37"].Value = cdpCargo;
                    ws.Range["G36"].Value = gerente;
                    ws.Range["F37"].Value = cargo;
                    ws.Range["B38"].Value = elaboro;
                    ws.Range["B39"].Value = reviso;
                    ws.Range["B40"].Value = aprobo;
                    break;

                case TipoFormato.DocenteCatedra:
                    ws.Range["F8"].Value = emp.Correo;
                    ws.Range["F9"].Value = emp.EntidadBancaria?.ToUpperInvariant();
                    ws.Range["C11"].Value = emp.Categoria?.ToUpperInvariant();
                    ws.Range["G13"].Value = aereoHoraIda;
                    ws.Range["G14"].Value = aereoHoraRegreso;
                    ws.Range["H15"].Value = aereoRequiereTiquete;
                    ws.Range["H16"].Value = aeroTerminal;
                    ws.Range["C16"].Value = v.CiudadDestino?.ToUpperInvariant();
                    ws.Range["C17"].Value = v.Objeto?.ToUpperInvariant();
                    ws.Range["E25"].Value = valorTransporteAereo;
                    ws.Range["E26"].Value = v.Transporte;
                    ws.Range["B29"].Value = cdpFull;
                    // DocenteCatedra: G34 = Gerente, F35 = Cargo Gerente; C34 = Responsable CDP, B35 = Cargo del responsable del CDP (B36 = Elaboró, B37 = Revisó, B38 = Aprobó)
                    ws.Range["G34"].Value = gerente;
                    ws.Range["F35"].Value = cargo;
                    ws.Range["C34"].Value = cdpResponsable;
                    ws.Range["B35"].Value = cdpCargo;
                    ws.Range["B36"].Value = elaboro;
                    ws.Range["B37"].Value = reviso;
                    ws.Range["B38"].Value = aprobo;
                    break;
            }

            ws.Calculate();
        }

        /// <summary>
        /// Crea una copia del formato lleno en un nuevo libro XLSX independiente y lo exporta a PDF.
        /// Fase 19: Garantiza el cierre de wbPdf en bloque finally para evitar fugas de objetos COM si falla la exportación PDF.
        /// </summary>
        /// <param name="app">Instancia de la aplicación Excel.</param>
        /// <param name="ws">Hoja con los datos llenos.</param>
        /// <param name="nameXls">Nombre del archivo Excel resultante.</param>
        /// <param name="namePdf">Nombre del archivo PDF resultante.</param>
        /// <param name="folder">Carpeta destino de guardado.</param>
        private void Guardar(dynamic app, dynamic ws, string nameXls, string namePdf, string folder)
        {
            string pdfPath = Path.Combine(folder, namePdf + ".pdf");
            string xlsPath = Path.Combine(folder, nameXls + ".xlsx");

            ws.Copy();
            dynamic wbNuevo = app.ActiveWorkbook;
            dynamic wsNueva = wbNuevo.Worksheets[1];

            try { wsNueva.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }
            wsNueva.UsedRange.Value = wsNueva.UsedRange.Value;

            dynamic wbPdf = null;
            app.DisplayAlerts = false;
            try
            {
                if (File.Exists(xlsPath)) File.Delete(xlsPath);

                if (File.Exists(ConstantsHelper.THEME_PATH))
                {
                    try { wbNuevo.Theme.ThemeColorScheme.Load(ConstantsHelper.THEME_PATH); }
                    catch (Exception ex) { LoggerHelper.Warning("No se pudo cargar el tema: " + ex.Message); }
                }

                wbNuevo.SaveAs(xlsPath, 51 /*xlOpenXMLWorkbook*/);

                // Reabrir para exportar PDF  (fiel al original: GuardarPDFIndividual)
                string savedPath = xlsPath;
                wbPdf = app.Workbooks.Open(savedPath);
                dynamic wsPdf = wbPdf.Worksheets[1];
                App.ConfigurarPagina(wsPdf, true);
                wsPdf.ExportAsFixedFormat(0 /*xlTypePDF*/, pdfPath);
            }
            finally
            {
                app.DisplayAlerts = true;
                if (wbPdf != null)
                {
                    try { wbPdf.Close(false); } catch { }
                }
                try { wbNuevo.Close(false); } catch { }
                // Fase 13.3: Eliminar archivo Excel temporal para evitar residuos .xlsx en la carpeta de salida
                try { if (File.Exists(xlsPath)) File.Delete(xlsPath); } catch { }
            }
        }

        private static string NombreArchivoPdf(string persona, DateTime fecha, int fila)
        {
            string p = ViaticosHelper.SinTildes(persona ?? "").Replace(" ", "-");
            if (p.Length > 20) p = p.Substring(0, 20);
            return $"{p}-[{fecha:dd-MM-yy}]-FILA-{fila:000}";
        }

        private void EliminarHoja(dynamic wb, string name)
        {
            wb.Application.DisplayAlerts = false;
            try { wb.Worksheets[name].Delete(); } catch { }
            wb.Application.DisplayAlerts = true;
        }
    }
}