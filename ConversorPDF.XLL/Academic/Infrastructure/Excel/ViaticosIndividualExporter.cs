using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Services;

namespace ConversorPDF.XLL.Academic.Infrastructure.Excel
{
    public class ViaticosIndividualExporter
    {
        // Fase 633: Prefijos auxiliares y nombres neutros centralizados para generadores y reconocedores
        private const string PREFIJO_PARACAIDAS = "__tmp_guard_";
        private const string PREFIJO_RESIDUO = "__res_";
        private const string NOMBRE_HOJA_NEUTRA = "Hoja1";

        private readonly UserService _userService;
        private readonly PdfMetadataService _metadataService;

        public ViaticosIndividualExporter()
        {
            _userService = new UserService();
            _metadataService = new PdfMetadataService();
        }

        /// <summary>
        /// Exporta bloques de viáticos a formatos PDF individuales.
        /// </summary>
        /// <param name="app">Instancia de la aplicación Excel (COM).</param>
        /// <param name="bloques">Lista de bloques de viáticos a exportar.</param>
        /// <param name="cfg">Configuración global de viáticos.</param>
        /// <param name="cacheTerceros">Caché de datos de terceros.</param>
        /// <param name="progressCallback">Opcional. Callback de notificación de progreso.</param>
        /// <returns>Número de archivos PDF generados exitosamente.</returns>
        public int Exportar(dynamic app, List<ViaticoBloque> bloques, ViaticosConfig cfg, IReadOnlyDictionary<string, Tercero> cacheTerceros, Action<int, int> progressCallback = null)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            if (cfg == null) throw new ArgumentNullException(nameof(cfg));
            if (bloques == null || bloques.Count == 0) return 0;

            if (string.IsNullOrWhiteSpace(cfg.Cdp))
            {
                LoggerHelper.Warning("[Exporter] CDP vacío en configuración. Los textos de CDP en PDF quedarán incompletos.");
            }

            int iteracion = 1;
            int procesados = 0;
            int generados = 0;
            int total = bloques.Count;

            string folderDestino = cfg.CarpetaDestino;
            if (!Directory.Exists(folderDestino)) Directory.CreateDirectory(folderDestino);

            Stopwatch sw = Stopwatch.StartNew();

            dynamic wb = app.ActiveWorkbook;
            string templateFileName = "GF-FO-023.xlsx";
            string templatesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTemplates", templateFileName);

            // Importar 7 hojas
            dynamic wbTemplate = null;
            bool importacionRealizada = false;

            // Fase 207, 226 & 626: Declarar colecciones de seguimiento antes del try
            var hojasImportadas = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            var hojasTemporales = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var nombresCopias = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Desactivar actualizaciones de pantalla y cálculos automáticos para evitar parpadeo y mejorar rendimiento
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;
            try { app.Calculation = -4135; /* xlCalculationManual */ } catch { }
            try
            {
                // Fase 177: Validar existencia de la plantilla ANTES de cualquier limpieza o modificación
                if (!File.Exists(templatesPath))
                {
                    throw new FileNotFoundException($"No se encontró la plantilla en: {templatesPath}");
                }

                // Fase 191: Abrir primero la plantilla
                wbTemplate = app.Workbooks.Open(templatesPath, ReadOnly: true);

                string[] sheetsToImport = { "Docentes catedra", "Contratistas", "Servidores", "DIVIPOLA", "Categorías 1 y 2", "TARIFAS ESAP", "Decreto 314 de 2026" };

                // Fase 175 y 201: Pre-limpieza de duplicados desambiguados (" (n)") y temporales en el libro activo
                LimpiarHojasPlantilla(app, wb, nombresImportados: null, hojasTemporales: null);

                // Fase 206: Limpieza previa de hojas canónicas preexistentes que sean idénticas a la plantilla
                LimpiarPreviaPlantillas(wb, wbTemplate, sheetsToImport);

                foreach (string shName in sheetsToImport)
                {
                    try
                    {
                        dynamic sh = wbTemplate.Worksheets[shName];

                        // Fase 213, 197 & 620: Captura de nombres previa a la copia para identificar determinísticamente la nueva hoja
                        var nombresAntes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        int countAntes = (int)wb.Sheets.Count;
                        for (int i = 1; i <= countAntes; i++)
                        {
                            try { nombresAntes.Add((string)wb.Sheets[i].Name); } catch { }
                        }

                        int antesImp = (int)wb.Worksheets.Count;
                        sh.Copy(Type.Missing, wb.Worksheets[antesImp]);

                        if ((int)wb.Worksheets.Count != antesImp + 1)
                        {
                            LoggerHelper.Warning($"[Exporter] Copy no insertó la hoja '{shName}'. Importación de la hoja omitida.");
                            continue;
                        }

                        // Fase 620: Localizar determinísticamente la nueva hoja por diferencia de nombres (evita desincronización por índice posicional)
                        dynamic copia = null;
                        int countDespues = (int)wb.Sheets.Count;
                        for (int i = 1; i <= countDespues; i++)
                        {
                            try
                            {
                                dynamic cand = wb.Sheets[i];
                                string candName = (string)cand.Name;
                                if (!nombresAntes.Contains(candName))
                                {
                                    copia = cand;
                                    break;
                                }
                            }
                            catch { }
                        }

                        if (copia == null)
                        {
                            // Fase 628: Fallback posicional solo válido si la hoja en esa posición no existía antes de la copia
                            dynamic posicional = null;
                            try { posicional = wb.Worksheets[antesImp + 1]; } catch { }
                            string nomPos = null;
                            try { nomPos = posicional == null ? null : (string)posicional.Name; } catch { }

                            if (posicional != null && nomPos != null && !nombresAntes.Contains(nomPos))
                            {
                                copia = posicional;
                            }
                            else
                            {
                                LoggerHelper.Warning($"[Exporter] No se pudo identificar la copia de '{shName}'; importación omitida.");
                                continue;
                            }
                        }

                        hojasImportadas[shName] = copia;
                        string nombreCopia = (string)copia.Name;

                        // Fase 626: Registrar el nombre real de la copia para purga segura sin tocar hojas homónimas del usuario
                        nombresCopias.Add(nombreCopia);

                        if (!string.Equals(nombreCopia, shName, StringComparison.OrdinalIgnoreCase))
                        {
                            LoggerHelper.Warning($"[Exporter] Plantilla '{shName}' renombrada por Excel a '{nombreCopia}'.");
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Warning($"No se pudo importar '{shName}' de la plantilla: {ex.Message}");
                    }
                }

                if (hojasImportadas.Count > 0)
                {
                    importacionRealizada = true;
                }

                try { wbTemplate.Close(false); } catch { }
                wbTemplate = null;

                // Desbloquear configuración en las hojas importadas
                foreach (string baseName in new[] { "Docentes catedra", "Contratistas", "Servidores" })
                {
                    if (hojasImportadas.TryGetValue(baseName, out dynamic hb))
                    {
                        try { hb.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }
                    }
                }

                // Fase 182: Precalcular la identidad del usuario una sola vez fuera del ciclo
                string fullUser = _userService.GetUsuarioWindows();
                var parts = fullUser.Split('\\');
                string cleanUser = parts.Length > 1 ? parts[1].ToUpper() : fullUser.ToUpper();

                // Fase 173: Contador monotónico independiente para evitar colisiones de nombre en la colección Worksheets
                int consecutivoHoja = 0;

                foreach (var bloque in bloques)
                {
                    // Fase 194 & 210: Notificar progreso por total de procesados protegiendo contra excepciones de UI
                    procesados++;
                    try { progressCallback?.Invoke(procesados, total); }
                    catch (Exception ex) { LoggerHelper.Warning($"[Exporter] Callback de progreso falló: {ex.Message}"); }

                    string cedulaBloque = bloque.Sesiones.FirstOrDefault()?.CedulaTercero;
                    if (string.IsNullOrWhiteSpace(cedulaBloque)) continue;

                    if (!cacheTerceros.TryGetValue(cedulaBloque, out Tercero tercero))
                    {
                        LoggerHelper.Warning($"[Exporter] Tercero no encontrado en caché para cédula {cedulaBloque}");
                        continue; // No procesar si no se halla el tercero
                    }

                    // Identificar tipo de plantilla
                    string tipoStr = (tercero.Tipo ?? "").ToUpperInvariant().Trim();
                    TipoFormato tipoFormato;
                    string nombreHojaBase;

                    if (tipoStr == "CATEDRA") { tipoFormato = TipoFormato.DocenteCatedra; nombreHojaBase = "Docentes catedra"; }
                    else if (tipoStr == "CONTRATISTA") { tipoFormato = TipoFormato.Contratista; nombreHojaBase = "Contratistas"; }
                    else { tipoFormato = TipoFormato.Servidor; nombreHojaBase = "Servidores"; }

                    // Fase 223: Sanitizar el prefijo eliminando caracteres inválidos para hojas de Excel y truncando a 24 caracteres max
                    string tipoPrefijo = new string((tipoStr ?? "").Where(ch => char.IsLetterOrDigit(ch) || ch == '-').ToArray());
                    if (tipoPrefijo.Length == 0) tipoPrefijo = "FORMATO";
                    if (tipoPrefijo.Length > 24) tipoPrefijo = tipoPrefijo.Substring(0, 24);

                    consecutivoHoja++;
                    string sheetName = $"{tipoPrefijo}-{consecutivoHoja}";
                    while (ExisteHoja(wb, sheetName))
                    {
                        consecutivoHoja++;
                        sheetName = $"{tipoPrefijo}-{consecutivoHoja}";
                    }

                    // Fase 226: Registrar el nombre exacto de la hoja temporal creada en el registro dinámico
                    hojasTemporales.Add(sheetName);

                    // Mapeo
                    var emp = MapearEmpleado(tercero, cfg);
                    var viaje = MapearViaje(bloque, tercero, cfg);

                    // Fase 176 y 184: Obtener referencia directa a la hoja base importada con TryGetValue
                    if (!hojasImportadas.TryGetValue(nombreHojaBase, out dynamic wsBase) || wsBase == null)
                    {
                        LoggerHelper.Warning($"[Exporter] Plantilla '{nombreHojaBase}' no importada. Bloque {cedulaBloque} omitido.");
                        continue;
                    }

                    // Fase 170, 185 y 218: Duplicar y verificar explícitamente que Worksheets.Count incrementó
                    dynamic wsNueva = null;
                    try
                    {
                        try { wsBase.Visible = -1; /* xlSheetVisible */ }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[Exporter] No se pudo mostrar la base '{nombreHojaBase}': {ex.Message}");
                            continue;
                        }

                        int antesDup = (int)wb.Worksheets.Count;
                        dynamic ultimoSheet = wb.Worksheets[antesDup];
                        wsBase.Copy(Type.Missing, ultimoSheet);

                        if ((int)wb.Worksheets.Count != antesDup + 1)
                        {
                            LoggerHelper.Error($"[Exporter] Copy no creó la hoja temporal para {cedulaBloque} ('{sheetName}'). Bloque omitido.");
                            continue;
                        }

                        wsNueva = wb.Worksheets[antesDup + 1];
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error($"[Exporter] Falló la duplicación de '{nombreHojaBase}' para {cedulaBloque}", ex);
                        continue;
                    }
                    finally
                    {
                        // Fase 172: Restaurar estado oculto con protección try/catch
                        try { wsBase.Visible = 2; /* xlSheetVeryHidden */ } catch { }
                    }

                    // Fase 171: Encapsular el procesamiento del bloque en try/catch/finally por iteración
                    try
                    {
                        try { wsNueva.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }
                        wsNueva.Name = sheetName;

                        // Llenar Formato
                        LlenarFormato(app, wsNueva, cfg, emp, viaje, tipoFormato);

                        // Fase 183: Congelar valores con referencia cacheada y protegida
                        try
                        {
                            dynamic rangoUsado = wsNueva.UsedRange;
                            rangoUsado.Value = rangoUsado.Value;
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[Exporter] No se pudo congelar valores en '{sheetName}': {ex.Message}");
                        }

                        // Fase 224 y 229: Sanitizar y truncar nombres de archivo PDF a 40 caracteres para evitar desbordamiento MAX_PATH
                        string baseRutaName = SanitizarNombreArchivo(viaje.CiudadDestino, "VARIOS");
                        if (baseRutaName.Length > 40) baseRutaName = baseRutaName.Substring(0, 40).TrimEnd('.', ' ', '_');

                        string p = SanitizarNombreArchivo(ViaticosHelper.SinTildes(emp.NombreCompleto ?? "").Replace(" ", "-"), "SIN-NOMBRE");
                        if (p.Length > 20) p = p.Substring(0, 20);

                        string nomPdf = $"[{iteracion}]-{p}-[{viaje.FechaInicio:dd-MM-yy}]-{baseRutaName}";

                        // Fase 230: Desambiguación de nombres si el archivo PDF resultante ya existe en la carpeta destino
                        string rutaPdf = Path.Combine(folderDestino, nomPdf + ".pdf");
                        for (int k = 2; File.Exists(rutaPdf); k++)
                        {
                            nomPdf = $"[{iteracion}]-{p}-[{viaje.FechaInicio:dd-MM-yy}]-{baseRutaName}-({k})";
                            rutaPdf = Path.Combine(folderDestino, nomPdf + ".pdf");
                        }

                        // Fase 227: Generar token único Guid para asegurar el nombre del archivo temporal XLSX entre corridas
                        string tokenTemporal = Guid.NewGuid().ToString("N").Substring(0, 8);
                        GuardarPDF(app, wsNueva, $"xV-TEMP-{tokenTemporal}-{consecutivoHoja:000}", nomPdf, folderDestino);

                        // Fase 181: Incrementar contador inmediatamente tras el éxito del PDF
                        generados++;
                        iteracion++;

                        // Fase 181: Guardar Metadatos en try/catch no fatal
                        try
                        {
                            string fullPdfPath = Path.Combine(folderDestino, nomPdf + ".pdf");
                            _metadataService.GuardarMetadatos(fullPdfPath, new ConversorPDF.XLL.Core.Models.Metadata
                            {
                                Titulo = nomPdf,
                                Autor = cleanUser,
                                Asunto = viaje.Objeto,
                                PalabrasClave = $"{emp.Cedula}, {emp.NombreCompleto}, {emp.Dependencia}",
                                Fecha = viaje.FechaInicio,
                                MetaDocumento = ConversorPDF.XLL.Core.Enums.TipoMetadata.VIA_SOLICITUD,
                                CDP = cfg.Cdp,
                                Tipo = "Solicitud",
                                Comision = "0"
                            });
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[Exporter] Metadatos no aplicados a '{nomPdf}': {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error($"[Exporter] Error procesando bloque {iteracion} (cédula {cedulaBloque}). Continúa con el siguiente.", ex);
                    }
                    finally
                    {
                        // Fase 174, 186, 198, 200: Eliminación defensiva de la hoja temporal comprobando EsUnicaHojaVisible y desocultando temporalmente si aplica
                        if (wsNueva != null)
                        {
                            try
                            {
                                app.DisplayAlerts = false;
                                EliminarHojaTemporal(wb, wsNueva, sheetName, excluir: null);
                            }
                            catch (Exception ex)
                            {
                                LoggerHelper.Warning($"[Exporter] No se pudo eliminar la hoja temporal '{sheetName}': {ex.Message}");
                            }
                        }
                    }
                }
            }
            finally
            {
                // Fase 178: Cierre garantizado del libro de plantilla si permanecía abierto
                try { if (wbTemplate != null) wbTemplate.Close(false); } catch { }
                wbTemplate = null;

                // Fase 175, 191, 201, 226 & 626: Limpieza final pasando nombres reales de copias importadas y registro dinámico de temporales
                if (importacionRealizada || hojasTemporales.Count > 0)
                {
                    LimpiarHojasPlantilla(app, wb, nombresCopias, hojasTemporales);
                }

                app.DisplayAlerts = true;
                app.ScreenUpdating = true;
                try { app.Calculation = -4105; /* xlCalculationAutomatic */ } catch { }
            }

            sw.Stop();
            LoggerHelper.Info($"[Exporter] Proceso de {total} bloques completado en {sw.ElapsedMilliseconds} ms. Generados: {generados}");

            return generados; // Retornar cantidad real de archivos PDF generados exitosamente
        }

        private Empleado MapearEmpleado(Tercero t, ViaticosConfig cfg)
        {
            return new Empleado
            {
                Cedula = t.Cedula,
                NombreCompleto = t.Nombre,
                Salario = t.Salario,
                Rol = t.Tipo,
                Categoria = t.Categoria,
                EntidadBancaria = t.EntidadBancaria,
                NumeroCuenta = t.CuentaBancaria,
                TipoCuenta = t.TipoCuentaBancaria,
                Telefono = t.Telefono,
                Correo = t.Correo,
                FechaNacimiento = t.FechaDeNacimiento,
                Contrato = t.Contrato,
                Dependencia = cfg.FullDependencia // o usar t.Dependencia si hubiera
            };
        }

        /// <summary>
        /// Mapea la información de un bloque de sesiones académicas, tercero y configuración global a la entidad Viaje.
        /// Se extrae la configuración de transporte aéreo (isTerminalAereo, idTerminales, terminalAereaNombre, gastoTerminalAerea).
        /// </summary>
        /// <param name="b">Bloque de sesiones agrupadas para la comisión.</param>
        /// <param name="t">Entidad del tercero comisionado.</param>
        /// <param name="cfg">Configuración global de viáticos.</param>
        /// <returns>Entidad Viaje mapeada con rutas, fechas, indicadores de terminal y valores de transporte.</returns>
        private Viaje MapearViaje(ViaticoBloque b, Tercero t, ViaticosConfig cfg)
        {
            string origen = !string.IsNullOrWhiteSpace(t.Residencia) ? t.Residencia : cfg.CiudadBase;
            string destino = b.Sesiones.LastOrDefault(s => !string.IsNullOrWhiteSpace(s.SedeCiudad))?.SedeCiudad ?? "S/D";
            string iti = ConstruirItinerario(b, origen, destino);

            // Armar el objeto según el original
            string objBase = b.Sesiones.FirstOrDefault()?.AsignaturaNombre ?? "ACTIVIDADES ACADEMICAS";
            DateTime fIni = b.FechaMinima ?? DateTime.Now;
            DateTime fFin = b.FechaFinComision ?? b.FechaMaxima ?? DateTime.Now;
            // dur = días calendario de comisión (redondeo AwayFromZero: 0.5->1, 1.5->2, 2.5->3); b.DuracionTotal = días de viático
            int dur = (int)Math.Round(b.DuracionTotal, MidpointRounding.AwayFromZero);
            string fechasObjeto = dur > 1
                ? $" DEL {fIni:dd/MM/yyyy} AL {fFin:dd/MM/yyyy}"
                : $" EL {fIni:dd/MM/yyyy}";

            string objetoFinal = ViaticosHelper.LimpiarTexto($"PAGO VIATICOS PARA DICTAR ASIGNATURA DE {objBase} EN EL CETAP DE {destino}{fechasObjeto}");

            decimal transporte = 0;
            var ultimaSesion = b.Sesiones.OrderBy(s => s.FechaSesion).LastOrDefault();
            if (ultimaSesion?.TransporteSede.HasValue == true)
            {
                transporte = ultimaSesion.TransporteSede.Value;
            }

            // Detección de configuración aérea en las sesiones del bloque
            var sesionAerea = b.Sesiones.FirstOrDefault(s => s.IsTerminalAereo == 1) ?? b.Sesiones.FirstOrDefault();
            int isTerminalAereo = sesionAerea?.IsTerminalAereo ?? 0;
            int idTerminales = sesionAerea?.IdTerminales ?? 0;
            string terminalAereaNombre = sesionAerea?.TerminalAereaNombre ?? "";
            decimal gastoTerminalAerea = (decimal)(sesionAerea?.GastoTerminalAerea ?? 0);

            decimal transporteAereo = 0m;
            string terminalNombre = string.Empty;

            if (isTerminalAereo == 1)
            {
                if (gastoTerminalAerea > 0)
                {
                    transporteAereo = gastoTerminalAerea;
                    terminalNombre = terminalAereaNombre;
                }
                else
                {
                    // Fase 246: Resolver tope del catálogo si en datos legacy o vinculación no venía gasto definido
                    if (ViaticosHelper.ResolverTopeTerminal(idTerminales, terminalAereaNombre, "Exporter.MapearViaje",
                        out int idRes, out string nomRes, out decimal maxVal))
                    {
                        transporteAereo = maxVal;
                        gastoTerminalAerea = maxVal;
                        terminalNombre = !string.IsNullOrWhiteSpace(terminalAereaNombre) ? terminalAereaNombre : nomRes;
                        if (idTerminales <= 1) idTerminales = idRes;
                    }
                    else
                    {
                        terminalNombre = terminalAereaNombre;
                    }
                }
            }

            return new Viaje
            {
                cedula = t.Cedula,
                FechaInicio = fIni,
                FechaFin = fFin,
                Duracion = dur,
                CiudadOrigen = origen,
                CiudadDestino = destino,
                Itinerario = iti,
                Transporte = transporte,
                TransporteAereo = transporteAereo,
                Objeto = objetoFinal,
                terminal = terminalNombre,
                IsTerminalAereo = isTerminalAereo,
                IdTerminales = idTerminales,
                TerminalAereaNombre = terminalAereaNombre,
                GastoTerminalAerea = gastoTerminalAerea
            };
        }

        private string ConstruirItinerario(ViaticoBloque bloque, string origen, string destino)
        {
            var sedesOrdenadas = bloque.Sesiones
                .Where(s => !string.IsNullOrWhiteSpace(s.SedeCiudad) || !string.IsNullOrWhiteSpace(s.Subsede))
                .OrderBy(s => s.FechaSesion)
                .Select(s => (string.IsNullOrWhiteSpace(s.SedeCiudad) ? s.Subsede : s.SedeCiudad).ToUpper())
                .ToList();

            var rutaCompleta = new List<string>();
            rutaCompleta.Add(origen.ToUpper());
            rutaCompleta.AddRange(sedesOrdenadas);

            var rutaLimpia = new List<string>();
            for (int i = 0; i < rutaCompleta.Count; i++)
            {
                if (i == 0 || rutaCompleta[i] != rutaCompleta[i - 1])
                    rutaLimpia.Add(rutaCompleta[i]);
            }

            string escala = null;
            if (rutaLimpia.Count > 2)
            {
                escala = string.Join("-", rutaLimpia.Skip(1).Take(rutaLimpia.Count - 2));
            }

            var segmentos = new List<string>();
            segmentos.Add(origen.ToUpper());

            if (!string.IsNullOrWhiteSpace(escala))
            {
                segmentos.AddRange(escala.Split('-').Select(s => s.Trim().ToUpper()).Where(s => !string.IsNullOrWhiteSpace(s)));
            }

            if (!string.IsNullOrWhiteSpace(destino)) segmentos.Add(destino.ToUpper());
            if (!string.Equals(destino, origen, StringComparison.OrdinalIgnoreCase))
                segmentos.Add(origen.ToUpper());

            return string.Join(" - ", segmentos);
        }

        /// <summary>
        /// Llena los datos en la hoja de Excel mapeando las entidades Empleado y Viaje a las celdas de la plantilla correspondiente.
        /// Sincroniza los campos de transporte aéreo (horas de ida/regreso, requerimiento de tiquete y nombre de terminal/aeropuerto)
        /// según el tipo de formato (Contratista, Servidor, DocenteCatedra).
        /// </summary>
        /// <param name="app">Instancia COM de la aplicación Excel, usada para forzar recálculo.</param>
        /// <param name="ws">Objeto dinámico de la hoja de cálculo de Excel (Worksheet).</param>
        /// <param name="cfg">Configuración global de viáticos (CDP, dependencias, firmas, etc.).</param>
        /// <param name="emp">Entidad del Empleado con la información personal y bancaria.</param>
        /// <param name="v">Entidad del Viaje con rutas, fechas, indicadores de terminal y valores calculados.</param>
        /// <param name="t">Tipo de formato a utilizar, derivado del tipo de tercero.</param>
        private void LlenarFormato(dynamic app, dynamic ws, ViaticosConfig cfg, Empleado emp, Viaje v, TipoFormato t)
        {
            ws.Range["C7"].Value = emp.NombreCompleto?.ToUpperInvariant();
            ws.Range["J7"].Value = emp.FechaNacimiento?.ToString()?.ToUpperInvariant() ?? "";
            ws.Range["C8"].Value = emp.Cedula;
            ws.Range["C9"].Value = emp.Telefono;
            ws.Range["G10"].Value = emp.TipoCuenta?.ToUpperInvariant() ?? "";
            ws.Range["G11"].Value = emp.NumeroCuenta;
            ws.Range["E13"].Value = v.FechaInicio;
            ws.Range["E14"].Value = v.FechaFin;
            ws.Range["H14"].Value = v.Itinerario?.ToUpperInvariant();

            string depLarga = string.IsNullOrWhiteSpace(cfg.FullDependencia) ? cfg.CortaDependencia : cfg.FullDependencia;
            string cdpFull = $"El pago de la presente comisión de servicios / autorización de desplazamiento, se hará con cargo a {depLarga}, del Rubro {cfg.Rubro}, según Certificado de Disponibilidad Presupuestal No. {cfg.Cdp} del {cfg.Fcdp}.";
            string cdpResponsable = cfg.ResponsableCDP ?? "";
            string cdpCargo = cfg.CargoResponsableCDP ?? "";
            string elaboro = string.IsNullOrWhiteSpace(cfg.Elaboro) ? "" : "Elaboró: " + cfg.Elaboro;
            string reviso = string.IsNullOrWhiteSpace(cfg.Reviso) ? "" : "Revisó: " + cfg.Reviso;

            string gerente;
            string cargo;
            ViaticosHelper.ResolverAprobador(cfg, emp.Cedula, emp.NombreCompleto, "Exporter", out gerente, out cargo);
            string aprobo = string.IsNullOrWhiteSpace(gerente) ? "" : "Aprobó: " + gerente + (string.IsNullOrWhiteSpace(cargo) ? "" : ", " + cargo);

            // Variables de transporte aéreo comunes (unificadas vía ViaticosHelper)
            var (aereoHoraIda, aereoHoraRegreso, aereoRequiereTiquete, aeroTerminal) = 
                ViaticosHelper.ResolverAereo(
                    v.IsTerminalAereo, 
                    !string.IsNullOrWhiteSpace(v.TerminalAereaNombre) ? v.TerminalAereaNombre : v.terminal, 
                    docente: t == TipoFormato.DocenteCatedra);

            // Fase 246: Valor consolidado de transporte aéreo para el formato
            decimal valorTransporteAereo = v.IsTerminalAereo == 1
                ? (v.TransporteAereo > 0 ? v.TransporteAereo : v.GastoTerminalAerea)
                : 0;

            LoggerHelper.Debug($"[Exporter] [Formato] Cedula='{emp.Cedula}', Nombre='{emp.NombreCompleto}' | aeroTerminal='{aeroTerminal}', valorTransporteAereo={valorTransporteAereo:N0}, Transporte={v.Transporte:N0} | Formato={t}");

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

            // Nota sobre cálculo: Se asume que la hoja está desprotegida temporalmente (línea 113) 
            // Corrección Fase 6: Forzar recálculo completo antes de exportar, manejando el escenario donde Excel opera en cálculo manual.
            try { app.CalculateUntilAsyncQueriesDone(); } catch { }
            ws.Calculate();
        }

        /// <summary>
        /// Determina si una hoja o gráfico de Excel es visible, considerando devoluciones de Variant booleanos o enteros (-1 = xlSheetVisible).
        /// Fase 199: Resuelve la lectura errónea de Variant booleano en Excel Interop.
        /// </summary>
        private static bool EsVisible(dynamic hoja)
        {
            if (hoja == null) return false;
            object v;
            try { v = hoja.Visible; } catch { return false; }
            if (v is bool b) return b;
            try { return Convert.ToInt32(v) == -1; } catch { return false; }
        }

        /// <summary>
        /// Verifica si una hoja con el nombre especificado existe en el libro activo.
        /// Fase 173: Previene colisiones de nombre al renombrar hojas temporales.
        /// </summary>
        private static bool ExisteHoja(dynamic wb, string nombre)
        {
            if (wb == null || string.IsNullOrWhiteSpace(nombre)) return false;
            int n = (int)wb.Worksheets.Count;
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    if (string.Equals((string)wb.Worksheets[i].Name, nombre, StringComparison.OrdinalIgnoreCase)) return true;
                }
                catch { }
            }
            return false;
        }

        /// <summary>
        /// Verifica si la hoja especificada por su nombre es la única hoja visible en todo el libro (incluyendo ChartSheets).
        /// Fase 200 y 199: Inspecciona wb.Sheets por nombre y evalúa visibilidad con EsVisible para evitar errores COM 1004.
        /// </summary>
        private static bool EsUnicaHojaVisible(dynamic wb, string nombreExcluido)
        {
            if (wb == null || string.IsNullOrWhiteSpace(nombreExcluido)) return false;
            int n = (int)wb.Sheets.Count;
            for (int i = 1; i <= n; i++)
            {
                dynamic s;
                try { s = wb.Sheets[i]; } catch { continue; }
                try
                {
                    if (string.Equals((string)s.Name, nombreExcluido, StringComparison.OrdinalIgnoreCase)) continue;
                }
                catch { continue; }

                if (EsVisible(s)) return false;
            }
            return true;
        }

        /// <summary>
        /// Sanitiza una cadena para su uso seguro como parte del nombre de un archivo en el sistema de archivos local,
        /// reemplazando los caracteres no válidos de ruta por guiones bajos y descartando nombres reservados de Windows.
        /// Fase 224 y 231: Previene excepciones de I/O por rutas inválidas o nombres reservados (CON, PRN, AUX, NUL, COM1..4, LPT1..3).
        /// </summary>
        private static string SanitizarNombreArchivo(string valor, string fallback)
        {
            if (string.IsNullOrWhiteSpace(valor)) return fallback;
            var invalidos = Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder(valor.Length);
            foreach (char ch in valor) sb.Append(Array.IndexOf(invalidos, ch) >= 0 ? '_' : ch);
            string limpio = sb.ToString().Trim();
            if (limpio.Length == 0) return fallback;

            // Fase 231: Descartar nombres reservados de Windows
            string[] reservados = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "LPT1", "LPT2", "LPT3" };
            string sinExt = Path.GetFileNameWithoutExtension(limpio);
            if (Array.Exists(reservados, r => string.Equals(r, sinExt, StringComparison.OrdinalIgnoreCase)))
            {
                limpio = "_" + limpio;
            }
            return limpio;
        }

        /// <summary>
        /// Recorre las primeras celdas del UsedRange de una hoja (hasta 25x25) para obtener el primer valor no nulo ni vacío,
        /// devolviendo además su dirección de celda y las dimensiones totales del UsedRange.
        /// Fase 219: Proporciona una firma de contenido robusta para comparar hojas sin depender de celdas A1/B1 fijas.
        /// </summary>
        private static object PrimerValorNoVacio(dynamic hoja, out string direccion, out int filas, out int cols)
        {
            direccion = null; filas = 0; cols = 0;
            try
            {
                dynamic ur = hoja.UsedRange;
                filas = (int)ur.Rows.Count;
                cols = (int)ur.Columns.Count;
                int maxF = Math.Min(filas, 25), maxC = Math.Min(cols, 25);
                for (int f = 1; f <= maxF; f++)
                {
                    for (int c = 1; c <= maxC; c++)
                    {
                        dynamic celda = ur.Cells[f, c];
                        object v = null;
                        try { v = celda.Value; } catch { }
                        if (v == null || (v is string s && string.IsNullOrWhiteSpace(s))) continue;
                        direccion = (string)celda.Address(false, false);
                        return v;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Intenta eliminar la hoja temporal de trabajo. Si es la única visible, desoculta temporalmente otra hoja
        /// (buscando en Sheets y excluyendo hojas pendientes de borrado, con fallback si todo el libro es residual).
        /// Fase 198, 214, 220, 221, 222, 225, 232 y 621: Garantiza eliminación mediante desprotección defensiva y paracaídas visible.
        /// </summary>
        private static void EliminarHojaTemporal(dynamic wb, dynamic wsNueva, string sheetName, HashSet<string> excluir = null)
        {
            if (wb == null || wsNueva == null) return;

            // Fase 621: Desprotección defensiva previa (hojas protegidas rechazan Delete en Excel)
            try { wsNueva.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }
            try { wsNueva.Unprotect(); } catch { }

            bool borrada = false;
            try { wsNueva.Delete(); borrada = true; } catch { }

            if (!borrada)
            {
                dynamic temporal = null;
                bool otraVisibleForzada = false;
                try
                {
                    // Fase 221 & 232: Buscar auxiliar sobre la colección Sheets comprobando tipos válidos (xlWorksheet | xlChart)
                    int nSheets = (int)wb.Sheets.Count;
                    for (int j = 1; j <= nSheets; j++)
                    {
                        dynamic cand = null;
                        try { cand = wb.Sheets[j]; } catch { continue; }

                        int tipoCand = -1;
                        try { tipoCand = Convert.ToInt32(cand.Type); } catch { continue; }
                        if (tipoCand != -4167 && tipoCand != -4168) continue;

                        string nomCand;
                        try { nomCand = (string)cand.Name; } catch { continue; }
                        if (string.Equals(nomCand, sheetName, StringComparison.OrdinalIgnoreCase)) continue;

                        // Fase 222: Excluir hojas que también van a ser eliminadas en esta misma pasada
                        if (excluir != null && excluir.Contains(nomCand)) continue;

                        temporal = cand;
                        break;
                    }

                    // Fase 225: Fallback si todo el libro es residual (excluir descartó todas las hojas)
                    if (temporal == null)
                    {
                        for (int j = 1; j <= nSheets; j++)
                        {
                            dynamic cand = null;
                            try { cand = wb.Sheets[j]; } catch { continue; }

                            int tipoCand = -1;
                            try { tipoCand = Convert.ToInt32(cand.Type); } catch { continue; }
                            if (tipoCand != -4167 && tipoCand != -4168) continue;

                            string nomCand;
                            try { nomCand = (string)cand.Name; } catch { continue; }
                            if (string.Equals(nomCand, sheetName, StringComparison.OrdinalIgnoreCase)) continue;

                            temporal = cand;
                            otraVisibleForzada = true;
                            break;
                        }
                    }

                    if (temporal != null)
                    {
                        int visibilidadPrevia = 0;
                        bool otraVisibleCreada = false;
                        try
                        {
                            visibilidadPrevia = Convert.ToInt32(temporal.Visible);
                            if (visibilidadPrevia != -1)
                            {
                                temporal.Visible = -1; // xlSheetVisible
                                otraVisibleCreada = true;
                            }
                        }
                        catch { }

                        try { wsNueva.Delete(); borrada = true; } catch { }

                        if (otraVisibleCreada && !otraVisibleForzada)
                        {
                            // Fase 220: Trazar advertencia explícita si Excel impide restaurar la visibilidad previa de la hoja auxiliar
                            try { temporal.Visible = visibilidadPrevia; }
                            catch (Exception ex)
                            {
                                LoggerHelper.Warning($"[Exporter] Auxiliar '{temporal.Name}' queda visible (Excel exige >= 1 hoja visible): {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[Exporter] Hoja '{sheetName}' no eliminable tras desocultar aux: {ex.Message}");
                }

                // Fase 621, 624 & 625: Fallback con paracaídas temporal insertado al final y purga garantizada
                if (!borrada)
                {
                    dynamic paracaidas = null;
                    try
                    {
                        // Fase 624: Insertar al final para no desplazar índices del bucle descendente de purga
                        paracaidas = wb.Worksheets.Add(Type.Missing, wb.Worksheets[(int)wb.Worksheets.Count]);
                        try { paracaidas.Name = $"{PREFIJO_PARACAIDAS}{Guid.NewGuid().ToString("N").Substring(0, 8)}"; }
                        catch (Exception exNom) { LoggerHelper.Warning($"[Exporter] Paracaídas sin renombrar ({exNom.Message})."); }
                        paracaidas.Visible = -1; // xlSheetVisible
                        try { wsNueva.Delete(); borrada = true; } catch { }
                    }
                    catch { }
                    finally
                    {
                        // Fase 625: Purgar el paracaídas; se conserva únicamente si es la última hoja visible del libro
                        if (borrada && paracaidas != null)
                        {
                            try
                            {
                                string nomPar = null;
                                try { nomPar = (string)paracaidas.Name; } catch { }
                                if (!string.IsNullOrEmpty(nomPar) && !EsUnicaHojaVisible(wb, nomPar))
                                {
                                    paracaidas.Delete();
                                }
                            }
                            catch { }
                        }
                    }
                }
            }

            if (!borrada)
            {
                LoggerHelper.Error($"[Exporter] Residuo visible '{sheetName}' permanece en el libro.");
            }
            else
            {
                LoggerHelper.Info($"[Exporter] Hoja temporal '{sheetName}' eliminada exitosamente.");
            }
        }

        /// <summary>
        /// Exporta la hoja de Excel activa a formato PDF y limpia archivos temporales intermedios.
        /// Fase 202, 215 y 216: Identifica inequívocamente el nuevo libro creado por ws.Copy(), marca Saved=true antes de cerrar
        /// y registra advertencias si la eliminación de archivos temporales .xlsx falla tras reintentos.
        /// </summary>
        /// <param name="app">Instancia COM de la aplicación Excel.</param>
        /// <param name="ws">La hoja de cálculo que contiene el formato lleno.</param>
        /// <param name="nameXls">Nombre temporal del archivo XLSX.</param>
        /// <param name="namePdf">Nombre final del archivo PDF resultante.</param>
        /// <param name="folder">Ruta de la carpeta de destino donde se guardarán los archivos.</param>
        private void GuardarPDF(dynamic app, dynamic ws, string nameXls, string namePdf, string folder)
        {
            string pdfPath = Path.Combine(folder, namePdf + ".pdf");
            string xlsPath = Path.Combine(folder, nameXls + ".xlsx");

            var nombresAntes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (dynamic w in app.Workbooks)
                {
                    try { nombresAntes.Add((string)w.Name); } catch { }
                }
            }
            catch { }

            ws.Copy();

            dynamic wbNuevo = null;
            try
            {
                foreach (dynamic w in app.Workbooks)
                {
                    string nom;
                    try { nom = (string)w.Name; } catch { continue; }
                    if (!nombresAntes.Contains(nom))
                    {
                        wbNuevo = w;
                        break;
                    }
                }
            }
            catch { }

            if (wbNuevo == null || (int)wbNuevo.Worksheets.Count != 1)
            {
                throw new InvalidOperationException("[Exporter] ws.Copy() no creó un libro nuevo; exportación abortada para no alterar el libro del usuario.");
            }

            dynamic wsPdfNueva = wbNuevo.Worksheets[1];
            try { wsPdfNueva.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }

            try
            {
                if (File.Exists(xlsPath)) File.Delete(xlsPath);

                if (File.Exists(ConstantsHelper.THEME_PATH))
                {
                    try { wbNuevo.Theme.ThemeColorScheme.Load(ConstantsHelper.THEME_PATH); }
                    catch { }
                }

                wbNuevo.SaveAs(xlsPath, 51 /*xlOpenXMLWorkbook*/);
                dynamic wsPdf = wbNuevo.Worksheets[1];
                App.ConfigurarPagina(wsPdf, true);
                wsPdf.ExportAsFixedFormat(0 /*xlTypePDF*/, pdfPath);
            }
            finally
            {
                // Fase 216: Marcar Saved=true para evitar cuadros modales de guardado antes de Close(false)
                try { wbNuevo.Saved = true; } catch { }
                try { wbNuevo.Close(false); }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[Exporter] No se pudo cerrar libro temporal '{xlsPath}'", ex);
                }

                // Fase 215: Registrar advertencia con la última excepción si la eliminación del .xlsx falla
                Exception ultimoError = null;
                for (int i = 0; i < 3 && File.Exists(xlsPath); i++)
                {
                    try { File.Delete(xlsPath); }
                    catch (Exception ex) { ultimoError = ex; Thread.Sleep(150); }
                }
                if (File.Exists(xlsPath))
                {
                    LoggerHelper.Warning($"[Exporter] Residuo no eliminado '{xlsPath}': {ultimoError?.Message}");
                }
            }
        }

        /// <summary>
        /// Examina el libro del usuario en busca de hojas preexistentes con nombres canónicos de plantilla ("DIVIPOLA", "Docentes catedra", etc.).
        /// Utiliza PrimerValorNoVacio para verificar si el contenido de la hoja existente es idéntico al de la plantilla oficial con margen de tolerancia de UsedRange.
        /// Fase 209, 212, 219 y 228: Elimina la hoja idéntica previa desocultando una auxiliar si aplica y previene borrados erróneos.
        /// </summary>
        private static void LimpiarPreviaPlantillas(dynamic wb, dynamic wbTemplate, string[] targetNames)
        {
            if (wb == null || wbTemplate == null || targetNames == null) return;

            foreach (string name in targetNames)
            {
                try
                {
                    if (ExisteHoja(wb, name))
                    {
                        dynamic plantilla = null;
                        try { plantilla = wbTemplate.Worksheets[name]; } catch { }
                        dynamic existente = null;
                        try { existente = wb.Worksheets[name]; } catch { }

                        if (plantilla != null && existente != null)
                        {
                            // Fase 219 y 228: Inspección dinámica del primer valor no vacío y dimensiones de UsedRange con margen
                            string dirTpl = null, dirExi = null;
                            int fTpl = 0, cTpl = 0, fExi = 0, cExi = 0;
                            object valTpl = PrimerValorNoVacio(plantilla, out dirTpl, out fTpl, out cTpl);
                            object valExi = PrimerValorNoVacio(existente, out dirExi, out fExi, out cExi);

                            bool identica = valTpl != null
                                         && string.Equals(dirTpl, dirExi, StringComparison.OrdinalIgnoreCase)
                                         && Equals(valTpl, valExi)
                                         && Math.Abs(fTpl - fExi) <= 2
                                         && Math.Abs(cTpl - cExi) <= 2;

                            if (identica)
                            {
                                // Fase 212 & 621: Usar EliminarHojaTemporal para desocultar auxiliar o usar paracaídas si es la única visible
                                EliminarHojaTemporal(wb, existente, name);
                                if (ExisteHoja(wb, name))
                                {
                                    // Fase 622: Desambiguar residuo renombrándolo para liberar el nombre canónico y evitar colisión en cadena
                                    try
                                    {
                                        string renomDef = $"{PREFIJO_RESIDUO}{Guid.NewGuid().ToString("N").Substring(0, 4)}_{name}";
                                        if (renomDef.Length > 31) renomDef = renomDef.Substring(0, 31);
                                        existente.Name = renomDef;
                                        LoggerHelper.Warning($"[Exporter] Hoja previa '{name}' no eliminable; renombrada defensivamente a '{renomDef}' para liberar nombre canónico.");
                                    }
                                    catch (Exception exRenom)
                                    {
                                        LoggerHelper.Warning($"[Exporter] Hoja previa '{name}' no eliminable ni renombrable ({exRenom.Message}); el formato usará '{name} (n)'.");
                                    }
                                }
                                else
                                {
                                    LoggerHelper.Info($"[Exporter] Hoja plantilla previa idéntica '{name}' eliminada.");
                                }
                            }
                            else
                            {
                                LoggerHelper.Warning($"[Exporter] Hoja '{name}' difiere de la plantilla; se conserva. Importación usará '{name} (n)'.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[Exporter] Error verificando hoja preexistente '{name}': {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Realiza una limpieza exhaustiva en el libro activo de todas las hojas plantilla e impresiones temporales residuales.
        /// Recolecta primero en un conjunto las hojas a eliminar e incluye las registradas dinámicamente en hojasTemporales.
        /// Fase 211, 217, 222 y 226: Elimina activamente hojas residuales sin usar como auxiliar otras víctimas en la misma pasada.
        /// </summary>
        /// <param name="app">Instancia COM de Excel.</param>
        /// <param name="wb">Libro activo de Excel.</param>
        /// <param name="nombresImportados">Nombres REALES de las hojas copiadas desde la plantilla en esta ejecución (Fase 626). Si es null, opera en modo pre-limpieza y purga duplicados desambiguados "X (n)".</param>
        /// <param name="hojasTemporales">Registro dinámico de nombres exactos de hojas temporales creadas en la corrida.</param>
        private static void LimpiarHojasPlantilla(dynamic app, dynamic wb, ICollection<string> nombresImportados = null, ISet<string> hojasTemporales = null)
        {
            if (app == null || wb == null) return;
            string[] targetNames = {
                "Docentes catedra", "Contratistas", "Servidores", "DIVIPOLA",
                "Categorías 1 y 2", "TARIFAS ESAP", "Decreto 314 de 2026"
            };

            try
            {
                app.DisplayAlerts = false;

                // Fase 222: Primera pasada para recolectar hojas pendientes de eliminación
                var pendientes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int totalSheets = (int)wb.Sheets.Count;
                for (int i = 1; i <= totalSheets; i++)
                {
                    try
                    {
                        dynamic sh = wb.Sheets[i];
                        string name = sh.Name?.ToString() ?? "";
                        // Fase 626: Identidad estricta por nombre real de copia importada.
                        // La desambiguación canónica solo aplica en la pre-limpieza (nombresImportados == null).
                        bool esImportada = nombresImportados != null && System.Linq.Enumerable.Contains(nombresImportados, name, StringComparer.OrdinalIgnoreCase);
                        bool esDesambiguada = nombresImportados == null && System.Linq.Enumerable.Any(targetNames, tn =>
                            name.StartsWith(tn + " (", StringComparison.OrdinalIgnoreCase) ||
                            name.StartsWith(tn + "(", StringComparison.OrdinalIgnoreCase));

                        // Fase 226 & 623: Verificar si está contenida en el registro dinámico de hojasTemporales o prefijos auxiliares (__tmp_guard_, __res_)
                        bool esHojaTemp = (hojasTemporales != null && hojasTemporales.Contains(name)) ||
                                          name.StartsWith("CATEDRA-", StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith("CONTRATISTA-", StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith("SERVIDOR-", StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith("FORMATO-", StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith("xV-TEMP-", StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith(PREFIJO_PARACAIDAS, StringComparison.OrdinalIgnoreCase) ||
                                          name.StartsWith(PREFIJO_RESIDUO, StringComparison.OrdinalIgnoreCase);

                        if (esHojaTemp || esImportada || esDesambiguada)
                        {
                            pendientes.Add(name);
                        }
                    }
                    catch { }
                }

                // Segunda pasada: Recálculo dinámico de Sheets.Count en cada iteración descendente excluyendo hojas pendientes como auxiliar
                for (int i = (int)wb.Sheets.Count; i >= 1; i--)
                {
                    try
                    {
                        dynamic sh = wb.Sheets[i];
                        string name = sh.Name?.ToString() ?? "";

                        if (pendientes.Contains(name))
                        {
                            if ((int)wb.Sheets.Count <= 1) break; // Excel exige mantener al menos 1 hoja activa

                            // Fase 211 & 217: Eliminar activamente la hoja si es la única visible comprobando el tipo xlWorksheet (-4167)
                            if (EsVisible(sh) && EsUnicaHojaVisible(wb, name))
                            {
                                bool esHojaCalculo = false;
                                try { esHojaCalculo = Convert.ToInt32(sh.Type) == -4167; } catch { }

                                if (esHojaCalculo)
                                {
                                    EliminarHojaTemporal(wb, sh, name, pendientes);
                                    continue;
                                }
                            }

                            // Fase 621 & 627: Desprotección previa y registro de log únicamente si Delete() tiene éxito
                            try { sh.Unprotect(ConstantsHelper.FORMATO_PASSWORD); } catch { }
                            try
                            {
                                sh.Delete();
                                LoggerHelper.Info($"[Exporter] Hoja residual eliminada: '{name}'.");
                            }
                            catch (Exception exDel)
                            {
                                LoggerHelper.Warning($"[Exporter] No se pudo eliminar la hoja residual '{name}' (índice {i}): {exDel.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Warning($"[Exporter] No se pudo procesar eliminación de hoja residual en índice {i}: {ex.Message}");
                    }
                }

                // Fase 629: Reconciliar hojas auxiliares supervivientes (Excel exige >= 1 hoja visible)
                try
                {
                    int restantes = (int)wb.Sheets.Count;
                    for (int i = 1; i <= restantes; i++)
                    {
                        dynamic shRes = wb.Sheets[i];
                        string nameRes = shRes.Name?.ToString() ?? "";

                        if (nameRes.StartsWith(PREFIJO_RESIDUO, StringComparison.OrdinalIgnoreCase))
                        {
                            int sep = nameRes.IndexOf('_', PREFIJO_RESIDUO.Length); // PREFIJO_RESIDUO{token}_{NombreOriginal}
                            string original = (sep > 0 && nameRes.Length > sep + 1) ? nameRes.Substring(sep + 1) : null;
                            if (!string.IsNullOrWhiteSpace(original) && !ExisteHoja(wb, original))
                            {
                                try
                                {
                                    shRes.Name = original;
                                    LoggerHelper.Info($"[Exporter] Nombre canónico '{original}' restaurado en hoja residual superviviente.");
                                }
                                catch (Exception exRes)
                                {
                                    LoggerHelper.Warning($"[Exporter] Hoja residual '{nameRes}' conserva nombre técnico (no renombrable): {exRes.Message}");
                                }
                            }
                        }
                        else if (nameRes.StartsWith(PREFIJO_PARACAIDAS, StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                shRes.Name = NOMBRE_HOJA_NEUTRA;
                                LoggerHelper.Warning($"[Exporter] Hoja auxiliar '{nameRes}' renombrada a '{NOMBRE_HOJA_NEUTRA}' (superviviente de limpieza).");
                            }
                            catch (Exception exGua)
                            {
                                LoggerHelper.Warning($"[Exporter] Hoja auxiliar '{nameRes}' conserva nombre técnico (no renombrable): {exGua.Message}");
                            }
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[Exporter] Error en rutina LimpiarHojasPlantilla", ex);
            }
        }
    }
}
