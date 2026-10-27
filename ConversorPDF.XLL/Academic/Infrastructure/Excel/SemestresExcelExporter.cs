using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Excel
{
    /// <summary>
    /// Servicio de exportación a Excel para la vista de sesiones académicas agrupadas por semestres.
    /// Utiliza optimización de COM de Excel para volcados de matriz directa y coloración de calendarios.
    /// </summary>
    public class SemestresExcelExporter
    {
        private readonly dynamic _excelApp;

        public SemestresExcelExporter(dynamic excelApp)
        {
            _excelApp = excelApp;
        }

        /// <summary>
        /// Realiza la exportación de las sesiones hacia la hoja activa de Excel, dibujando un calendario dinámico con las horas consolidadas por día,
        /// diferenciando los días festivos y los días de sesiones presenciales mediante un código de colores específico.
        /// </summary>
        /// <param name="items">Colección de sesiones académicas a consolidar.</param>
        /// <param name="periodoNombre">Nombre descriptivo del periodo académico.</param>
        /// <param name="festivos">Lista de rangos de fechas (Inicio, Fin) que se consideran como días FESTIVOS.</param>
        /// <param name="presenciales">Lista de rangos de fechas (Inicio, Fin) que se consideran como sesiones PRESENCIALES para viáticos.</param>
        public void Exportar(List<SesionListItem> items, string periodoNombre, List<(DateTime Inicio, DateTime Fin)> festivos, List<(DateTime Inicio, DateTime Fin)> presenciales, Dictionary<string, Tercero> terceros)
        {
            dynamic ws = null;
            dynamic allCells = null;
            
            dynamic tituloC1 = null;
            dynamic tituloC2 = null;
            dynamic rangeTitulo = null;
            dynamic tituloFont = null;
            
            dynamic rangeFechas = null;
            
            dynamic rangeDias = null;
            
            dynamic subtotalHeaderCell = null;
            
            dynamic row1 = null;
            dynamic row2 = null;
            dynamic fontRow1 = null;
            dynamic fontRow2 = null;
            
            dynamic startCell = null;
            dynamic endCell = null;
            dynamic dataRange = null;
            
            dynamic horasFmtC1 = null;
            dynamic horasFmtC2 = null;
            dynamic horasFormatRange = null;
            
            dynamic dateFmtC1 = null;
            dynamic dateFmtC2 = null;
            dynamic dateFormatRange = null;
            
            dynamic subFmtC1 = null;
            dynamic subFmtC2 = null;
            dynamic subtotalFormatRange = null;
            
            dynamic borderC1 = null;
            dynamic borderC2 = null;
            dynamic borderRange = null;
            dynamic borders = null;
            
            dynamic colAll = null;
            dynamic col2 = null;
            dynamic col5 = null;
            dynamic col6 = null;
            dynamic col7 = null;
            dynamic col8 = null;

            if (items == null || !items.Any(s => s.FechaSesion.HasValue))
            {
                return;
            }

            try
            {
                dynamic wb = null;
                try
                {
                    wb = _excelApp.ActiveWorkbook;
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[SemestresExcelExporter] Error al consultar ActiveWorkbook: {ex.Message}");
                }

                if (wb == null)
                {
                    try
                    {
                        wb = _excelApp.Workbooks.Add();
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("No se pudo obtener ni crear un libro de trabajo en Excel.", ex);
                    }
                }

                // Fase 639: Insertar siempre una hoja nueva al final del libro para preservar intactas las hojas del usuario
                dynamic hojas = null;
                dynamic ultimoSheet = null;
                try
                {
                    hojas = wb.Worksheets;
                    int count = (int)hojas.Count;
                    ultimoSheet = hojas[count];
                    ws = hojas.Add(Type.Missing, ultimoSheet);

                    // Asignar nombre descriptivo "Semestres" con desambiguación segura si ya existe
                    try
                    {
                        string baseNombre = "Semestres";
                        string nombreHoja = baseNombre;
                        int sufijo = 2;
                        bool nombreAsignado = false;

                        while (!nombreAsignado && sufijo <= 50)
                        {
                            try
                            {
                                ws.Name = nombreHoja;
                                nombreAsignado = true;
                            }
                            catch
                            {
                                nombreHoja = $"{baseNombre} ({sufijo})";
                                sufijo++;
                            }
                        }
                    }
                    catch { }

                    ws.Activate();
                }
                finally
                {
                    if (ultimoSheet != null) { try { Marshal.ReleaseComObject(ultimoSheet); } catch { } }
                    if (hojas != null) { try { Marshal.ReleaseComObject(hojas); } catch { } }
                    if (wb != null) { try { Marshal.ReleaseComObject(wb); } catch { } }
                }

                if (ws == null) throw new InvalidOperationException("No se pudo inicializar la hoja de cálculo en Excel.");

                using (ExcelOptimizeHelper.Optimize(_excelApp))
                {
                    allCells = ws.Cells;
                    if (allCells == null)
                        throw new InvalidOperationException(
                            "No se pudo acceder al objeto Cells de la hoja activa. " +
                            "Excel puede estar en estado de recuperación o la hoja estar corrupta. " +
                            "Cierre Excel, verifique que no haya diálogos de error abiertos e inténtelo de nuevo.");

                    try { allCells.Clear(); }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        throw new InvalidOperationException(
                            "No se pudo limpiar la hoja activa de Excel. " +
                            "Causas posibles: la hoja está protegida, Excel no responde, o error de comunicación COM. " +
                            "Asegúrese de tener una hoja de cálculo desprotegida y activa.", ex);
                    }

                    // 1. Calcular fechas globales
                    var fechasSesiones = items
                        .Where(s => s.FechaSesion.HasValue)
                        .Select(s => s.FechaSesion.Value.Date)
                        .ToList();

                    if (!fechasSesiones.Any())
                        return;

                    DateTime minDate = fechasSesiones.Min();
                    DateTime maxDate = fechasSesiones.Max();
                    List<DateTime> allDates = new List<DateTime>();
                    for (DateTime d = minDate; d <= maxDate; d = d.AddDays(1))
                        allDates.Add(d);

                    var diasFestivos = new HashSet<DateTime>();
                    foreach (var f in festivos)
                        for (DateTime d = f.Inicio; d <= f.Fin; d = d.AddDays(1))
                            diasFestivos.Add(d.Date);

                    var diasPresenciales = new HashSet<DateTime>();
                    foreach (var p in presenciales)
                        for (DateTime d = p.Inicio; d <= p.Fin; d = d.AddDays(1))
                            diasPresenciales.Add(d.Date);

                    // 2. Título Combinado en Fila 1 de Col 1 a 11
                    tituloC1 = ws.Cells[1, 1];
                    tituloC2 = ws.Cells[1, 11];
                    rangeTitulo = ws.Range[tituloC1, tituloC2];
                    rangeTitulo.Merge();
                    rangeTitulo.Value = "PROGRAMACION DOCENTE PERIODO " + (periodoNombre ?? "").Trim();
                    tituloFont = rangeTitulo.Font;
                    tituloFont.Bold = true;
                    tituloFont.Size = 12;
                    try { Marshal.ReleaseComObject(tituloFont); } catch { }
                    tituloFont = null;
                    rangeTitulo.HorizontalAlignment = -4108;

                    // 3. Dibujar Cabeceras
                    string[] cabecerasFijas = {
                        "No.", "Programa", "Sem.", "Horas", "Asignatura",
                        "Sede", "Residencia", "Docente", "Fecha Inicio", "Fecha Final", "Días/Horas Semana"
                    };

                    for (int i = 0; i < cabecerasFijas.Length; i++)
                    {
                        dynamic headerCell = null;
                        try
                        {
                            headerCell = ws.Cells[2, i + 1];
                            headerCell.Value = cabecerasFijas[i];
                        }
                        catch (Exception ex) when (IsNonCriticalException(ex))
                        {
                            // Ignorar fallo en cabecera individual; continuar con las demás
                        }
                        finally
                        {
                            if (headerCell != null) { try { Marshal.ReleaseComObject(headerCell); } catch { } }
                        }
                    }

                    int colBase = 12;
                    
                    object[,] fechasRow1 = new object[1, allDates.Count];
                    object[,] diasRow2 = new object[1, allDates.Count];

                    for (int i = 0; i < allDates.Count; i++)
                    {
                        DateTime date = allDates[i];
                        fechasRow1[0, i] = date.ToOADate();
                        diasRow2[0, i] = GetDiaSemanaInicial(date.DayOfWeek);
                    }

                    rangeFechas = null;
                    rangeDias = null;
                    try
                    {
                        rangeFechas = ws.Range[ws.Cells[1, colBase], ws.Cells[1, colBase + allDates.Count - 1]];
                        if (rangeFechas != null)
                        {
                            rangeFechas.Value = fechasRow1;
                            rangeFechas.NumberFormat = "dd/mm/yyyy";
                            rangeFechas.Orientation = 90;
                            rangeFechas.ColumnWidth = 2;
                            rangeFechas.HorizontalAlignment = -4108; // xlCenter
                        }

                        rangeDias = ws.Range[ws.Cells[2, colBase], ws.Cells[2, colBase + allDates.Count - 1]];
                        if (rangeDias != null)
                        {
                            rangeDias.Value = diasRow2;
                            rangeDias.HorizontalAlignment = -4108; // xlCenter
                        }
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en formateo de cabeceras de fechas
                    }
                    finally
                    {
                        if (rangeDias != null)   { try { Marshal.ReleaseComObject(rangeDias); } catch { } rangeDias = null; }
                        if (rangeFechas != null) { try { Marshal.ReleaseComObject(rangeFechas); } catch { } rangeFechas = null; }
                    }

                    int colSubTotal = colBase + allDates.Count;
                    try
                    {
                        subtotalHeaderCell = ws.Cells[2, colSubTotal];
                        if (subtotalHeaderCell != null) subtotalHeaderCell.Value = "SubTotal";
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en escritura de cabecera Subtotal
                    }
                    finally
                    {
                        if (subtotalHeaderCell != null) { try { Marshal.ReleaseComObject(subtotalHeaderCell); } catch { } subtotalHeaderCell = null; }
                    }

                    row1 = ws.Rows[1];
                    row2 = ws.Rows[2];
                    if (row1 != null)
                    {
                        fontRow1 = row1.Font;
                        if (fontRow1 != null)
                        {
                            fontRow1.Bold = true;
                            try { Marshal.ReleaseComObject(fontRow1); } catch { }
                            fontRow1 = null;
                        }
                        row1.RowHeight = 65;
                    }
                    if (row2 != null)
                    {
                        fontRow2 = row2.Font;
                        if (fontRow2 != null)
                        {
                            fontRow2.Bold = true;
                            try { Marshal.ReleaseComObject(fontRow2); } catch { }
                            fontRow2 = null;
                        }
                    }

                    // 4. Agrupar Datos (Fase 488/490: incluir Grupo y Modalidad para no colapsar cohortes homónimas)
                    var grupos = items
                        .GroupBy(s => new
                        {
                            s.ProgramaDetalle,
                            s.Semestre,
                            s.AsignaturaNombre,
                            SedeCompleta = string.IsNullOrWhiteSpace(s.Subsede) ? s.SedeCiudad : $"{s.SedeCiudad}({s.Subsede})",
                            Residencia = terceros != null && terceros.TryGetValue(s.CedulaTercero ?? "", out var t) ? t.Residencia : "",
                            s.DocenteNombre,
                            s.Grupo,
                            s.Modalidad
                        })
                        .OrderBy(g => g.Key.ProgramaDetalle)
                        .ThenBy(g => SemestreParser.Parsear(g.Key.Semestre))
                        .ThenBy(g => g.Key.Semestre)
                        .ThenBy(g => g.Key.AsignaturaNombre)
                        .ThenBy(g => g.Key.SedeCompleta)
                        .ThenBy(g => g.Key.DocenteNombre)
                        .ToList();

                    // Fase 512: Logging de proyección con detección de duplicados
                    var dups = grupos.Select(g => $"{g.Key.Semestre}|{g.Key.Grupo}|{g.Key.AsignaturaNombre}|{g.Key.DocenteNombre}|{g.Key.Modalidad}").ToList();
                    var rep = dups.GroupBy(x => x).Where(x => x.Count() > 1).Select(x => $"{x.Key}×{x.Count()}").ToList();
                    LoggerHelper.Info($"[Proyeccion] filas={grupos.Count} | clavesRepetidas={(rep.Count > 0 ? string.Join(" ; ", rep) : "NINGUNA")} | columnasDia={allDates.Count}");

                    if (grupos.Count == 0) return;

                    // 5. Preparar Matriz para Escritura Masiva
                    int filasTotales = grupos.Count;
                    int columnasTotales = colSubTotal;
                    object[,] matrizDatos = new object[filasTotales, columnasTotales];
                    var celdasViatica = new HashSet<(int Row, int Col)>();

                    for (int r = 0; r < grupos.Count; r++)
                    {
                        var grupo = grupos[r];
                        var sesionesGrupo = grupo.ToList();

                        matrizDatos[r, 0] = r + 1;
                        matrizDatos[r, 1] = grupo.Key.ProgramaDetalle;
                        matrizDatos[r, 2] = grupo.Key.Semestre;
                        
                        double totalHoras = Math.Round(sesionesGrupo.Sum(s => s.Horas), 1);
                        matrizDatos[r, 3] = totalHoras;
                        matrizDatos[r, 4] = grupo.Key.AsignaturaNombre;
                        matrizDatos[r, 5] = grupo.Key.SedeCompleta;
                        matrizDatos[r, 6] = grupo.Key.Residencia;
                        matrizDatos[r, 7] = grupo.Key.DocenteNombre;

                        var fechasGrupoValidas = sesionesGrupo
                            .Where(s => s.FechaSesion.HasValue)
                            .Select(s => s.FechaSesion.Value.Date)
                            .OrderBy(d => d)
                            .ToList();

                        if (fechasGrupoValidas.Any())
                        {
                            matrizDatos[r, 8] = fechasGrupoValidas.First().ToOADate();
                            matrizDatos[r, 9] = fechasGrupoValidas.Last().ToOADate();
                        }

                        // Fases 876, 878 & 1136: Columna K canónica "DIA|horas|horario" vía DiasHorasSemanaFormatter
                        var colK = DiasHorasSemanaFormatter.DesdeSesionListItems(sesionesGrupo);
                        matrizDatos[r, 10] = !string.IsNullOrWhiteSpace(colK) ? colK : null;

                        double subTotalFila = 0;
                        for (int i = 0; i < allDates.Count; i++)
                        {
                            DateTime currentDate = allDates[i];
                            int matCol = (colBase - 1) + i;
                            var sesionesDia = sesionesGrupo
                                .Where(s => s.FechaSesion.HasValue && s.FechaSesion.Value.Date == currentDate)
                                .ToList();

                            if (sesionesDia.Any())
                            {
                                double horasDia = sesionesDia.Sum(s => s.Horas);
                                matrizDatos[r, matCol] = horasDia;
                                subTotalFila += horasDia;

                                if (sesionesDia.Any(s => s.Viatica))
                                {
                                    celdasViatica.Add((r + 3, colBase + i));
                                }
                            }
                        }
                        matrizDatos[r, colSubTotal - 1] = subTotalFila > 0 ? (object)subTotalFila : null;
                    }

                    // Fase 495 & Fase 503: Limpiar bloque previo antes del volcado a todo el ancho disponible para evitar filas y columnas residuales
                    try
                    {
                        dynamic rangoLimpieza = ws.Range[ws.Cells[3, 1], ws.Cells[ws.Rows.Count, ws.Columns.Count]];
                        rangoLimpieza.ClearContents();
                        Marshal.ReleaseComObject(rangoLimpieza);
                    }
                    catch { }

                    // 6. Volcar Datos en una sola operación
                    startCell = ws.Cells[3, 1];
                    endCell = ws.Cells[3 + filasTotales - 1, columnasTotales];
                    
                    if (startCell == null || endCell == null)
                        throw new InvalidOperationException("No se pudo definir el rango de datos en Excel. Verifique que la hoja tenga suficientes filas y columnas disponibles.");

                    dataRange = ws.Range[startCell, endCell];
                    if (dataRange == null)
                        throw new InvalidOperationException("No se pudo crear el rango de datos en Excel.");

                    try { dataRange.Value = matrizDatos; }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        throw new InvalidOperationException(
                            "No se pudieron escribir los datos en Excel. " +
                            "Causas posibles: Excel está ocupado, la hoja está protegida, o error de comunicación COM. " +
                            "Cierre cualquier diálogo abierto en Excel e inténtelo de nuevo.", ex);
                    }

                    // Formatear Horas (Col 4)
                    try
                    {
                        horasFmtC1 = ws.Cells[3, 4];
                        horasFmtC2 = ws.Cells[3 + filasTotales - 1, 4];
                        horasFormatRange = ws.Range[horasFmtC1, horasFmtC2];
                        if (horasFormatRange != null) horasFormatRange.NumberFormat = "0.0";
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en formato de horas
                    }
                    finally
                    {
                        if (horasFormatRange != null) { try { Marshal.ReleaseComObject(horasFormatRange); } catch { } horasFormatRange = null; }
                        if (horasFmtC2 != null)       { try { Marshal.ReleaseComObject(horasFmtC2); } catch { } horasFmtC2 = null; }
                        if (horasFmtC1 != null)       { try { Marshal.ReleaseComObject(horasFmtC1); } catch { } horasFmtC1 = null; }
                    }

                    // Formatear Fechas (Col 9 y 10)
                    try
                    {
                        dateFmtC1 = ws.Cells[3, 9];
                        dateFmtC2 = ws.Cells[3 + filasTotales - 1, 10];
                        dateFormatRange = ws.Range[dateFmtC1, dateFmtC2];
                        if (dateFormatRange != null) dateFormatRange.NumberFormat = "dd/mm/yyyy";
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en formato de fechas
                    }
                    finally
                    {
                        if (dateFormatRange != null) { try { Marshal.ReleaseComObject(dateFormatRange); } catch { } dateFormatRange = null; }
                        if (dateFmtC2 != null)       { try { Marshal.ReleaseComObject(dateFmtC2); } catch { } dateFmtC2 = null; }
                        if (dateFmtC1 != null)       { try { Marshal.ReleaseComObject(dateFmtC1); } catch { } dateFmtC1 = null; }
                    }

                    // Formatear SubTotal
                    try
                    {
                        subFmtC1 = ws.Cells[3, colSubTotal];
                        subFmtC2 = ws.Cells[3 + filasTotales - 1, colSubTotal];
                        subtotalFormatRange = ws.Range[subFmtC1, subFmtC2];
                        if (subtotalFormatRange != null) subtotalFormatRange.NumberFormat = "0.0";
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en formato de subtotal
                    }
                    finally
                    {
                        if (subtotalFormatRange != null) { try { Marshal.ReleaseComObject(subtotalFormatRange); } catch { } subtotalFormatRange = null; }
                        if (subFmtC2 != null)            { try { Marshal.ReleaseComObject(subFmtC2); } catch { } subFmtC2 = null; }
                        if (subFmtC1 != null)            { try { Marshal.ReleaseComObject(subFmtC1); } catch { } subFmtC1 = null; }
                    }

                    // Colores (Festivos y Presenciales)
                    // Nota: Si un día coincide en ambos rangos, se prioriza visualmente el Festivo.
                    for (int i = 0; i < allDates.Count; i++)
                    {
                        DateTime currentDate = allDates[i];
                        bool esFestivo = diasFestivos.Contains(currentDate.Date);
                        bool esPresencial = diasPresenciales.Contains(currentDate.Date);

                        if (esFestivo)
                        {
                            AplicarColorColumna(ws, colBase + i, filasTotales, 0xB9B9FD);
                        }
                        else if (esPresencial)
                        {
                            AplicarColorColumna(ws, colBase + i, filasTotales, 0xE0E0E0);
                            
                            dynamic headerCell = null;
                            dynamic headerInterior = null;
                            try
                            {
                                headerCell = ws.Cells[1, colBase + i];
                                headerInterior = headerCell.Interior;
                                headerInterior.Color = 0xFFCC99;
                            }
                            catch (Exception ex) when (IsNonCriticalException(ex))
                            {
                                // Ignorar fallo
                            }
                            finally
                            {
                                if (headerInterior != null) { try { Marshal.ReleaseComObject(headerInterior); } catch { } }
                                if (headerCell != null) { try { Marshal.ReleaseComObject(headerCell); } catch { } }
                            }
                        }
                    }

                    // Colores Específicos para Viáticos (Prioridad sobre Festivo/Presencial)
                    foreach (var (row, col) in celdasViatica)
                    {
                        dynamic cell = null;
                        dynamic cellInterior = null;
                        try
                        {
                            cell = ws.Cells[row, col];
                            cellInterior = cell.Interior;
                            cellInterior.Color = 0x5AEC84; // #84EC5A
                        }
                        catch (Exception ex) when (IsNonCriticalException(ex))
                        {
                            // Ignorar fallo en celda individual; continuar con las demás celdas viática
                        }
                        finally
                        {
                            if (cellInterior != null) { try { Marshal.ReleaseComObject(cellInterior); } catch { } }
                            if (cell != null) { try { Marshal.ReleaseComObject(cell); } catch { } }
                        }
                    }

                    // Aplicar formato de cuadrícula al rango explícito
                    try
                    {
                        borderC1 = ws.Cells[1, 1];
                        borderC2 = ws.Cells[3 + filasTotales - 1, colSubTotal];
                        borderRange = ws.Range[borderC1, borderC2];
                        borders = borderRange.Borders;
                        borders.LineStyle = 1;
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en bordes; el documento sigue funcional sin cuadrícula
                    }
                    finally
                    {
                        if (borders != null)     { try { Marshal.ReleaseComObject(borders); } catch { } borders = null; }
                        if (borderRange != null) { try { Marshal.ReleaseComObject(borderRange); } catch { } borderRange = null; }
                        if (borderC2 != null)    { try { Marshal.ReleaseComObject(borderC2); } catch { } borderC2 = null; }
                        if (borderC1 != null)    { try { Marshal.ReleaseComObject(borderC1); } catch { } borderC1 = null; }
                    }

                    try
                    {
                        colAll = ws.Columns;
                        colAll.AutoFit();
                        col2 = ws.Columns[2];
                        col5 = ws.Columns[5];
                        col6 = ws.Columns[6];
                        col7 = ws.Columns[7];
                        col8 = ws.Columns[8];
                        if (col2 != null) col2.ColumnWidth = 15;
                        if (col5 != null) col5.ColumnWidth = 25;
                        if (col6 != null) col6.ColumnWidth = 25;
                        if (col7 != null) col7.ColumnWidth = 15;
                        if (col8 != null) col8.ColumnWidth = 25;
                    }
                    catch (Exception ex) when (IsNonCriticalException(ex))
                    {
                        // Ignorar fallo en autoajuste; el documento se exporta con anchos por defecto
                    }
                }
            }
            finally
            {
                // Liberar en orden inverso (hijos primero, padres después)
                if (col8 != null)          { try { Marshal.ReleaseComObject(col8); } catch { } }
                if (col7 != null)          { try { Marshal.ReleaseComObject(col7); } catch { } }
                if (col6 != null)          { try { Marshal.ReleaseComObject(col6); } catch { } }
                if (col5 != null)          { try { Marshal.ReleaseComObject(col5); } catch { } }
                if (col2 != null)          { try { Marshal.ReleaseComObject(col2); } catch { } }
                if (colAll != null)        { try { Marshal.ReleaseComObject(colAll); } catch { } }
                if (borders != null)       { try { Marshal.ReleaseComObject(borders); } catch { } }
                if (borderRange != null)   { try { Marshal.ReleaseComObject(borderRange); } catch { } }
                if (borderC2 != null)      { try { Marshal.ReleaseComObject(borderC2); } catch { } }
                if (borderC1 != null)      { try { Marshal.ReleaseComObject(borderC1); } catch { } }
                if (subtotalFormatRange != null) { try { Marshal.ReleaseComObject(subtotalFormatRange); } catch { } }
                if (subFmtC2 != null)      { try { Marshal.ReleaseComObject(subFmtC2); } catch { } }
                if (subFmtC1 != null)      { try { Marshal.ReleaseComObject(subFmtC1); } catch { } }
                if (dateFormatRange != null)     { try { Marshal.ReleaseComObject(dateFormatRange); } catch { } }
                if (dateFmtC2 != null)     { try { Marshal.ReleaseComObject(dateFmtC2); } catch { } }
                if (dateFmtC1 != null)     { try { Marshal.ReleaseComObject(dateFmtC1); } catch { } }
                if (horasFormatRange != null)    { try { Marshal.ReleaseComObject(horasFormatRange); } catch { } }
                if (horasFmtC2 != null)    { try { Marshal.ReleaseComObject(horasFmtC2); } catch { } }
                if (horasFmtC1 != null)    { try { Marshal.ReleaseComObject(horasFmtC1); } catch { } }
                if (dataRange != null)     { try { Marshal.ReleaseComObject(dataRange); } catch { } }
                if (endCell != null)       { try { Marshal.ReleaseComObject(endCell); } catch { } }
                if (startCell != null)     { try { Marshal.ReleaseComObject(startCell); } catch { } }
                if (fontRow2 != null)      { try { Marshal.ReleaseComObject(fontRow2); } catch { } }
                if (fontRow1 != null)      { try { Marshal.ReleaseComObject(fontRow1); } catch { } }
                if (row2 != null)          { try { Marshal.ReleaseComObject(row2); } catch { } }
                if (row1 != null)          { try { Marshal.ReleaseComObject(row1); } catch { } }
                if (subtotalHeaderCell != null) { try { Marshal.ReleaseComObject(subtotalHeaderCell); } catch { } }
                if (rangeDias != null)     { try { Marshal.ReleaseComObject(rangeDias); } catch { } }
                if (rangeFechas != null)   { try { Marshal.ReleaseComObject(rangeFechas); } catch { } }
                if (tituloFont != null)    { try { Marshal.ReleaseComObject(tituloFont); } catch { } }
                if (rangeTitulo != null)   { try { Marshal.ReleaseComObject(rangeTitulo); } catch { } }
                if (tituloC2 != null)      { try { Marshal.ReleaseComObject(tituloC2); } catch { } }
                if (tituloC1 != null)      { try { Marshal.ReleaseComObject(tituloC1); } catch { } }
                if (allCells != null)      { try { Marshal.ReleaseComObject(allCells); } catch { } }
                if (ws != null)            { try { Marshal.ReleaseComObject(ws); } catch { } }
            }
        }

        private string GetDiaSemanaInicial(DayOfWeek day)
        {
            switch (day)
            {
                case DayOfWeek.Sunday: return "D";
                case DayOfWeek.Monday: return "L";
                case DayOfWeek.Tuesday: return "Ma";
                case DayOfWeek.Wednesday: return "Mi";
                case DayOfWeek.Thursday: return "J";
                case DayOfWeek.Friday: return "V";
                case DayOfWeek.Saturday: return "S";
                default: return null;
            }
        }

        private void AplicarColorColumna(dynamic ws, int colIndex, int filasTotales, int colorBgr)
        {
            dynamic cellTop = null;
            dynamic cellBottom = null;
            dynamic rangeColumn = null;
            dynamic interior = null;
            try
            {
                cellTop = ws.Cells[1, colIndex];
                cellBottom = ws.Cells[3 + filasTotales - 1, colIndex];
                rangeColumn = ws.Range[cellTop, cellBottom];
                interior = rangeColumn.Interior;
                interior.Color = colorBgr;
            }
            catch (Exception ex) when (IsNonCriticalException(ex))
            {
                // Ignorar fallo en columna individual; continuar con las demás columnas
            }
            finally
            {
                if (interior != null)    { try { Marshal.ReleaseComObject(interior); } catch { } }
                if (rangeColumn != null) { try { Marshal.ReleaseComObject(rangeColumn); } catch { } }
                if (cellBottom != null)  { try { Marshal.ReleaseComObject(cellBottom); } catch { } }
                if (cellTop != null)     { try { Marshal.ReleaseComObject(cellTop); } catch { } }
            }
        }

        private static bool IsNonCriticalException(Exception ex)
        {
            return !(ex is OutOfMemoryException || ex is System.Threading.ThreadAbortException);
        }

        // Fase 876: Mapeo de DayOfWeek a nombre canónico en español
        private static string ObtenerNombreDia(DayOfWeek d)
        {
            switch (d)
            {
                case DayOfWeek.Monday: return "LUNES";
                case DayOfWeek.Tuesday: return "MARTES";
                case DayOfWeek.Wednesday: return "MIERCOLES";
                case DayOfWeek.Thursday: return "JUEVES";
                case DayOfWeek.Friday: return "VIERNES";
                case DayOfWeek.Saturday: return "SABADO";
                default: return "DOMINGO";
            }
        }

        // Fase 876: Orden canónico de días para presentación uniforme en columna K
        private static int ObtenerOrdenDia(string d)
        {
            var orden = new[] { "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO" };
            string normalizado = d.Replace("É", "E").Replace("Á", "A");
            int idx = Array.IndexOf(orden, normalizado);
            return idx >= 0 ? idx : 999;
        }

        // Fase 878: Normaliza nombres de días eliminando tildes y diéresis para evitar duplicados en columna K
        private static string NormalizarDia(string d)
        {
            return (d ?? string.Empty).Trim().ToUpperInvariant()
                .Replace("Á", "A").Replace("É", "E").Replace("Í", "I")
                .Replace("Ó", "O").Replace("Ú", "U").Replace("Ü", "U");
        }
    }
}
