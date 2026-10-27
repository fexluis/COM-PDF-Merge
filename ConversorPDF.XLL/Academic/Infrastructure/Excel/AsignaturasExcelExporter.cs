using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.UI.Enums;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Excel
{
    public class AsignaturasExcelExporter
    {
        private readonly dynamic _excelApp;
        private readonly ITrayectoEncoderService _trayectoEncoderService;

        private static readonly IReadOnlyList<(ExcelColumnas Inicio, ExcelColumnas Fin)> ColumnasComisiones = new (ExcelColumnas Inicio, ExcelColumnas Fin)[]
        {
            (ExcelColumnas.FechaInicio1, ExcelColumnas.FechaFin1),
            (ExcelColumnas.FechaInicio2, ExcelColumnas.FechaFin2),
            (ExcelColumnas.FechaInicio3, ExcelColumnas.FechaFin3),
            (ExcelColumnas.FechaInicio4, ExcelColumnas.FechaFin4),
            (ExcelColumnas.FechaInicio5, ExcelColumnas.FechaFin5),
            (ExcelColumnas.FechaInicio6, ExcelColumnas.FechaFin6),
            (ExcelColumnas.FechaInicio7, ExcelColumnas.FechaFin7),
            (ExcelColumnas.FechaInicio8, ExcelColumnas.FechaFin8)
        };

        private const int FILA_MAXIMA_PLANTILLA_DEFAULT = 234;

        private static readonly string[] RangosData = GenerarRangosData();

        private readonly List<string> _ultimosAvisos = new List<string>();

        /// <summary>
        /// Lista de advertencias y avisos generados durante la última exportación (reasignaciones, continuaciones, ventanas).
        /// </summary>
        public IReadOnlyList<string> UltimosAvisos => _ultimosAvisos.AsReadOnly();

        /// <summary>
        /// Número de bloques reasignados por saturación en la última exportación.
        /// </summary>
        public int UltimosBloquesReasignados { get; private set; }

        /// <summary>
        /// Número de filas de continuación generadas en la última exportación.
        /// </summary>
        public int UltimasFilasContinuacion { get; private set; }

        public AsignaturasExcelExporter(dynamic excelApp, ITrayectoEncoderService trayectoEncoderService)
        {
            _excelApp = excelApp;
            _trayectoEncoderService = trayectoEncoderService;
        }

        /// <summary>
        /// Exporta el listado de asignaturas y distribuye los viáticos del docente por vinculación (R1–R7).
        /// Máximo 8 comisiones por fila; las filas de continuación se generan sólo como último recurso.
        /// </summary>
        /// <param name="items">Lista de sesiones filtradas para exportar.</param>
        /// <param name="cacheTerceros">Diccionario de docentes consultado previamente para mejorar el rendimiento.</param>
        public void Exportar(List<SesionListItem> items, Dictionary<string, Tercero> cacheTerceros, ConversorPDF.XLL.UI.ProgressManager progress = null)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            if (cacheTerceros == null)
                throw new ArgumentNullException(nameof(cacheTerceros));

            _ultimosAvisos.Clear();
            UltimosBloquesReasignados = 0;
            UltimasFilasContinuacion = 0;

            dynamic ws = null;
            bool hojaDesprotegida = false; // Fase 734: Solo reproteger si la hoja fue desprotegida por el sistema
            try
            {
                dynamic wb = null;
                try
                {
                    wb = _excelApp.ActiveWorkbook;
                }
                catch { } // Fase 1: Suprimir excepción si no hay libro activo

                string[] requiredSheets = { "NOTOCAR", "PROGRAMACIÓN", "Base asignaturas", "Base docente", "Base gastos terrestre" };
                bool hasAllSheets = true;
                
                if (wb == null)
                {
                    hasAllSheets = false; // Fase 1: Forzar flujo de plantilla si no hay libro
                }
                else
                {
                    foreach (string sheetName in requiredSheets)
                    {
                        bool sheetExists = false;
                        foreach (dynamic sheet in wb.Worksheets)
                        {
                            // Fase 4: Comparación case-insensitive
                            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                sheetExists = true;
                                break;
                            }
                        }
                        if (!sheetExists)
                        {
                            hasAllSheets = false;
                            break;
                        }
                    }
                }

                string newFilePath = null;
                if (!hasAllSheets)
                {
                    string templateFileName = "FV-FO-033.xlsx";
                    string templatesPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTemplates", templateFileName);
                    
                    if (!System.IO.File.Exists(templatesPath))
                        throw new System.IO.FileNotFoundException($"No se encontró la plantilla de programación presupuestal en: {templatesPath}");

                    // Fase 14 y 15: Copiar archivo físico a Mis Documentos, manejando colisiones y locks
                    string docsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string newFileName = $"FV-FO-033-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
                    newFilePath = System.IO.Path.Combine(docsPath, newFileName);

                    try
                    {
                        System.IO.File.Copy(templatesPath, newFilePath, false);
                    }
                    catch (System.IO.IOException)
                    {
                        newFileName = $"FV-FO-033-{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N").Substring(0, 4)}.xlsx";
                        newFilePath = System.IO.Path.Combine(docsPath, newFileName);
                        System.IO.File.Copy(templatesPath, newFilePath, false);
                    }

                    // Fase 16: Eliminar atributo ReadOnly propagado desde la carpeta del instalador
                    System.IO.File.SetAttributes(newFilePath, System.IO.FileAttributes.Normal);

                    // Abrir el archivo copiado
                    wb = _excelApp.Workbooks.Open(newFilePath);
                }

                // Fase 2: Forzar siempre selección de hoja PROGRAMACIÓN, en lugar de confiar en ActiveSheet
                ws = wb.Worksheets["PROGRAMACIÓN"];
                ws.Activate();
                try
                {
                    ws.Unprotect(ConstantsHelper.FORMATO_PASSWORD);
                    hojaDesprotegida = true; // Fase 736: Unprotect exitoso sin excepción => hoja desprotegida
                }
                catch { }
                
                // Fase 8 y 10: Validar que efectivamente la hoja se haya desprotegido, caso contrario es una contraseña diferente
                if (ws.ProtectContents)
                {
                    hojaDesprotegida = false; // Fase 736: Si sigue protegida tras Unprotect, invalidar bandera
                    if (!hasAllSheets) // Fase 10 y 17: flujo plantilla, descartar libro y archivo huérfano
                    {
                        bool prevAlerts = _excelApp.DisplayAlerts;
                        _excelApp.DisplayAlerts = false;
                        try { wb.Close(false); } catch { }
                        _excelApp.DisplayAlerts = prevAlerts;
                        try { if (newFilePath != null) System.IO.File.Delete(newFilePath); } catch { }
                        ws = null; // evitar Protect/Release sobre hoja de libro cerrado en el finally
                    }
                    throw new InvalidOperationException("La hoja 'PROGRAMACIÓN' está protegida con una contraseña distinta a la interna del sistema.");
                }
                dynamic colA = null;
                dynamic lastCell = null;
                try
                {
                    // Buscar en columna A (índice 1 en Excel COM) para no afectar fórmulas de totales en otras columnas
                    colA = ws.Columns[1];
                    const int xlByRows = 1;
                    const int xlPrevious = 2;
                    lastCell = colA.Find("*", System.Reflection.Missing.Value, 
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, 
                        xlByRows, 
                        xlPrevious, 
                        false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                    if (lastCell != null)
                    {
                        int ultimaFila = lastCell.Row;
                        if (ultimaFila >= 14)
                        {
                            int filaLimite = Math.Min(ultimaFila, 10000); // Límite de seguridad
                            foreach (var rango in RangosData)
                            {
                                var partes = rango.Split(':');
                                dynamic colRange = null;
                                try
                                {
                                    colRange = ws.Range[$"{partes[0]}14:{partes[1]}{filaLimite}"];
                                    colRange.ClearContents();
                                }
                                finally
                                {
                                    if (colRange != null) Marshal.ReleaseComObject(colRange);
                                }
                            }
                        }
                    }
                }
                finally
                {
                    if (lastCell != null) Marshal.ReleaseComObject(lastCell);
                    if (colA != null) Marshal.ReleaseComObject(colA);
                }



                using (ExcelOptimizeHelper.Optimize(_excelApp))
                {
                    int filaExcel = 14;

                    var docentes = items
                        .GroupBy(s => s.DocenteNombre)
                        .OrderBy(g => g.Key)
                        .ToList();

                    double progressStep = docentes.Count > 0 ? 100.0 / docentes.Count : 100;
                    int numeroFila = 1; // Movido fuera del bucle para numeración global continua
                    int filaMaximaPlantilla = ObtenerFilaMaximaPlantilla(ws);
                    bool capacidadExcedida = false;
                    int filasOmitidas = 0;
                    string docenteCorte = null;

                    foreach (var grupoDocente in docentes)
                    {
                        // 1. Calcular Viáticos Globales del Docente
                        var sesionesGlobalesDocente = grupoDocente.Where(s => s.Viatica).OrderBy(s => s.FechaSesion).ToList();
                        var bloquesGlobales = ViaticoBloque.AgruparConsecutivos(sesionesGlobalesDocente);

                        // 2. Obtener Asignaturas del Docente incluyendo la ventana de vinculación
                        var asignaturas = grupoDocente
                            .GroupBy(s => new { s.SedeCiudad, s.Subsede, s.AsignaturaNombre, s.Semestre, s.VinculacionFechaInicio, s.VinculacionFechaFin })
                            .OrderBy(g => g.Key.SedeCiudad)
                            .ThenBy(g => g.Key.Subsede )
                            .ThenBy(g => RomanToInt(g.Key.Semestre ?? "0"))
                            .ThenBy(g => g.Key.AsignaturaNombre)
                            .Select(g => (IReadOnlyList<SesionListItem>)g.ToList())
                            .ToList();

                        if (asignaturas.Count == 0) continue;

                        var avisosDocente = new List<string>();
                        var filasPlan = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloquesGlobales, avisosDocente);

                        foreach (var aviso in avisosDocente)
                        {
                            LoggerHelper.Warning($"[AsignaturasExcelExporter] Docente '{grupoDocente.Key}': {aviso}");
                        }

                        bool planCompleto = true;

                        for (int indiceFila = 0; indiceFila < filasPlan.Count; indiceFila++)
                        {
                            if (filaExcel > filaMaximaPlantilla)
                            {
                                capacidadExcedida = true;
                                planCompleto = false;
                                filasOmitidas += filasPlan.Count - indiceFila;
                                docenteCorte = grupoDocente.Key;
                                break;
                            }

                            var filaPlan = filasPlan[indiceFila];
                            bool esFilaDuplicada = filaPlan.EsContinuacion;
                            if (esFilaDuplicada)
                            {
                                UltimasFilasContinuacion++;
                            }

                            var sesiones = filaPlan.Asignatura;
                            var primera = sesiones.First();

                            string cedula = primera.CedulaTercero ?? string.Empty;
                            cacheTerceros.TryGetValue(cedula, out Tercero tercero);

                            var bloquesParaEstaFila = filaPlan.Bloques;
                            bool requiereViatico = bloquesParaEstaFila.Count > 0;

                            if (esFilaDuplicada)
                            {
                                // Las horas (L–O) no se replican para no duplicar HorasAsignadasDocente2/AA–AF.
                                // U, V, Grupo, Prestaciones y Puntaje Ocasionales sí se replican (Fases 762/771).
                                // Nota: Se mantiene el incremento de 'numeroFila' en filas duplicadas 
                                // para garantizar un índice secuencial continuo a lo largo del reporte.
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Num, numeroFila++);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Territorial, primera.SedeDepartamento );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Cetap, BuildCetapName(primera.SedeCiudad, primera.Subsede));
                                EscribirCelda(ws, filaExcel, ExcelColumnas.CodMateria, primera.AsignaturaCodMateria );
                                // Walkthrough: Add Columns E-J to Duplicated Rows in Excel Export
                                EscribirCelda(ws, filaExcel, ExcelColumnas.NombreMateria, primera.AsignaturaNombre );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Creditos, primera.AsignaturaCreditos);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Modalidad, primera.ProgramaTipo );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Programa, primera.ProgramaNombre );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Metodologia, primera.AsignaturaMetodologia );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Nivel, primera.Nivel);
                                // Fin Walkthrough

                                int grupoNum = Enum.TryParse(primera.Grupo, true, out ValoresGrupo parsedGrupo)
                                    ? (int)parsedGrupo
                                    : (int)ValoresGrupo.Ninguno;
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Grupo, grupoNum);

                                EscribirFecha(ws, filaExcel, ExcelColumnas.FechaInicioVinculacion, primera.VinculacionFechaInicio);
                                EscribirFecha(ws, filaExcel, ExcelColumnas.FechaFinVinculacion, primera.VinculacionFechaFin);

                                EscribirCelda(ws, filaExcel, ExcelColumnas.Identificacion, tercero?.Cedula ?? cedula);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.RequiereViaticos, requiereViatico ? "SI" : "NO");

                                // Campos adicionales requeridos en filas duplicadas (Fase 19)
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Vinculacion, tercero?.Tipo );

                                EscribirCelda(ws, filaExcel, ExcelColumnas.Prestaciones, tercero != null ? (tercero.Pensionado == 1 ? "PENSIONADO" : "NO PENSIONADO") : null);

                                if (tercero != null && tercero.PuntajeOcasionales > 0)
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.PuntajeOcasionales, (double)tercero.PuntajeOcasionales);
                                }
                                else
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.PuntajeOcasionales, null);
                                }

                                if (tercero != null && requiereViatico && !string.Equals(tercero.Tipo, "CATEDRA", StringComparison.OrdinalIgnoreCase))
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.SalarioParaCalculoVia, (double)tercero.Salario);
                                }

                                string valorTarifaCatedraFilaDup = requiereViatico ? "Tarifa A" : null;
                                EscribirCelda(ws, filaExcel, ExcelColumnas.TarifaCatedra, valorTarifaCatedraFilaDup);
                            }
                            else
                            {
                                // ── SECCIÓN 1: DATOS DE LA ASIGNATURA ──────────────────
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Num, numeroFila++);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Territorial, primera.SedeDepartamento );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Cetap, BuildCetapName(primera.SedeCiudad, primera.Subsede));
                                EscribirCelda(ws, filaExcel, ExcelColumnas.CodMateria, primera.AsignaturaCodMateria );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.NombreMateria, primera.AsignaturaNombre );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Creditos, primera.AsignaturaCreditos);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Modalidad, primera.ProgramaTipo );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Programa, primera.ProgramaNombre );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Metodologia, primera.AsignaturaMetodologia );
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Nivel, primera.Nivel);

                                int grupoNum = Enum.TryParse(primera.Grupo, true, out ValoresGrupo parsedGrupo)
                                    ? (int)parsedGrupo
                                    : (int)ValoresGrupo.Ninguno;
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Grupo, grupoNum);

                                double totalHoras = Math.Round(sesiones.Sum(s => s.Horas), 1);
                                int totalTutorias = sesiones.Count;
                                int tutoriasPresencial = sesiones.Count(s => s.Viatica);
                                int tutoriasVirtual = sesiones.Count(s => !s.Viatica);

                                EscribirCelda(ws, filaExcel, ExcelColumnas.HorasAsignadasDocente1, totalHoras);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.NumTutorias, totalTutorias);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.NumTutoriasPresenciales, tutoriasPresencial);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.NumTutoriasVirtuales, tutoriasVirtual);

                                // ── SECCIÓN 2: DATOS DEL DOCENTE ───────────────────────
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Identificacion, tercero?.Cedula ?? cedula);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Vinculacion, tercero?.Tipo );

                                EscribirFecha(ws, filaExcel, ExcelColumnas.FechaInicioVinculacion, primera.VinculacionFechaInicio);
                                EscribirFecha(ws, filaExcel, ExcelColumnas.FechaFinVinculacion, primera.VinculacionFechaFin);


                                EscribirCelda(ws, filaExcel, ExcelColumnas.Prestaciones, tercero != null ? (tercero.Pensionado == 1 ? "PENSIONADO" : "NO PENSIONADO") : null);

                                if (tercero != null && tercero.PuntajeOcasionales > 0)
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.PuntajeOcasionales, (double)tercero.PuntajeOcasionales);
                                }
                                else
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.PuntajeOcasionales, null);
                                }

                                // ── SECCIÓN 3: VIÁTICOS ────────────────────────────────
                                EscribirCelda(ws, filaExcel, ExcelColumnas.RequiereViaticos, requiereViatico ? "SI" : "NO");

                                if (tercero != null && requiereViatico && !string.Equals(tercero.Tipo, "CATEDRA", StringComparison.OrdinalIgnoreCase))
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.SalarioParaCalculoVia, (double)tercero.Salario);
                                }

                                string valorTarifaCatedra = requiereViatico ? "Tarifa A" : null;
                                EscribirCelda(ws, filaExcel, ExcelColumnas.TarifaCatedra, valorTarifaCatedra);
                            }

                            // ── SECCIÓN 4: COMISIONES (usar bloquesParaEstaFila ya calculado) ────
                            if (bloquesParaEstaFila.Count > 0)
                            {
                                for (int i = 0; i < bloquesParaEstaFila.Count && i < ColumnasComisiones.Count; i++)
                                {
                                    DateTime? rangoFin = ViaticoBloqueFormatter.ObtenerRangoComision(bloquesParaEstaFila[i])?.Fin;
                                    if (bloquesParaEstaFila[i].FechaMinima.HasValue && rangoFin.HasValue)
                                    {
                                        var (fechaInicioCol, fechaFinCol) = ColumnasComisiones[i];
                                        EscribirCelda(ws, filaExcel, fechaInicioCol, bloquesParaEstaFila[i].FechaMinima.Value.Date);
                                        EscribirCelda(ws, filaExcel, fechaFinCol, rangoFin.Value.Date);
                                    }
                                }

                                // Limpiar comisiones no utilizadas en esta fila para evitar datos residuales
                                for (int i = bloquesParaEstaFila.Count; i < ColumnasComisiones.Count; i++)
                                {
                                    var (fechaInicioCol, fechaFinCol) = ColumnasComisiones[i];
                                    EscribirCelda(ws, filaExcel, fechaInicioCol, null);
                                    EscribirCelda(ws, filaExcel, fechaFinCol, null);
                                }
                            }
                            else
                            {
                                // Limpiar fechas de comisión 1-8 para evitar que datos de la fila anterior persistan
                                foreach (var (fechaInicioCol, fechaFinCol) in ColumnasComisiones)
                                {
                                    EscribirCelda(ws, filaExcel, fechaInicioCol, null);
                                    EscribirCelda(ws, filaExcel, fechaFinCol, null);
                                }
                            }

                            // ── SECCIÓN 5: DATOS DE RUTA ───────────────────────────
                            if (bloquesParaEstaFila.Count > 0)
                            {
                                string residencia = tercero?.Residencia ;

                                var sedesOrdenadas = bloquesParaEstaFila
                                    .SelectMany(b => b.Sesiones)
                                    .Where(s => !string.IsNullOrWhiteSpace(s.SedeCiudad) || !string.IsNullOrWhiteSpace(s.Subsede))
                                    .OrderBy(s => s.FechaSesion)
                                    .ThenBy(s => string.IsNullOrWhiteSpace(s.SedeCiudad) ? s.Subsede : s.SedeCiudad)
                                    // Para la ruta: priorizar SedeCiudad; si está en blanco, usar Subsede.
                                    .Select(s => (string.IsNullOrWhiteSpace(s.SedeCiudad) ? s.Subsede : s.SedeCiudad).ToUpper())
                                    .ToList();

                                var rutaCompleta = new List<string>();
                                if (!string.IsNullOrWhiteSpace(residencia))
                                {
                                    rutaCompleta.Add(residencia.ToUpper());
                                }
                                rutaCompleta.AddRange(sedesOrdenadas);

                                var rutaLimpia = new List<string>();
                                for (int i = 0; i < rutaCompleta.Count; i++)
                                {
                                    if (i == 0 || rutaCompleta[i] != rutaCompleta[i - 1])
                                        rutaLimpia.Add(rutaCompleta[i]);
                                }

                                string origen = !string.IsNullOrWhiteSpace(residencia) ? residencia.ToUpper() : null;
                                string destino = rutaLimpia.LastOrDefault() ;

                                string escala = null;
                                // IMPORTANTE: Los nombres de ciudades no deben contener guiones ("-") ya que se utilizan 
                                // como delimitador de segmentos en el cálculo de la escala.
                                if (rutaLimpia.Count > 2)
                                {
                                    escala = string.Join("-", rutaLimpia.Skip(1).Take(rutaLimpia.Count - 2));
                                }

                                EscribirCelda(ws, filaExcel, ExcelColumnas.Origen, origen);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Escala, escala);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Destino, destino);

                                var segmentos = new List<string>();
                                if (!string.IsNullOrWhiteSpace(residencia)) segmentos.Add(residencia.ToUpper());

                                if (!string.IsNullOrWhiteSpace(escala))
                                {
                                    segmentos.AddRange(escala.Split('-').Select(s => s.Trim().ToUpper()).Where(s => !string.IsNullOrWhiteSpace(s)));
                                }

                                if (!string.IsNullOrWhiteSpace(destino)) segmentos.Add(destino);
                                if (!string.IsNullOrWhiteSpace(residencia) && !string.Equals(destino, residencia, StringComparison.OrdinalIgnoreCase)) segmentos.Add(residencia.ToUpper());

                                string verificacion = segmentos.Count >= 2 
                                    ? string.Join(" - ", segmentos) 
                                    : null;
                                EscribirCelda(ws, filaExcel, ExcelColumnas.VerificacionRuta, verificacion);
                                
                                var ultimaSesionBloque = bloquesParaEstaFila
                                    .SelectMany(b => b.Sesiones)
                                    .OrderBy(s => s.FechaSesion)
                                    .ThenBy(s => s.SedeCiudad)
                                    .LastOrDefault();

                                // Fase 19: Gastos terrestres debe imprimirse siempre, incluso en filas duplicadas
                                EscribirCelda(ws, filaExcel, ExcelColumnas.GastosTerrestres, 
                                    ultimaSesionBloque?.TransporteSede.HasValue == true 
                                        ? (object)(double)ultimaSesionBloque.TransporteSede.Value 
                                        : 0);

                                if (!string.IsNullOrEmpty(verificacion) && _trayectoEncoderService != null)
                                {
                                    string territorial = primera.SedeDepartamento ;
                                    string codRuta = _trayectoEncoderService.GetOrCreateCodigoTrayecto(territorial, verificacion, ultimaSesionBloque?.TransporteSede);
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.CodRuta, codRuta);
                                }
                                else
                                {
                                    EscribirCelda(ws, filaExcel, ExcelColumnas.CodRuta, null);
                                }
                                
                                var sesionAereaBloque = bloquesParaEstaFila
                                    .SelectMany(b => b.Sesiones)
                                    .FirstOrDefault(s => s.IsTerminalAereo == 1)
                                    ?? bloquesParaEstaFila.SelectMany(b => b.Sesiones).FirstOrDefault();

                                var (_, _, tiquete, aeroTerminal) = ViaticosHelper.ResolverAereo(
                                    sesionAereaBloque?.IsTerminalAereo ?? 0,
                                    sesionAereaBloque?.TerminalAereaNombre,
                                    docente: false);

                                EscribirCelda(ws, filaExcel, ExcelColumnas.RequiereTiquetesAereos, tiquete);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.TerminalAerea, 
                                    tiquete == "SI" ? aeroTerminal : ViaticosHelper.SIN_TERMINAL);
                            }
                            else
                            {
                                // Limpiar columnas de ruta para evitar que datos de la fila anterior persistan
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Origen, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Escala, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.Destino, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.VerificacionRuta, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.CodRuta, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.GastosTerrestres, null);
                                EscribirCelda(ws, filaExcel, ExcelColumnas.RequiereTiquetesAereos, "NO");
                                EscribirCelda(ws, filaExcel, ExcelColumnas.TerminalAerea, ViaticosHelper.SIN_TERMINAL);
                            }

                            filaExcel++;
                        }

                        if (planCompleto)
                        {
                            foreach (var aviso in avisosDocente)
                            {
                                _ultimosAvisos.Add(aviso);
                            }
                            UltimosBloquesReasignados += avisosDocente.Count(a => a.IndexOf("saturación", StringComparison.OrdinalIgnoreCase) >= 0);
                        }

                        if (capacidadExcedida) break;
                        progress?.AddProgress(progressStep);
                    }

                    if (capacidadExcedida)
                    {
                        string aviso = $"Capacidad de T_Principal8 excedida (fila {filaMaximaPlantilla}). " +
                                       $"Filas omitidas del docente '{docenteCorte}': {filasOmitidas} " +
                                       $"(docentes posteriores no procesados).";
                        LoggerHelper.Error($"[AsignaturasExcelExporter] {aviso}");
                        _ultimosAvisos.Add(aviso);
                    }
                }
            }
            finally
            {
                // Fase 5: Mover protección a finally e instanciar release de COM
                // Fase 728: Proteger hoja permitiendo explícitamente "Usar Autofiltro" (AllowFiltering = true)
                // Fase 730 / 735: Fallback a ws.Protect con trazabilidad en LoggerHelper
                // Fase 734: Solo reproteger si el sistema efectivamente desprotegió la hoja
                if (ws != null)
                {
                    if (hojaDesprotegida)
                    {
                        try
                        {
                            ProtegerHojaPermitiendoAutofiltro(ws, ConstantsHelper.FORMATO_PASSWORD);
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[AsignaturasExcelExporter] Protect por reflexión falló; aplicando fallback: {ex.Message}");
                            try
                            {
                                ws.Protect(ConstantsHelper.FORMATO_PASSWORD);
                            }
                            catch (Exception exFallback)
                            {
                                LoggerHelper.Error("[AsignaturasExcelExporter] No fue posible reproteger la hoja 'PROGRAMACIÓN'.", exFallback);
                            }
                        }
                    }
                    Marshal.ReleaseComObject(ws);
                }
            }
        }

        /// <summary>
        /// Protege la hoja conservando los valores por defecto de Excel, habilitando únicamente
        /// la casilla "Usar Autofiltro" (AllowFiltering) del cuadro "Proteger hoja".
        /// Nota (Fase 733): AllowSorting permanece deshabilitado dado que Excel exige celdas desbloqueadas
        /// para permitir ordenamiento en hojas protegidas, condición no aplicable a esta plantilla.
        /// Firma posicional de Worksheet.Protect (16 parámetros opcionales en COM):
        /// 0: Password, 1: DrawingObjects, 2: Contents, 3: Scenarios, 4: UserInterfaceOnly,
        /// 5: AllowFormattingCells, 6: AllowFormattingColumns, 7: AllowFormattingRows,
        /// 8: AllowInsertingColumns, 9: AllowInsertingRows, 10: AllowInsertingHyperlinks,
        /// 11: AllowDeletingColumns, 12: AllowDeletingRows, 13: AllowSorting,
        /// 14: AllowFiltering, 15: AllowUsingPivotTables.
        /// </summary>
        /// <param name="ws">Instancia COM dynamic de la hoja de cálculo (Worksheet).</param>
        /// <param name="password">Contraseña interna para la protección del formato.</param>
        private static void ProtegerHojaPermitiendoAutofiltro(dynamic ws, string password)
        {
            var args = new object[16];
            for (int i = 0; i < args.Length; i++)
            {
                args[i] = Missing.Value;
            }

            args[0] = password; // Conserva contraseña interna del formato
            // args[13] = true; // REVERTIDO (Fase 733): inefectivo con celdas bloqueadas en plantilla protegida
            args[14] = true;    // AllowFiltering -> "Usar Autofiltro" (Fase 728)

            ws.GetType().InvokeMember(
                "Protect",
                BindingFlags.InvokeMethod,
                null,
                ws,
                args);
        }

        private static void EscribirCelda(dynamic ws, int fila, ExcelColumnas col, object valor)
        {
            dynamic cell = null;
            try
            {
                cell = ws.Cells[fila, (int)col + 1];
                if (valor == null)
                {
                    var t = cell.GetType();
                    t.InvokeMember("ClearContents", System.Reflection.BindingFlags.InvokeMethod, null, cell, new object[0]);
                }
                else
                {
                    cell.Value = valor;
                }
            }
            finally
            {
                if (cell != null) Marshal.ReleaseComObject(cell);
            }
        }

        private static void EscribirFecha(dynamic ws, int fila, ExcelColumnas col, string fechaStr)
        {
            if (DateTime.TryParse(fechaStr, out DateTime dt))
            {
                EscribirCelda(ws, fila, col, dt);
            }
            else
            {
                EscribirCelda(ws, fila, col, null);
            }
        } // Los métodos AgruparEnBloques y ExtraerCuatroComisiones han sido eliminados a favor de ViaticoBloque.AgruparConsecutivos

        private static string BuildCetapName(string sedeCiudad, string subsede)
        {
            if (!string.IsNullOrWhiteSpace(subsede))
            {
                return string.IsNullOrWhiteSpace(sedeCiudad)
                    ? subsede
                    : $"{sedeCiudad} ({subsede})";
            }
            return sedeCiudad ?? string.Empty;
        }

        private static int RomanToInt(string roman)
        {
            if (string.IsNullOrEmpty(roman)) return 0;
            var map = new Dictionary<char, int> { { 'I', 1 }, { 'V', 5 }, { 'X', 10 } };
            int total = 0, prev = 0;
            for (int i = roman.Length - 1; i >= 0; i--)
            {
                if (!map.TryGetValue(roman[i], out int curr)) return 0;
                total += curr < prev ? -curr : curr;
                prev = curr;
            }
            return total;
        }

        private static string[] GenerarRangosData()
        {
            // NOTA DE DISEÑO: Los rangos definidos omiten intencionalmente las columnas que
            // en la plantilla de Excel corresponden a fórmulas automáticas (ej. NombreDocente,
            // TarifaPlanta, TotalViatico, etc.). Borrar dichas columnas destruiría los cálculos
            // de la hoja de cálculo. Por lo tanto, no se incluyen en esta lista dinámica.
            var rangos = new List<string>
            {
                $"{GetColumnLetter(ExcelColumnas.Num)}:{GetColumnLetter(ExcelColumnas.NumTutoriasVirtuales)}",
                $"{GetColumnLetter(ExcelColumnas.Identificacion)}:{GetColumnLetter(ExcelColumnas.Identificacion)}",
                // Se separa el rango para excluir la columna T (HorasAsignadasDocente2) que contiene una fórmula
                $"{GetColumnLetter(ExcelColumnas.Vinculacion)}:{GetColumnLetter(ExcelColumnas.Vinculacion)}",
                $"{GetColumnLetter(ExcelColumnas.FechaInicioVinculacion)}:{GetColumnLetter(ExcelColumnas.FechaFinVinculacion)}",
                $"{GetColumnLetter(ExcelColumnas.Prestaciones)}:{GetColumnLetter(ExcelColumnas.Prestaciones)}",
                $"{GetColumnLetter(ExcelColumnas.PuntajeOcasionales)}:{GetColumnLetter(ExcelColumnas.PuntajeOcasionales)}",
                $"{GetColumnLetter(ExcelColumnas.RequiereViaticos)}:{GetColumnLetter(ExcelColumnas.SalarioParaCalculoVia)}",
                $"{GetColumnLetter(ExcelColumnas.TarifaCatedra)}:{GetColumnLetter(ExcelColumnas.TarifaCatedra)}"
            };

            foreach (var comision in ColumnasComisiones)
            {
                rangos.Add($"{GetColumnLetter(comision.Inicio)}:{GetColumnLetter(comision.Fin)}");
            }

            rangos.Add($"{GetColumnLetter(ExcelColumnas.Origen)}:{GetColumnLetter(ExcelColumnas.TerminalAerea)}");

            return rangos.ToArray();
        }

        private static string GetColumnLetter(ExcelColumnas col)
        {
            int index = (int)col + 1;
            if (index <= 0) throw new ArgumentOutOfRangeException(nameof(col), "El índice de columna debe ser mayor que cero.");
            
            string letters = "";
            while (index > 0)
            {
                index--;
                letters = (char)('A' + (index % 26)) + letters;
                index /= 26;
            }
            return letters;
        }

        private static int ObtenerFilaMaximaPlantilla(dynamic ws)
        {
            dynamic coleccion = null;
            dynamic tabla = null;
            dynamic rango = null;
            dynamic filas = null;
            try
            {
                coleccion = ws.ListObjects;
                tabla = coleccion["T_Principal8"];
                rango = tabla.Range;
                filas = rango.Rows;
                int maxima = (int)rango.Row + (int)filas.Count - 1;
                return maxima > 13 ? maxima : FILA_MAXIMA_PLANTILLA_DEFAULT;
            }
            catch
            {
                return FILA_MAXIMA_PLANTILLA_DEFAULT;
            }
            finally
            {
                if (filas != null) { try { Marshal.ReleaseComObject(filas); } catch { } }
                if (rango != null) { try { Marshal.ReleaseComObject(rango); } catch { } }
                if (tabla != null) { try { Marshal.ReleaseComObject(tabla); } catch { } }
                if (coleccion != null) { try { Marshal.ReleaseComObject(coleccion); } catch { } }
            }
        }
    }
}

