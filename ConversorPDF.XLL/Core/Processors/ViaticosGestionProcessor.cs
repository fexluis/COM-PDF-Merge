using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Processors
{
    /// <summary>
    /// Procesador encargado de la gestión documental de viáticos.
    /// Realiza tareas de organización de archivos, renombrado automático mediante OCR/Metadatos y unificación de expedientes.
    /// </summary>
    internal class ViaticosGestionProcessor : IViaticosGestionProcessor
    {
        private readonly IExcelRepository _excelRepository;
        private readonly IViaticosRepository _viaticosRepository;
        private readonly IPdfToTexto _pdfToTexto;
        private readonly IPdfMetadataService _metadataService;

        /// <summary>
        /// Inicializa una nueva instancia de la clase <see cref="ViaticosGestionProcessor"/>.
        /// </summary>
        /// <param name="excelRepository">Acceso a Excel.</param>
        /// <param name="viaticosRepository">Lógica de procesos de viáticos.</param>
        /// <param name="PdfToTextoHelper">Servicio de extracción de texto de PDFs (OCR).</param>
        /// <param name="metadataService">Servicio de gestión de metadatos PDF.</param>
        public ViaticosGestionProcessor(IExcelRepository excelRepository, IViaticosRepository viaticosRepository, IPdfToTexto PdfToTextoHelper, IPdfMetadataService metadataService)
        {
            _excelRepository = excelRepository;
            _viaticosRepository = viaticosRepository;
            _pdfToTexto = PdfToTextoHelper;
            _metadataService = metadataService;
        }

        /// <inheritdoc/>
        public int RenombrarSolicitudes(string folder, int inicio = 0, int fin = 0)
        {
            ResultadoRenombrado resultado = RenombrarSolicitudesConDetalle(folder, inicio, fin);
            return resultado.HayError ? -1 : resultado.Renombrados;
        }

        /// <inheritdoc/>
        public ResultadoRenombrado RenombrarSolicitudesConDetalle(string folder, int inicio = 0, int fin = 0)
        {
            var r = new ResultadoRenombrado { CarpetaRaiz = folder };
            try
            {
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    r.Error = $"La carpeta '{folder}' no existe.";
                    LoggerHelper.Error("RenombrarSolicitudes: " + r.Error);
                    return r;
                }

                // Comisiones de CONSULTAS filtradas por el rango solicitado
                var solicitudes = GetSolicitudes(inicio, fin);
                r.ComisionesEnRango = solicitudes.Count;
                if (solicitudes.Count == 0)
                {
                    LoggerHelper.Warning($"RenombrarSolicitudes: no se encontraron comisiones en CONSULTAS para el rango {inicio}-{fin}.");
                    return r;
                }

                // Raíz del libro + subcarpeta estándar de solicitudes
                foreach (string carpeta in CarpetasDeSolicitudes(folder))
                {
                    r.CarpetasProcesadas.Add(carpeta);
                    Estandarizar(carpeta);

                    foreach (string file in Directory.GetFiles(carpeta, "*.pdf"))
                    {
                        r.PdfsExaminados++;
                        string motivo = RenombrarFila(file, solicitudes, carpeta);
                        string carpetaClave = Path.GetFileName(carpeta.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                        string clave = carpetaClave + "/" + Path.GetFileName(file);
                        r.Detalle[clave] = motivo;
                        Contabilizar(r, motivo);
                        LoggerHelper.Debug($"RenombrarSolicitudes: '{Path.GetFileName(file)}' -> {motivo}");
                    }
                }

                LoggerHelper.Resumen("RENOMBRAR",
                    ("carpetas", r.CarpetasProcesadas.Count),
                    ("comisionesEnRango", r.ComisionesEnRango),
                    ("pdfs", r.PdfsExaminados),
                    ("renombrados", r.Renombrados),
                    ("yaCorrectos", r.YaCorrectos),
                    ("colisiones", r.Colisiones),
                    ("duplicados", r.Duplicados),
                    ("ambiguos", r.Ambiguos),
                    ("sinComision", r.SinComision),
                    ("sinDatos", r.SinDatos),
                    ("prefijoInvalido", r.PrefijoInvalido),
                    ("erroresIo", r.ErroresIo));

                return r;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("RenombrarSolicitudes", ex);
                r.Error = ex.Message;
                return r;
            }
        }

        /// <inheritdoc/>
        public int UnificarSolicitudes(string folder, int inicio = 0, int fin = 0)
        {
            try
            {
                if (!Directory.Exists(folder)) return -1;

                // Asegura que los archivos en la carpeta raíz sigan el estándar de nombres
                Estandarizar(folder);

                // Define las rutas de las subcarpetas estándar del sistema
                string pRP = Path.Combine(folder, App.FolderNames.RPS);
                string pCDP = Path.Combine(folder, App.FolderNames.CDPS);
                string pAuto = Path.Combine(folder, App.FolderNames.AUTORIZADOS);
                string pVeri = Path.Combine(folder, App.FolderNames.VERIFICADOS);
                string pUnif = Path.Combine(folder, App.FolderNames.UNIFICADOS);
                string pSoli = Path.Combine(folder, App.FolderNames.SOLICITUDES);

                if (!Directory.Exists(pUnif)) Directory.CreateDirectory(pUnif);

                // Itera sobre todas las hojas del libro (excepto las técnicas) para unificar sus expedientes
                dynamic wb = _excelRepository.GetActiveWorkbook();
                int total = wb.Sheets.Count;
                int count = 0;

                for (int i = 1; i <= total; i++)
                {
                    dynamic ws = wb.Sheets[i];
                    string n = ws.Name;

                    // Omite hojas de configuración o masivas
                    if (n == "Masivo" || n == "CONSULTAS" || n == "Programación" || n == "Config" || n == "Datos" || n == "Dependencias") continue;

                    // Lógica de unificación por hoja aplicando filtro de rango
                    if (Unificar(ws, pSoli, pAuto, pRP, pCDP, pVeri, pUnif, inicio, fin)) count++;
                }

                // Abre la carpeta de resultados si hubo éxito
                if (count > 0) System.Diagnostics.Process.Start("explorer.exe", pUnif);
                return count;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("UnificarSolicitudes", ex);
                return -1;
            }
        }

        /// <summary>
        /// Obtiene la información básica de comisiones (Número, CC, Nombre) desde la hoja CONSULTAS, filtrando por rango de consecutivo/año.
        /// </summary>
        /// <param name="inicio">Número de viático inicial (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final (0 para sin filtro).</param>
        /// <returns>Lista de solicitudes que cumplen el filtro.</returns>
        internal List<SInfo> GetSolicitudes(int inicio = 0, int fin = 0)
        {
            var list = new List<SInfo>();
            object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
            if (data == null) return list;

            int rows = data.GetLength(0);
            int total = 0;
            int sinCedula = 0;
            for (int i = 2; i <= rows; i++)
            {
                string num = (data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(num)) continue;

                total++;
                if (!ViaticoNumeroComision.EnRango(num, inicio, fin))
                {
                    continue;
                }

                string cc = (data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString() ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(cc))
                {
                    sinCedula++;
                    continue;
                }

                list.Add(new SInfo
                {
                    Num = num,
                    CC = cc,
                    Nom = Clean(data[i, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO]?.ToString() ?? string.Empty)
                });
            }

            if (sinCedula > 0)
            {
                LoggerHelper.Warning($"CONSULTAS: {sinCedula} fila(s) sin cédula omitida(s).");
            }

            LoggerHelper.Resumen("CONSULTAS", ("filas", total), ("enRango", list.Count), ("sinCedula", sinCedula),
                                 ("rango", ViaticoNumeroComision.HayFiltro(inicio, fin) ? $"{inicio}-{fin}" : "sin filtro"));
            return list;
        }

        /// <summary>
        /// Limpia el nombre del funcionario para usarlo en nombres de archivos (PascalCase sin espacios).
        /// Trunca a un máximo de 18 caracteres basándose en la longitud de la cadena procesada para evitar ArgumentOutOfRangeException.
        /// </summary>
        /// <param name="s">Nombre original con espacios o caracteres diversos.</param>
        /// <returns>Nombre limpio y truncado a un máximo de 18 caracteres.</returns>
        private string Clean(string s)
        {
            string limpio = Regex.Replace(App.CapitalizarPalabras(s ?? string.Empty), @"\s+", "");
            return limpio.Substring(0, Math.Min(limpio.Length, 18));
        }

        /// <summary>
        /// Enumera la carpeta raíz y, si existe, la subcarpeta estándar de solicitudes.
        /// Evita duplicar la raíz cuando ya es la carpeta de solicitudes.
        /// </summary>
        private static IEnumerable<string> CarpetasDeSolicitudes(string folder)
        {
            yield return folder;

            string pSoli = Path.Combine(folder, App.FolderNames.SOLICITUDES);
            if (!string.Equals(pSoli, folder, StringComparison.OrdinalIgnoreCase) && Directory.Exists(pSoli))
            {
                yield return pSoli;
            }
        }

        /// <summary> Incrementa el contador correspondiente al motivo devuelto por <see cref="RenombrarFila"/>. </summary>
        private static void Contabilizar(ResultadoRenombrado r, string motivo)
        {
            switch (motivo)
            {
                case MotivoRenombrado.Renombrado: r.Renombrados++; break;
                case MotivoRenombrado.YaCorrecto: r.YaCorrectos++; break;
                case MotivoRenombrado.Colision: r.Colisiones++; break;
                case MotivoRenombrado.Duplicado: r.Duplicados++; break;
                case MotivoRenombrado.Ambigua: r.Ambiguos++; break;
                case MotivoRenombrado.SinComision: r.SinComision++; break;
                case MotivoRenombrado.SinDatos: r.SinDatos++; break;
                case MotivoRenombrado.PrefijoInvalido: r.PrefijoInvalido++; break;
                default: r.ErroresIo++; break;
            }
        }

        /// <summary>
        /// Normaliza los nombres de todos los archivos en un directorio reemplazando espacios y guiones bajos por un guion único.
        /// Previene problemas de búsqueda por discrepancias mínimas en el nombre.
        /// </summary>
        /// <param name="f">Ruta del directorio a estandarizar.</param>
        private void Estandarizar(string f)
        {
            foreach (string file in Directory.GetFiles(f, "*.pdf"))
            {
                string old = Path.GetFileName(file);
                string newN = Regex.Replace(old.Replace("_", "-").Replace(" ", "-"), @"-+", "-");
                if (old != newN)
                {
                    try
                    {
                        File.Move(file, Path.Combine(f, newN));
                        LoggerHelper.Debug($"Estandarizar: '{old}' -> '{newN}'");
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Warning($"Estandarizar: '{old}' no normalizado: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Intenta identificar y renombrar un PDF aplicando la política de desambiguación:
        /// N0 candidato único por (Num, Cédula) → N1 metadatos → N2 OCR → N3 sin renombrar.
        /// </summary>
        /// <returns>Motivo del resultado (ver <see cref="MotivoRenombrado"/>).</returns>
        private string RenombrarFila(string file, List<SInfo> sols, string carpeta)
        {
            string oldN = Path.GetFileNameWithoutExtension(file);
            string ext = Path.GetExtension(file);

            // El nombre debe iniciar con el número de comisión (admite '-' o '_' como separador)
            Match prefijo = Regex.Match(oldN, @"^(\d+)");
            if (!prefijo.Success) return MotivoRenombrado.PrefijoInvalido;
            string num = prefijo.Groups[1].Value;

            // Fallback para prefijos de 6 dígitos con dígito espurio (ej. 202426 -> 20426).
            // Sólo se acepta si la comisión resultante existe, es única y el nombre del archivo coincide.
            if (num.Length == 6 && !sols.Any(s => s.Num == num))
            {
                for (int pos = 1; pos <= 3; pos++)
                {
                    string candidatoNum = num.Remove(pos, 1);
                    var posibles = sols.Where(s => s.Num == candidatoNum).ToList();
                    if (posibles.Count == 1 && CoincideNombre(oldN, posibles[0].Nom))
                    {
                        LoggerHelper.Debug($"RenombrarSolicitudes: prefijo '{num}' remapeado a '{candidatoNum}' por coincidencia de nombre.");
                        num = candidatoNum;
                        break;
                    }
                }
            }

            // N0: un solo candidato por (Num, Cédula); las filas duplicadas idénticas colapsan aquí
            var candidatos = sols.Where(s => s.Num == num)
                                 .GroupBy(s => NormalizarDocumento(s.CC))
                                 .Select(g => g.First())
                                 .ToList();
            if (candidatos.Count == 0) return MotivoRenombrado.SinComision;

            var completos = candidatos.Where(c => !string.IsNullOrWhiteSpace(c.CC) && !string.IsNullOrWhiteSpace(c.Nom)).ToList();
            if (completos.Count == 0) return MotivoRenombrado.SinDatos;
            candidatos = completos;

            SInfo elegido = candidatos.Count == 1 ? candidatos[0] : null;

            // N1: metadatos del propio PDF
            if (elegido == null) elegido = IdentificarPorMetadatos(file, candidatos);

            // N2: OCR de la primera página validado contra las cédulas candidatas
            if (elegido == null) elegido = IdentificarPorOcr(file, candidatos);

            // N3: nada se renombra a ciegas
            if (elegido == null) return MotivoRenombrado.Ambigua;

            string nuevoNombre = $"{elegido.Num}-{elegido.CC}-{elegido.Nom}{ext}";
            if (string.Equals(oldN + ext, nuevoNombre, StringComparison.OrdinalIgnoreCase))
                return MotivoRenombrado.YaCorrecto;

            string destino = Path.Combine(carpeta, nuevoNombre);
            if (File.Exists(destino))
            {
                if (!MismoContenido(file, destino)) return MotivoRenombrado.Colision;

                // Duplicado exacto: se archiva en Solicitudes\_Duplicados conservando su nombre.
                string pDup = Path.Combine(carpeta, App.FolderNames.DUPLICADOS);
                Directory.CreateDirectory(pDup);
                string dupDestino = Path.Combine(pDup, Path.GetFileName(file));
                if (File.Exists(dupDestino))
                {
                    dupDestino = Path.Combine(pDup, oldN + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ext);
                }

                try
                {
                    File.Move(file, dupDestino);
                    return MotivoRenombrado.Duplicado;
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"RenombrarSolicitudes: no se pudo archivar duplicado '{Path.GetFileName(file)}': {ex.Message}");
                    return MotivoRenombrado.ErrorIo;
                }
            }

            try
            {
                File.Move(file, destino);
                return MotivoRenombrado.Renombrado;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"RenombrarSolicitudes: no se pudo renombrar '{Path.GetFileName(file)}': {ex.Message}");
                return MotivoRenombrado.ErrorIo;
            }
        }

        /// <summary> N1: identifica al funcionario por los metadatos del PDF (cédula en PalabrasClave). </summary>
        private SInfo IdentificarPorMetadatos(string file, List<SInfo> candidatos)
        {
            try
            {
                var meta = _metadataService.LeerMetadatos(file);
                if (meta == null || meta.MetaDocumento != TipoMetadata.VIA_SOLICITUD || string.IsNullOrWhiteSpace(meta.PalabrasClave))
                    return null;

                string doc = NormalizarDocumento(meta.PalabrasClave.Split(',')[0]);
                return candidatos.FirstOrDefault(s => NormalizarDocumento(s.CC) == doc);
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"RenombrarSolicitudes: metadatos ilegibles en '{Path.GetFileName(file)}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// N2: identifica al funcionario por la cédula leída del PDF.
        /// </summary>
        private SInfo IdentificarPorOcr(string file, List<SInfo> candidatos)
        {
            var cedulas = new HashSet<string>(candidatos.Select(s => NormalizarDocumento(s.CC)));
            string doc = FindCC(file, cedulas);
            if (string.IsNullOrEmpty(doc)) return null;
            return candidatos.FirstOrDefault(s => NormalizarDocumento(s.CC) == doc);
        }

        /// <summary> Quita separadores de miles y espacios de un número de documento. </summary>
        internal static string NormalizarDocumento(string valor)
        {
            return Regex.Replace(valor ?? string.Empty, @"[\s\.,]", string.Empty);
        }

        /// <summary>
        /// Determina si al menos un token significativo del nombre de CONSULTAS (5+ letras)
        /// aparece en el nombre del archivo, sin distinguir mayúsculas ni acentos.
        /// </summary>
        private static bool CoincideNombre(string nombreArchivo, string nombreConsultas)
        {
            string limpio = SinAcentos(nombreConsultas);
            foreach (string token in Regex.Split(SinAcentos(nombreArchivo), "[^A-Z0-9]+"))
            {
                if (token.Length < 5) continue;

                // Coincidencia tolerante: se acepta un prefijo de 5+ letras (erratas tipo ILDEBRANDO/ILDELBRANDO).
                for (int len = token.Length; len >= 5; len--)
                {
                    if (limpio.Contains(token.Substring(0, len))) return true;
                }
            }
            return false;
        }

        /// <summary> Pasa a mayúsculas y elimina diacríticos para comparaciones tolerantes. </summary>
        private static string SinAcentos(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            string norm = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder(norm.Length);
            foreach (char c in norm)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            return sb.ToString().ToUpperInvariant();
        }

        /// <summary> Compara dos archivos por tamaño y contenido byte a byte. </summary>
        private static bool MismoContenido(string a, string b)
        {
            var fa = new FileInfo(a);
            var fb = new FileInfo(b);
            if (fa.Length != fb.Length) return false;

            using (FileStream sa = File.OpenRead(a))
            using (FileStream sb = File.OpenRead(b))
            {
                int x;
                int y;
                do
                {
                    x = sa.ReadByte();
                    y = sb.ReadByte();
                    if (x != y) return false;
                }
                while (x != -1);
            }
            return true;
        }

        /// <inheritdoc/>
        public RangoSolicitudes ObtenerRangoSolicitudes()
        {
            var rango = new RangoSolicitudes();
            object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
            if (data == null) return rango;

            for (int i = 2; i <= data.GetLength(0); i++)
            {
                string num = Texto(data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]);
                if (!int.TryParse(num, out int n) || n <= 0) continue;

                if (!rango.HayDatos)
                {
                    rango.Inicio = n;
                    rango.HayDatos = true;
                }
                rango.Fin = n;
            }

            if (rango.HayDatos && rango.Inicio > rango.Fin)
            {
                int tmp = rango.Inicio;
                rango.Inicio = rango.Fin;
                rango.Fin = tmp;
            }

            return rango;
        }

        /// <summary>
        /// Extrae de la primera página una cédula que pertenezca al conjunto de cédulas candidatas.
        /// Patrón 1: cédula seguida de correo (formato GF-FO-023). Patrón 2: cualquier secuencia
        /// numérica que coincida exactamente con una cédula candidata (evita tomar el rubro presupuestal).
        /// </summary>
        /// <param name="f">Ruta del PDF.</param>
        /// <param name="cedulasCandidatas">Cédulas normalizadas de las filas de CONSULTAS para esa comisión.</param>
        /// <returns>Cédula normalizada encontrada, o cadena vacía.</returns>
        private string FindCC(string f, ICollection<string> cedulasCandidatas)
        {
            if (cedulasCandidatas == null || cedulasCandidatas.Count == 0) return string.Empty;

            string t = _pdfToTexto.ExtraerTextoPagina(f, 1);
            if (string.IsNullOrWhiteSpace(t)) return string.Empty;

            foreach (Match m in Regex.Matches(t, @"(\d[\d\.,\s]{5,})\s+[a-zA-Z0-9._%+-]+@"))
            {
                string candidata = NormalizarDocumento(m.Groups[1].Value);
                if (cedulasCandidatas.Contains(candidata)) return candidata;
            }

            var halladas = Regex.Matches(t, @"(?<!\d)\d[\d\.,\s]{4,18}\d(?!\d)")
                                .Cast<Match>()
                                .Select(m => NormalizarDocumento(m.Value))
                                .Where(c => c.Length >= 6 && c.Length <= 11 && cedulasCandidatas.Contains(c))
                                .Distinct()
                                .ToList();

            return halladas.Count == 1 ? halladas[0] : string.Empty;
        }

        /// <inheritdoc/>
        public ResultadoVerificacionSolicitudes VerificarSolicitudesPresentes(string folder, int inicio, int fin)
        {
            var r = new ResultadoVerificacionSolicitudes();
            try
            {
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    r.Error = $"La carpeta '{folder}' no existe.";
                    return r;
                }

                object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
                if (data == null)
                {
                    r.Error = "No se pudo leer la hoja CONSULTAS.";
                    return r;
                }

                dynamic ws = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS);
                if (ws == null)
                {
                    r.Error = "No se encontró la hoja CONSULTAS en el libro activo.";
                    return r;
                }

                HashSet<string> indice = ConstruirIndiceArchivos(folder);

                _excelRepository.ScreenUpdating(false);
                try
                {
                    for (int i = 2; i <= data.GetLength(0); i++)
                    {
                        string num = Texto(data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]);
                        if (num.Length == 0) continue;
                        if (!ViaticoNumeroComision.EnRango(num, inicio, fin)) continue;

                        r.FilasRevisadas++;

                        string estado = Norm(Texto(data[i, ConstantsHelper.CONSULTAS_ESTADO]));
                        string informe = Norm(Texto(data[i, ConstantsHelper.CONSULTAS_INFORME]));
                        string cc = NormalizarDocumento(Texto(data[i, ConstantsHelper.CONSULTAS_CEDULA]));

                        // D-5: protección extendida, sin acceso a disco ni escritura
                        if (informe == EstadoComision.Verificado || informe == EstadoComision.Reintegro)
                        {
                            r.OmitidasProtegidas++;
                            continue;
                        }

                        string destino;
                        if (estado == EstadoComision.Revocada || estado == EstadoComision.Rechazada || estado == EstadoComision.Anulada)
                        {
                            destino = estado;
                        }
                        else if (indice.Contains(num + "|" + cc))
                        {
                            // D-6: si el archivo existe y la celda tiene una marca de ausencia, se corrige
                            if (EsMarcaFaltante(informe))
                            {
                                _excelRepository.SetCellValue(ws, "H" + i, string.Empty);
                                r.CeldasLimpiadas++;
                            }
                            else
                            {
                                r.ConArchivo++;
                            }
                            continue;
                        }
                        else
                        {
                            destino = EstadoComision.SinSolicitud;
                            r.SinArchivo++;
                        }

                        if (informe == Norm(destino))
                        {
                            r.YaMarcadas++;
                            continue;
                        }

                        // D-5b: se preservan marcas operativas ajenas (ERROR, SIN INFORME, DUPLICADA, ...)
                        if (informe.Length > 0 && !EsMarcaFaltante(informe))
                        {
                            r.OmitidasOtrasMarcas++;
                            continue;
                        }

                        _excelRepository.SetCellValue(ws, "H" + i, destino);
                        r.CeldasEscritas++;
                        if (destino == estado) r.CopiadasEstado++;
                    }
                }
                finally
                {
                    _excelRepository.ScreenUpdating(true);
                }

                if (r.CeldasEscritas + r.CeldasLimpiadas > 0)
                {
                    try
                    {
                        _excelRepository.SaveActiveWorkbook();
                        r.LibroGuardado = true;
                    }
                    catch (Exception ex)
                    {
                        r.ErrorGuardado = ex.Message;
                        LoggerHelper.Warning($"VerificarSolicitudesPresentes: no se pudo guardar el libro: {ex.Message}");
                    }
                }

                LoggerHelper.Resumen("VERIFICAR_SOLICITUDES",
                    ("filas", r.FilasRevisadas),
                    ("protegidas", r.OmitidasProtegidas),
                    ("otrasMarcas", r.OmitidasOtrasMarcas),
                    ("copiadasEstado", r.CopiadasEstado),
                    ("conArchivo", r.ConArchivo),
                    ("sinArchivo", r.SinArchivo),
                    ("yaMarcadas", r.YaMarcadas),
                    ("celdasEscritas", r.CeldasEscritas),
                    ("celdasLimpiadas", r.CeldasLimpiadas),
                    ("libroGuardado", r.LibroGuardado));

                return r;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("VerificarSolicitudesPresentes", ex);
                r.Error = ex.Message;
                return r;
            }
        }

        /// <summary>
        /// Indexa los PDF de la raíz y de la subcarpeta Solicitudes usando la clave "comisión|cédula".
        /// La subcarpeta _Duplicados queda excluida por no ser recursiva.
        /// </summary>
        private static HashSet<string> ConstruirIndiceArchivos(string folder)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string carpeta in CarpetasDeSolicitudes(folder))
            {
                foreach (string file in Directory.GetFiles(carpeta, "*.pdf"))
                {
                    string[] partes = Path.GetFileNameWithoutExtension(file).Split('-');
                    if (partes.Length < 2) continue;

                    string num = partes[0].Trim();
                    string cc = NormalizarDocumento(partes[1]);
                    if (num.Length == 0 || cc.Length == 0) continue;

                    set.Add(num + "|" + cc);
                }
            }
            return set;
        }

        /// <summary> True si la celda INFORME contiene una marca de ausencia de soporte. </summary>
        private static bool EsMarcaFaltante(string informe)
        {
            return informe == EstadoComision.SinSolicitud || informe == EstadoComision.FaltaSolicitud;
        }

        /// <summary> Normaliza una celda de texto para comparaciones. </summary>
        private static string Norm(string valor)
        {
            return (valor ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary> Devuelve el texto de una celda de la matriz de Excel. </summary>
        private static string Texto(object valor)
        {
            return (valor?.ToString() ?? string.Empty).Trim();
        }

        /// <summary>
        /// Orquesta la unión de múltiples PDFs para formar un expediente unificado.
        /// Recolecta Solicitud, Resolución, RP, CDP y Cumplimiento basándose en los datos de la hoja.
        /// Aplica filtro opcional por rango del número de orden (P26).
        /// </summary>
        /// <param name="ws">Hoja de cálculo individual del comisionado.</param>
        /// <param name="s">Carpeta de Solicitudes.</param>
        /// <param name="a">Carpeta de Autorizaciones.</param>
        /// <param name="r">Carpeta de RPs.</param>
        /// <param name="c">Carpeta de CDPs.</param>
        /// <param name="v">Carpeta de Verificados.</param>
        /// <param name="outD">Carpeta de salida para el PDF unificado.</param>
        /// <param name="inicio">Número de viático inicial (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final (0 para sin filtro).</param>
        /// <returns>True si el expediente se unificó correctamente.</returns>
        private bool Unificar(dynamic ws, string s, string a, string r, string c, string v, string outD, int inicio, int fin)
        {
            string sheetName = ws.Name;
            try
            {
                // Extrae datos de las celdas fijas del formato de viáticos
                string order = ws.Range["P26"].Value?.ToString() ?? "";
                string cdp = extraerCDP(ws.Range["B15"].Value?.ToString() ?? "");
                string rp = ws.Range["F16"].Value?.ToString() ?? "";
                string name = ws.Range["I20"].Value?.ToString() ?? "";

                // Validación de datos mínimos requeridos para la búsqueda
                if (string.IsNullOrEmpty(order) || string.IsNullOrEmpty(rp) || string.IsNullOrEmpty(cdp))
                {
                    LoggerHelper.Debug($"Hoja '{sheetName}' omitida: Faltan datos (CDP, RP u Orden).");
                    return false;
                }

                // Filtrar por rango si está activo
                if (!ViaticoNumeroComision.EnRango(order, inicio, fin))
                {
                    LoggerHelper.Debug($"Hoja '{sheetName}' omitida: orden {order} fuera del rango {inicio}-{fin}.");
                    return false;
                }

                name = Regex.Replace(App.CapitalizarPalabras(name), @"\s+", "");
                string outN = $"S_{order}_RP_{rp}_{name}.pdf";

                // Búsqueda de archivos físicos en las respectivas carpetas
                string mS = Buscar(s, order); // Solicitud
                if (string.IsNullOrEmpty(mS)) { LoggerHelper.Warning($"Hoja '{sheetName}': No se encontró solicitud/soporte ({order}) en {s}"); return false; }

                string mR = Buscar(r, rp);    // Registro Presupuestal
                if (string.IsNullOrEmpty(mR)) { LoggerHelper.Warning($"Hoja '{sheetName}': No se encontró RP ({rp}) en {r}"); return false; }

                string mC = Buscar(c, cdp);   // CDP
                if (string.IsNullOrEmpty(mC)) { LoggerHelper.Warning($"Hoja '{sheetName}': No se encontró CDP ({cdp}) en {c}"); return false; }

                // Construcción de la lista de archivos para concatenación (orden específico)
                var fs = new List<string> { mS };
                string mA = Buscar(a, order); // Autorización (Resolución) - opcional
                if (!string.IsNullOrEmpty(mA)) fs.Add(mA);

                fs.Add(mR);
                fs.Add(mC);

                string mV = Buscar(v, order); // Verificado (Cumplimiento) - opcional
                if (!string.IsNullOrEmpty(mV)) fs.Add(mV);

                // Delegación en el servicio de PDF para la unificación física
                new App().CombinarArchivos(fs.ToArray(), Path.Combine(outD, outN));
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error unificando hoja '{sheetName}'", ex);
                return false;
            }
        }

        /// <summary>
        /// Extrae el número de CDP de una cadena de texto (ej: "Soportado con CDP No 1234").
        /// </summary>
        private string extraerCDP(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var match = Regex.Match(text, @"CDP\s*No[\s:]*(\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : "";
        }

        /// <summary>
        /// Busca el primer archivo PDF que contenga el patrón especificado dentro de una carpeta.
        /// Útil para encontrar archivos por ID (ej. número de comisión o CC).
        /// </summary>
        /// <param name="d">Directorio de búsqueda.</param>
        /// <param name="p">Patrón o prefijo a buscar.</param>
        /// <returns>La ruta completa del primer archivo encontrado o vacío.</returns>
        private string Buscar(string d, string p) => (!string.IsNullOrEmpty(p) && Directory.Exists(d)) ? Directory.GetFiles(d, "*.pdf").FirstOrDefault(f => f.Contains(p)) ?? "" : "";

        /// <summary>
        /// DTO interno para manejar la información de identificación de solicitudes.
        /// </summary>
        internal class SInfo { public string Num, CC, Nom; }
    }
}
