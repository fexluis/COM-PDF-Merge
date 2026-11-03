using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Processors
{
    /// <summary>
    /// Implementa la búsqueda de informes usando exclusivamente PdfPig (lectura de texto y
    /// conteo de páginas) y PDFsharp (unificación), sin binarios externos ni archivos temporales.
    /// </summary>
    internal class BusquedaInformesProcessor : IBusquedaInformesProcessor
    {
        private readonly IExcelRepository _excelRepository;
        private readonly IPdfToTexto _pdfToTexto;
        private readonly IPdfService _pdfService;

        /// <summary>
        /// Inicializa el procesador.
        /// </summary>
        /// <param name="excelRepository">Acceso a la hoja CONSULTAS.</param>
        /// <param name="pdfToTexto">Extracción de texto y conteo de páginas (PdfPig).</param>
        /// <param name="pdfService">Unificación de PDFs (PDFsharp).</param>
        public BusquedaInformesProcessor(IExcelRepository excelRepository, IPdfToTexto pdfToTexto, IPdfService pdfService)
        {
            _excelRepository = excelRepository;
            _pdfToTexto = pdfToTexto;
            _pdfService = pdfService;
        }

        /// <inheritdoc/>
        public ResultadoBusquedaInformes RenombrarYUnificar(string carpetaRaiz, int inicio, int fin,
                                                            bool unificarParejas, bool simular,
                                                            Action<int, int> progress = null)
        {
            var resultado = new ResultadoBusquedaInformes { CarpetaRaiz = carpetaRaiz, Simulado = simular };
            var cronometro = Stopwatch.StartNew();

            try
            {
                if (string.IsNullOrWhiteSpace(carpetaRaiz) || !Directory.Exists(carpetaRaiz))
                {
                    resultado.Error = $"La carpeta '{carpetaRaiz}' no existe.";
                    LoggerHelper.Warning("BusquedaInformes: " + resultado.Error);
                    return resultado;
                }

                var solicitudes = GetSolicitudes(inicio, fin);
                LoggerHelper.Resumen("INFORMES_CONSULTAS",
                    ("filas", solicitudes.Count),
                    ("rango", ViaticoNumeroComision.HayFiltro(inicio, fin) ? inicio + "-" + fin : "sin filtro"));

                if (solicitudes.Count == 0)
                {
                    LoggerHelper.Warning("BusquedaInformes: no se encontraron comisiones en CONSULTAS para el rango solicitado.");
                }

                var carpetas = EnumerarCarpetas(carpetaRaiz);
                int totalArchivos = 0;
                foreach (string carpeta in carpetas)
                {
                    try { totalArchivos += Directory.GetFiles(carpeta, "*.pdf").Length; }
                    catch (Exception ex) { LoggerHelper.Warning($"BusquedaInformes: no se pudo listar '{carpeta}': {ex.Message}"); }
                }

                int procesados = 0;

                foreach (string carpeta in carpetas)
                {
                    resultado.CarpetasProcesadas.Add(carpeta);
                    bool esEspecial = EsCarpetaEspecial(carpeta);
                    var pendientes = new List<InformePdf>();

                    string[] archivos;
                    try { archivos = Directory.GetFiles(carpeta, "*.pdf"); }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        LoggerHelper.Warning($"BusquedaInformes: no se pudo acceder a '{carpeta}': {ex.Message}");
                        continue;
                    }

                    foreach (string archivo in archivos)
                    {
                        resultado.PdfsExaminados++;
                        procesados++;

                        InformePdf info = ProcesarArchivo(archivo, carpeta, solicitudes, resultado, simular);
                        if (esEspecial && unificarParejas && info != null && info.EsParValido)
                        {
                            pendientes.Add(info);
                        }

                        Reportar(progress, procesados, totalArchivos);
                    }

                    if (esEspecial && unificarParejas && pendientes.Count > 0)
                    {
                        UnificarParejas(pendientes, carpetaRaiz, resultado, simular);
                    }
                }

                LoggerHelper.Resumen("BUSCAR_INFORMES_RENOMBRAR",
                    ("carpetas", resultado.CarpetasProcesadas.Count),
                    ("pdfs", resultado.PdfsExaminados),
                    ("renombrados", resultado.Renombrados),
                    ("yaCorrectos", resultado.YaCorrectos),
                    ("unificados", resultado.Unificados),
                    ("sinPareja", resultado.RechazadosSinPareja),
                    ("duplicados", resultado.Duplicados),
                    ("colisiones", resultado.Colisiones),
                    ("sinTexto", resultado.SinTexto),
                    ("sinComision", resultado.SinComision),
                    ("errores", resultado.Errores),
                    ("simulado", resultado.Simulado));
            }
            catch (Exception ex)
            {
                resultado.Error = ex.Message;
                LoggerHelper.Error("BusquedaInformes.RenombrarYUnificar", ex);
            }
            finally
            {
                cronometro.Stop();
                resultado.TiempoMs += cronometro.ElapsedMilliseconds;
            }

            return resultado;
        }

        /// <summary>
        /// Lee de CONSULTAS las filas del rango solicitado (número, cédula y nombre).
        /// </summary>
        /// <param name="inicio">Solicitud inicial (0 = sin filtro).</param>
        /// <param name="fin">Solicitud final (0 = sin filtro).</param>
        /// <returns>Filas válidas del rango.</returns>
        internal List<InformeSolicitud> GetSolicitudes(int inicio, int fin)
        {
            var lista = new List<InformeSolicitud>();

            object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
            if (data == null) return lista;

            int filas = data.GetLength(0);
            for (int i = 2; i <= filas; i++)
            {
                string num = (data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? string.Empty).Trim();
                if (num.Length == 0) continue;
                if (!ViaticoNumeroComision.EnRango(num, inicio, fin)) continue;

                lista.Add(new InformeSolicitud
                {
                    Numero = num,
                    Cedula = InformeNombreHelper.SoloDigitos(data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString()),
                    Nombre = (data[i, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO]?.ToString() ?? string.Empty).Trim()
                });
            }

            return lista;
        }

        /// <summary>
        /// Analiza un PDF, decide la cédula y el tipo definitivos y lo renombra al estándar.
        /// Devuelve la información resultante cuando el archivo queda como INF o CUM.
        /// </summary>
        private InformePdf ProcesarArchivo(string archivo, string carpeta, List<InformeSolicitud> solicitudes,
                                           ResultadoBusquedaInformes resultado, bool simular)
        {
            string clave = ClaveDetalle(archivo);
            string nombreActual = Path.GetFileNameWithoutExtension(archivo);

            // 1) ¿Ya cumple el estándar?
            InformePdf existente;
            if (InformeNombreHelper.TryParse(archivo, out existente))
            {
                resultado.YaCorrectos++;
                resultado.Detalle[clave] = MotivoInforme.YaCorrecto;
                existente.RutaOrigen = archivo;
                existente.NombreArchivo = Path.GetFileName(archivo);
                return existente;
            }

            // 2) Número de solicitud del nombre
            string numero = InformeNombreHelper.ExtraerNumeroSolicitud(nombreActual);
            if (string.IsNullOrEmpty(numero))
            {
                resultado.PrefijoInvalido++;
                resultado.Detalle[clave] = MotivoInforme.PrefijoInvalido;
                return null;
            }

            // 3) Datos de CONSULTAS para esa solicitud
            var filas = solicitudes.Where(s => s.Numero == numero).ToList();
            var cedulas = filas.Select(f => f.Cedula)
                               .Where(c => !string.IsNullOrEmpty(c))
                               .Distinct()
                               .ToList();
            string cedulaTabla = cedulas.Count > 0 ? cedulas[0] : string.Empty;
            string nombreTabla = filas.Select(f => f.Nombre).FirstOrDefault(n => !string.IsNullOrWhiteSpace(n));

            if (filas.Count == 0)
            {
                resultado.SinComision++;
            }
            else if (cedulas.Count == 0)
            {
                resultado.SinDatos++;
            }

            // 4) Texto del PDF (página 1 y, si está vacía, página 2) y tipo
            string texto = LeerTexto(archivo);
            if (string.IsNullOrEmpty(texto)) resultado.SinTexto++;

            string tipo = InformeTextoHelper.DetectarTipo(texto);
            ResultadoCedulaPdf cedulaPdf = InformeTextoHelper.BuscarCedula(texto, cedulas);

            // 5) Patrón heredado en el nombre (sirve como segunda fuente de contraste)
            string solPatron, ccPatron, tipoPatron, nomPatron;
            bool tienePatron = InformeNombreHelper.TryParseLoose(nombreActual, out solPatron, out ccPatron, out tipoPatron, out nomPatron);

            // 6) Decisión de cédula y tipo
            bool verificada = cedulaPdf.EsVerificacionFiable &&
                              (cedulas.Count == 0 || cedulas.Contains(cedulaPdf.Cedula));

            string cedulaFinal;
            if (verificada) cedulaFinal = cedulaPdf.Cedula;
            else if (!string.IsNullOrEmpty(cedulaTabla)) cedulaFinal = cedulaTabla;
            else if (tienePatron && !string.IsNullOrEmpty(ccPatron)) cedulaFinal = ccPatron;
            else if (cedulaPdf.Hallada) cedulaFinal = cedulaPdf.Cedula;
            else cedulaFinal = "000000";

            string tipoFinal = tipo;
            string motivo = MotivoInforme.Renombrado;

            if (tipo == InformeTipo.Err)
            {
                tipoFinal = InformeTipo.Err;
                motivo = string.IsNullOrEmpty(texto) ? MotivoInforme.SinTexto : MotivoInforme.SinTipo;
            }
            else if (!verificada)
            {
                tipoFinal = InformeTipo.Err;
                if (string.IsNullOrEmpty(texto)) motivo = MotivoInforme.SinTexto;
                else if (cedulaPdf.Hallada) motivo = MotivoInforme.CedulaNoCoincide;
                else motivo = MotivoInforme.CedulaNoHallada;
            }
            else if (tienePatron && !string.IsNullOrEmpty(ccPatron) &&
                     InformeNombreHelper.NormalizarDocumento(ccPatron) != InformeNombreHelper.NormalizarDocumento(cedulaPdf.Cedula))
            {
                tipoFinal = InformeTipo.Err;
                motivo = MotivoInforme.CedulaNoCoincide;
            }

            string nombreCom = string.IsNullOrWhiteSpace(nombreTabla) ? InformeNombreHelper.NombreDesconocido : nombreTabla;

            return MoverAlEstandar(archivo, nombreActual, numero, cedulaFinal, tipoFinal, nombreCom,
                                   carpeta, motivo, resultado, simular, clave);
        }

        /// <summary>
        /// Aplica el nombre estandarizado al archivo, resolviendo duplicados y colisiones sin
        /// sobrescribir nunca un contenido distinto.
        /// </summary>
        private InformePdf MoverAlEstandar(string archivo, string nombreActual, string numero, string cedula, string tipo,
                                           string nombreCom, string carpeta, string motivo,
                                           ResultadoBusquedaInformes resultado, bool simular, string clave)
        {
            string nuevoNombre = InformeNombreHelper.Construir(numero, cedula, tipo, nombreCom);

            var info = new InformePdf
            {
                NumeroSolicitud = InformeNombreHelper.SoloDigitos(numero),
                Cedula = InformeNombreHelper.SoloDigitos(cedula),
                Tipo = tipo,
                NombreComisionado = InformeNombreHelper.LimpiarNombre(nombreCom),
                RutaOrigen = archivo,
                NombreArchivo = nuevoNombre + ".pdf"
            };

            string destino = Path.Combine(carpeta, nuevoNombre + ".pdf");

            try
            {
                if (string.Equals(nombreActual, nuevoNombre, StringComparison.OrdinalIgnoreCase))
                {
                    resultado.YaCorrectos++;
                    resultado.Detalle[clave] = MotivoInforme.YaCorrecto;
                    return info;
                }

                if (simular)
                {
                    resultado.Renombrados++;
                    resultado.Detalle[clave] = motivo + " (" + MotivoInforme.Simulado + ")";
                    return info;
                }

                if (File.Exists(destino))
                {
                    if (EsMismoContenido(archivo, destino))
                    {
                        string carpetaDuplicados = Path.Combine(carpeta, App.FolderNames.INFORMES_DUPLICADOS);
                        Directory.CreateDirectory(carpetaDuplicados);
                        string rutaDuplicado = InformeNombreHelper.NombreUnico(carpetaDuplicados, Path.GetFileName(archivo));
                        File.Move(archivo, rutaDuplicado);

                        resultado.Duplicados++;
                        resultado.Detalle[clave] = MotivoInforme.Duplicado;
                        return null;
                    }

                    destino = InformeNombreHelper.NombreUnico(carpeta, nuevoNombre + ".pdf");
                    resultado.Colisiones++;
                }

                File.Move(archivo, destino);

                resultado.Renombrados++;
                resultado.Detalle[clave] = motivo;

                info.RutaOrigen = destino;
                info.NombreArchivo = Path.GetFileName(destino);
                return info;
            }
            catch (Exception ex)
            {
                resultado.Errores++;
                resultado.Detalle[clave] = MotivoInforme.ErrorIo;
                LoggerHelper.Warning($"BuscarInformes: no se pudo renombrar '{archivo}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extrae el texto de la primera página del PDF; si está vacío, intenta la segunda.
        /// Devuelve cadena vacía cuando el documento no tiene capa de texto.
        /// </summary>
        private string LeerTexto(string archivo)
        {
            string texto = _pdfToTexto.ExtraerTextoPagina(archivo, 1);
            if (EsTextoValido(texto)) return texto;

            int paginas = _pdfToTexto.ContarPaginas(archivo);
            if (paginas >= 2)
            {
                string segunda = _pdfToTexto.ExtraerTextoPagina(archivo, 2);
                if (EsTextoValido(segunda)) return segunda;
            }

            return string.Empty;
        }

        /// <summary>
        /// True si el texto extraído es utilizable (no vacío y sin el prefijo de error de PdfPig).
        /// </summary>
        private static bool EsTextoValido(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return false;
            return !texto.TrimStart().StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Forma parejas INF + CUM por (solicitud, nombre) y las unifica en la raíz de informes.
        /// </summary>
        private void UnificarParejas(List<InformePdf> pendientes, string carpetaRaiz,
                                     ResultadoBusquedaInformes resultado, bool simular)
        {
            foreach (var grupo in pendientes.GroupBy(p => p.NumeroSolicitud + "|" + p.NombreComisionado))
            {
                InformePdf cum = grupo.FirstOrDefault(p => p.Tipo == InformeTipo.Cum);
                InformePdf inf = grupo.FirstOrDefault(p => p.Tipo == InformeTipo.Inf);

                if (cum == null || inf == null)
                {
                    foreach (InformePdf pendiente in grupo)
                    {
                        resultado.RechazadosSinPareja++;
                        resultado.Detalle[ClaveDetalle(pendiente.RutaOrigen)] = MotivoInforme.SinPareja;
                    }
                    continue;
                }

                Unificar(cum, inf, carpetaRaiz, resultado, simular);
            }
        }

        /// <summary>
        /// Une CUM + INF en un archivo UNI (CUM primero), valida el número de páginas y solo
        /// entonces elimina los originales.
        /// </summary>
        private void Unificar(InformePdf cum, InformePdf inf, string carpetaRaiz,
                              ResultadoBusquedaInformes resultado, bool simular)
        {
            string nombreUni = InformeNombreHelper.Construir(cum.NumeroSolicitud, cum.Cedula, InformeTipo.Uni, cum.NombreComisionado);
            string salida = Path.Combine(carpetaRaiz, nombreUni + ".pdf");
            string clave = ClaveDetalle(cum.RutaOrigen);

            if (simular)
            {
                resultado.Unificados++;
                resultado.Detalle[clave] = MotivoInforme.Unificado + " (" + MotivoInforme.Simulado + ")";
                return;
            }

            string temporal = Path.Combine(carpetaRaiz, Guid.NewGuid().ToString("N") + ".tmp.pdf");

            try
            {
                int paginasEsperadas = _pdfToTexto.ContarPaginas(cum.RutaOrigen) + _pdfToTexto.ContarPaginas(inf.RutaOrigen);

                _pdfService.CombinarPDFs(new[] { cum.RutaOrigen, inf.RutaOrigen }, temporal);

                int paginasObtenidas = _pdfToTexto.ContarPaginas(temporal);
                if (paginasObtenidas <= 0 || (paginasEsperadas > 0 && paginasObtenidas != paginasEsperadas))
                {
                    throw new InvalidOperationException(
                        $"El PDF unificado tiene {paginasObtenidas} página(s) y se esperaban {paginasEsperadas}.");
                }

                if (File.Exists(salida))
                {
                    if (EsMismoContenido(temporal, salida))
                    {
                        File.Delete(temporal);
                        resultado.Duplicados++;
                    }
                    else
                    {
                        salida = InformeNombreHelper.NombreUnico(carpetaRaiz, nombreUni + ".pdf");
                        File.Copy(temporal, salida, true);
                        File.Delete(temporal);
                        resultado.Colisiones++;
                    }
                }
                else
                {
                    File.Move(temporal, salida);
                }

                BorrarSeguro(cum.RutaOrigen);
                BorrarSeguro(inf.RutaOrigen);

                resultado.Unificados++;
                resultado.Detalle[clave] = MotivoInforme.Unificado;
            }
            catch (Exception ex)
            {
                resultado.MergeFallidos++;
                resultado.Detalle[clave] = MotivoInforme.ErrorMerge;
                LoggerHelper.Warning(
                    $"BuscarInformes: falló la unificación de '{cum.NombreArchivo}' + '{inf.NombreArchivo}': {ex.Message}");
            }
            finally
            {
                if (File.Exists(temporal))
                {
                    try { File.Delete(temporal); }
                    catch (Exception ex) { LoggerHelper.Warning($"BuscarInformes: no se pudo borrar el temporal '{temporal}': {ex.Message}"); }
                }
            }
        }

        /// <summary>
        /// Enumera la carpeta raíz y todas sus subcarpetas, excluyendo VERIFICADOS, _Duplicados
        /// y cualquier carpeta que inicie con guion bajo.
        /// </summary>
        private static List<string> EnumerarCarpetas(string raiz)
        {
            var carpetas = new List<string>();
            var pendientes = new Queue<string>();
            pendientes.Enqueue(raiz);

            while (pendientes.Count > 0)
            {
                string actual = pendientes.Dequeue();
                carpetas.Add(actual);

                string[] hijas;
                try { hijas = Directory.GetDirectories(actual); }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"BuscarInformes: no se pudo listar subcarpetas de '{actual}': {ex.Message}");
                    continue;
                }

                foreach (string hija in hijas)
                {
                    if (EsCarpetaExcluida(Path.GetFileName(hija))) continue;
                    pendientes.Enqueue(hija);
                }
            }

            return carpetas;
        }

        /// <summary> True si la carpeta no debe procesarse (VERIFICADOS, _Duplicados, ocultas). </summary>
        private static bool EsCarpetaExcluida(string nombreCarpeta)
        {
            if (string.IsNullOrEmpty(nombreCarpeta)) return true;

            return string.Equals(nombreCarpeta, App.FolderNames.INFORMES_VERIFICADOS, StringComparison.OrdinalIgnoreCase)
                || string.Equals(nombreCarpeta, App.FolderNames.INFORMES_DUPLICADOS, StringComparison.OrdinalIgnoreCase)
                || nombreCarpeta.StartsWith("_", StringComparison.Ordinal);
        }

        /// <summary> True si la carpeta es una de las que admiten unificación de parejas. </summary>
        private static bool EsCarpetaEspecial(string carpeta)
        {
            string nombre = Path.GetFileName(carpeta);
            if (string.IsNullOrEmpty(nombre)) return false;

            foreach (string especial in ConstantsHelper.INFORMES_CARPETAS_ESPECIALES)
            {
                if (string.Equals(nombre, especial, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        /// <summary> Clave de trazabilidad carpeta/archivo para el diccionario de detalle. </summary>
        private static string ClaveDetalle(string ruta)
        {
            if (string.IsNullOrEmpty(ruta)) return string.Empty;

            string carpeta = Path.GetFileName(Path.GetDirectoryName(ruta) ?? string.Empty);
            return carpeta + "/" + Path.GetFileName(ruta);
        }

        /// <summary> Notifica el progreso sin permitir que un fallo del callback rompa el proceso. </summary>
        private static void Reportar(Action<int, int> progress, int actual, int total)
        {
            if (progress == null || total <= 0) return;

            try { progress(actual, total); }
            catch (Exception ex) { LoggerHelper.Warning($"BuscarInformes: callback de progreso falló: {ex.Message}"); }
        }

        /// <summary> True si dos archivos tienen exactamente el mismo contenido. </summary>
        private static bool EsMismoContenido(string rutaA, string rutaB)
        {
            try
            {
                var infoA = new FileInfo(rutaA);
                var infoB = new FileInfo(rutaB);
                if (!infoA.Exists || !infoB.Exists || infoA.Length != infoB.Length) return false;
                return CalcularHash(rutaA) == CalcularHash(rutaB);
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"BuscarInformes: no se pudo comparar '{rutaA}' con '{rutaB}': {ex.Message}");
                return false;
            }
        }

        /// <summary> Hash SHA-256 en hexadecimal del contenido de un archivo. </summary>
        private static string CalcularHash(string ruta)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(ruta))
            {
                return BitConverter.ToString(sha.ComputeHash(stream));
            }
        }

        /// <summary> Elimina un archivo sin propagar excepciones. </summary>
        private static bool BorrarSeguro(string ruta)
        {
            try
            {
                if (File.Exists(ruta)) File.Delete(ruta);
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"BuscarInformes: no se pudo eliminar '{ruta}': {ex.Message}");
                return false;
            }
        }

        /// <inheritdoc/>
        public ResultadoBusquedaInformes AuditarConsultas(string carpetaRaiz, int inicio, int fin,
                                                          bool simular, Action<int, int> progress = null)
        {
            var resultado = new ResultadoBusquedaInformes { CarpetaRaiz = carpetaRaiz, Simulado = simular };
            var cronometro = Stopwatch.StartNew();

            try
            {
                if (string.IsNullOrWhiteSpace(carpetaRaiz) || !Directory.Exists(carpetaRaiz))
                {
                    resultado.Error = $"La carpeta '{carpetaRaiz}' no existe.";
                    return resultado;
                }

                object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
                if (data == null)
                {
                    resultado.Error = "No se pudo leer la hoja CONSULTAS.";
                    return resultado;
                }

                dynamic hoja = _excelRepository.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS);
                if (hoja == null)
                {
                    resultado.Error = "No se encontró la hoja CONSULTAS en el libro activo.";
                    return resultado;
                }

                var indice = ConstruirIndice(carpetaRaiz);

                // 1) UNI presentes en la raíz: se validan contra CONSULTAS y se mueven a VERIFICADOS.
                var verificados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var conErrorUni = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                ProcesarUnificados(carpetaRaiz, data, indice, resultado, simular, verificados, conErrorUni);

                // 2) Filas de CONSULTAS dentro del rango: se recalcula y escribe la columna INFORME.
                int filas = data.GetLength(0);
                int total = Math.Max(filas - 1, 1);

                for (int i = 2; i <= filas; i++)
                {
                    string num = (data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? string.Empty).Trim();
                    if (num.Length == 0) continue;
                    if (!ViaticoNumeroComision.EnRango(num, inicio, fin)) continue;

                    resultado.FilasRevisadas++;
                    Reportar(progress, i, total);

                    string actual = (data[i, ConstantsHelper.CONSULTAS_INFORME]?.ToString() ?? string.Empty).Trim().ToUpperInvariant();
                    if (EsEstadoProtegido(actual))
                    {
                        resultado.OmitidasProtegidas++;
                        continue;
                    }

                    string cedula = InformeNombreHelper.SoloDigitos(data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString());
                    string deseado = CalcularEstado(num, cedula, verificados, conErrorUni, indice, resultado);

                    if (deseado == EstadoComision.ParOk) continue;
                    if (string.Equals(actual, deseado, StringComparison.Ordinal)) continue;

                    if (actual.Length > 0 && !EsEstadoValido(actual)) resultado.CeldasLimpiadas++;

                    if (!simular)
                    {
                        try
                        {
                            hoja.Cells[i, ConstantsHelper.CONSULTAS_INFORME] = deseado;
                        }
                        catch (Exception ex)
                        {
                            resultado.Errores++;
                            LoggerHelper.Warning($"BuscarInformes: no se pudo escribir INFORME en la fila {i}: {ex.Message}");
                            continue;
                        }
                    }

                    resultado.CeldasEscritas++;
                    ContabilizarEstado(resultado, deseado);
                    resultado.Detalle["CONSULTAS!" + i] = deseado;
                }

                if (!simular && resultado.CeldasEscritas > 0)
                {
                    try
                    {
                        _excelRepository.SaveActiveWorkbook();
                        resultado.LibroGuardado = true;
                    }
                    catch (Exception ex)
                    {
                        resultado.ErrorGuardado = ex.Message;
                        LoggerHelper.Error("BuscarInformes: no se pudo guardar el libro", ex);
                    }
                }

                LoggerHelper.Resumen("BUSCAR_INFORMES_AUDITORIA",
                    ("filas", resultado.FilasRevisadas),
                    ("protegidas", resultado.OmitidasProtegidas),
                    ("verificados", resultado.Verificados),
                    ("faltaCum", resultado.FaltaCum),
                    ("faltaInf", resultado.FaltaInf),
                    ("error", resultado.ConError),
                    ("sinInforme", resultado.SinInforme),
                    ("parOk", resultado.ParOk),
                    ("celdasEscritas", resultado.CeldasEscritas),
                    ("guardado", resultado.LibroGuardado),
                    ("simulado", resultado.Simulado));
            }
            catch (Exception ex)
            {
                resultado.Error = ex.Message;
                LoggerHelper.Error("BusquedaInformes.AuditarConsultas", ex);
            }
            finally
            {
                cronometro.Stop();
                resultado.TiempoMs += cronometro.ElapsedMilliseconds;
            }

            return resultado;
        }

        /// <summary>
        /// Indexa los PDF de \Informes (excluyendo VERIFICADOS y _Duplicados) por número de solicitud,
        /// indicando si están en la raíz y si están en una carpeta especial.
        /// </summary>
        private static Dictionary<string, List<ArchivoIndexado>> ConstruirIndice(string carpetaRaiz)
        {
            var indice = new Dictionary<string, List<ArchivoIndexado>>(StringComparer.OrdinalIgnoreCase);
            string raiz = Path.GetFullPath(carpetaRaiz).TrimEnd(Path.DirectorySeparatorChar);

            foreach (string carpeta in EnumerarCarpetas(carpetaRaiz))
            {
                bool esEspecial = EsCarpetaEspecial(carpeta);
                bool enRaiz = string.Equals(
                    Path.GetFullPath(carpeta).TrimEnd(Path.DirectorySeparatorChar),
                    raiz,
                    StringComparison.OrdinalIgnoreCase);

                string[] archivos;
                try { archivos = Directory.GetFiles(carpeta, "*.pdf"); }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"BuscarInformes: no se pudo listar '{carpeta}': {ex.Message}");
                    continue;
                }

                foreach (string archivo in archivos)
                {
                    string solicitud, cedula, tipo, nombre;
                    if (!InformeNombreHelper.TryParseLoose(archivo, out solicitud, out cedula, out tipo, out nombre)) continue;
                    if (!InformeTipo.EsTipoValido(tipo)) continue;

                    List<ArchivoIndexado> lista;
                    if (!indice.TryGetValue(solicitud, out lista))
                    {
                        lista = new List<ArchivoIndexado>();
                        indice[solicitud] = lista;
                    }

                    lista.Add(new ArchivoIndexado
                    {
                        Pdf = new InformePdf
                        {
                            NumeroSolicitud = solicitud,
                            Cedula = cedula,
                            Tipo = tipo,
                            NombreComisionado = nombre,
                            RutaOrigen = archivo,
                            NombreArchivo = Path.GetFileName(archivo)
                        },
                        EsCarpetaEspecial = esEspecial,
                        EnRaiz = enRaiz
                    });
                }
            }

            return indice;
        }

        /// <summary>
        /// Valida los archivos UNI de la raíz contra CONSULTAS: los que coinciden en cédula se mueven
        /// a VERIFICADOS y los que no se renombran a ERR.
        /// </summary>
        private void ProcesarUnificados(string carpetaRaiz, object[,] data,
                                        Dictionary<string, List<ArchivoIndexado>> indice,
                                        ResultadoBusquedaInformes resultado, bool simular,
                                        HashSet<string> verificados, HashSet<string> conErrorUni)
        {
            var unificados = indice.Values
                                   .SelectMany(l => l)
                                   .Where(a => a.EnRaiz && a.Pdf.Tipo == InformeTipo.Uni)
                                   .ToList();

            foreach (ArchivoIndexado archivo in unificados)
            {
                InformePdf pdf = archivo.Pdf;
                string cedula = InformeNombreHelper.NormalizarDocumento(pdf.Cedula);
                string claveFila = pdf.NumeroSolicitud + "|" + cedula;
                string claveDetalle = ClaveDetalle(pdf.RutaOrigen);

                if (CedulaCoincideEnConsultas(data, pdf.NumeroSolicitud, cedula))
                {
                    verificados.Add(claveFila);

                    if (simular)
                    {
                        resultado.Detalle[claveDetalle] = EstadoComision.Verificado + " (" + MotivoInforme.Simulado + ")";
                        continue;
                    }

                    try
                    {
                        string carpetaVerificados = Path.Combine(carpetaRaiz, App.FolderNames.INFORMES_VERIFICADOS);
                        Directory.CreateDirectory(carpetaVerificados);

                        string destino = InformeNombreHelper.NombreUnico(carpetaVerificados, pdf.NombreArchivo);
                        File.Move(pdf.RutaOrigen, destino);

                        resultado.Detalle[claveDetalle] = EstadoComision.Verificado;
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        LoggerHelper.Warning($"BuscarInformes: no se pudo mover '{pdf.NombreArchivo}' a VERIFICADOS: {ex.Message}");
                    }
                }
                else
                {
                    conErrorUni.Add(claveFila);

                    string nombreError = InformeNombreHelper.Construir(
                        pdf.NumeroSolicitud, pdf.Cedula, InformeTipo.Err, pdf.NombreComisionado);

                    if (!simular)
                    {
                        try
                        {
                            string destino = InformeNombreHelper.NombreUnico(carpetaRaiz, nombreError + ".pdf");
                            File.Move(pdf.RutaOrigen, destino);
                        }
                        catch (Exception ex)
                        {
                            resultado.Errores++;
                            LoggerHelper.Warning($"BuscarInformes: no se pudo renombrar '{pdf.NombreArchivo}' a ERR: {ex.Message}");
                        }
                    }

                    resultado.Detalle[claveDetalle] = MotivoInforme.CedulaNoCoincide;
                }
            }
        }

        /// <summary>
        /// True si alguna fila de CONSULTAS con esa solicitud tiene la cédula indicada.
        /// </summary>
        private static bool CedulaCoincideEnConsultas(object[,] data, string solicitud, string cedula)
        {
            if (string.IsNullOrEmpty(solicitud) || string.IsNullOrEmpty(cedula)) return false;

            int filas = data.GetLength(0);
            for (int i = 2; i <= filas; i++)
            {
                string num = (data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? string.Empty).Trim();
                if (!string.Equals(num, solicitud, StringComparison.Ordinal)) continue;

                string cc = InformeNombreHelper.SoloDigitos(data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString());
                if (string.Equals(cc, cedula, StringComparison.Ordinal)) return true;
            }

            return false;
        }

        /// <summary>
        /// Calcula el estado que debe llevar la columna INFORME para una solicitud/cédula.
        /// Devuelve <see cref="EstadoComision.ParOk"/> cuando la pareja está completa y no hay que escribir.
        /// </summary>
        private static string CalcularEstado(string numero, string cedula,
                                             HashSet<string> verificados, HashSet<string> conErrorUni,
                                             Dictionary<string, List<ArchivoIndexado>> indice,
                                             ResultadoBusquedaInformes resultado)
        {
            string claveFila = numero + "|" + InformeNombreHelper.NormalizarDocumento(cedula);

            if (verificados.Contains(claveFila)) return EstadoComision.Verificado;
            if (conErrorUni.Contains(claveFila)) return EstadoComision.Error;

            List<ArchivoIndexado> archivos;
            if (!indice.TryGetValue(numero, out archivos)) return EstadoComision.SinInforme;

            bool tieneErr = false;
            bool tieneInf = false;
            bool tieneCum = false;

            foreach (ArchivoIndexado archivo in archivos)
            {
                if (!archivo.EsCarpetaEspecial) continue;

                if (archivo.Pdf.Tipo == InformeTipo.Err) tieneErr = true;
                else if (archivo.Pdf.Tipo == InformeTipo.Inf) tieneInf = true;
                else if (archivo.Pdf.Tipo == InformeTipo.Cum) tieneCum = true;
            }

            if (tieneErr) return EstadoComision.Error;
            if (tieneInf && tieneCum)
            {
                resultado.ParOk++;
                return EstadoComision.ParOk;
            }
            if (tieneInf) return EstadoComision.FaltaCum;
            if (tieneCum) return EstadoComision.FaltaInf;

            return EstadoComision.SinInforme;
        }

        /// <summary> Suma el contador correspondiente al estado escrito en INFORME. </summary>
        private static void ContabilizarEstado(ResultadoBusquedaInformes resultado, string estado)
        {
            if (estado == EstadoComision.Verificado) resultado.Verificados++;
            else if (estado == EstadoComision.FaltaCum) resultado.FaltaCum++;
            else if (estado == EstadoComision.FaltaInf) resultado.FaltaInf++;
            else if (estado == EstadoComision.Error) resultado.ConError++;
            else if (estado == EstadoComision.SinInforme) resultado.SinInforme++;
        }

        /// <summary> Estados de la columna INFORME que no deben sobrescribirse. </summary>
        private static bool EsEstadoProtegido(string estado)
        {
            return estado == EstadoComision.Verificado
                || estado == EstadoComision.Reintegro
                || estado == EstadoComision.Revocada
                || estado == EstadoComision.Rechazada
                || estado == EstadoComision.Anulada;
        }

        /// <summary> Estados que el proceso reconoce como propios (para detectar marcas obsoletas). </summary>
        private static bool EsEstadoValido(string estado)
        {
            return EsEstadoProtegido(estado)
                || estado == EstadoComision.FaltaCum
                || estado == EstadoComision.FaltaInf
                || estado == EstadoComision.Error
                || estado == EstadoComision.SinInforme
                || estado == EstadoComision.SinSolicitud
                || estado == EstadoComision.FaltaSolicitud;
        }

        /// <summary> PDF indexado con la ubicación relativa que ocupa dentro de \Informes. </summary>
        private sealed class ArchivoIndexado
        {
            /// <summary> Datos del archivo interpretados desde su nombre. </summary>
            public InformePdf Pdf { get; set; }

            /// <summary> True si está en una carpeta especial (Academica, Capacitacion, Direccion, Fortalecimiento). </summary>
            public bool EsCarpetaEspecial { get; set; }

            /// <summary> True si está directamente en la raíz de \Informes. </summary>
            public bool EnRaiz { get; set; }
        }
    }
}
