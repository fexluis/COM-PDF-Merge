using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Helper principal de unificación de Ordenes de Pago (OP).
    /// Paridad funcional con BuscarOrdenPago / ListFilesContainingSearchTerms de VBA (AMenu.bas).
    /// Orquesta el inventario de boletines/soportes, la correspondencia por clave (D-6, D-8), la resolución
    /// de categorías (D-4), el saneamiento determinista del nombre de salida (D-1b, D-2, C2, C3, Fase 11) y la preservación en histórico.
    /// </summary>
    public static class UnificarOrdenPagoHelper
    {
        private const string DOC_TYPE_POR_DEFECTO = "OP";

        /// <summary>
        /// Unifica los soportes PDF (ubicados en la raíz) con el boletín correspondiente de cada hoja
        /// "ORDEN DE PAGO PRESUPUESTAL" del libro activo, generando 1 solo PDF unificado por hoja.
        /// </summary>
        /// <param name="activeWorkbook">Objeto COM del libro activo de Excel.</param>
        /// <param name="util">Instancia de la aplicación <see cref="App"/> (provee CombinarArchivos y resolución OneDrive).</param>
        /// <param name="docType">Tipo de documento para la clave de match: "OP" (por Orden, por defecto) u "OB" (por Obligación).</param>
        /// <param name="progressCallback">Callback para actualizar el progreso (hojaActual, totalHojas).</param>
        /// <returns>
        /// Mayor o igual a 0: Número total de PDFs unificados generados.
        /// -1: No se pudo resolver la ruta local del libro.
        /// -2: No hay boletines ni soportes activos para unificar.
        /// -3: Tipo de documento docType no válido.
        /// -4: Argumentos nulos (activeWorkbook o util).
        /// </returns>
        public static int UnificarOrdenPago(dynamic activeWorkbook, App util,
            string docType = DOC_TYPE_POR_DEFECTO, Action<int, int> progressCallback = null,
            bool forzarSobrescritura = false)
        {
            if (activeWorkbook == null || util == null)
            {
                LoggerHelper.Error("UnificarOrdenPago: workbook o util nulos.");
                return -4;
            }

            docType = string.IsNullOrWhiteSpace(docType)
                ? DOC_TYPE_POR_DEFECTO : docType.Trim().ToUpperInvariant();

            if (docType != "OP" && docType != "OB")
            {
                LoggerHelper.Error($"docType no válido: '{docType}'. Se esperaba 'OP' u 'OB'.");
                return -3;
            }

            // 1. Resolver ruta local sincronizada (Fase 2: Decisión D-9 sin parche hardcodeado)
            string localPath = FirmaPDFsHelper.ObtenerRutaLocal(activeWorkbook, util);
            if (string.IsNullOrEmpty(localPath))
            {
                LoggerHelper.Warning("No se pudo resolver la ruta local del libro activo.");
                return -1;
            }

            // 2. Definición y creación de carpetas
            string pathBoletines = Path.Combine(localPath, App.FolderNames.BOLETINES);
            string pathOk = Path.Combine(pathBoletines, App.FolderNames.OrdenesDePago);
            string pathUnif = Path.Combine(pathBoletines, App.FolderNames.UNIFICADOS);

            // Carpetas históricas legadas del VBA (Fase 5 / Hallazgo Legado)
            string pathOkLegacy = Path.Combine(pathBoletines, "OK");
            string pathUnLegacy = Path.Combine(pathBoletines, "UN");

            if (!Directory.Exists(pathBoletines))
            {
                LoggerHelper.Warning($"No existe el directorio de boletines: {pathBoletines}");
                return -2;
            }

            FirmaPDFsHelper.CrearCarpetaSiNoExiste(pathOk);
            FirmaPDFsHelper.CrearCarpetaSiNoExiste(pathUnif);

            // 3. Construir inventario de boletines excluyendo las carpetas históricas actuales y legadas (OK, UN)
            List<string> boletines = new List<string>();
            foreach (string f in Directory.GetFiles(pathBoletines, "*.pdf", SearchOption.AllDirectories))
            {
                if (EstaEnCarpeta(f, pathOk) || EstaEnCarpeta(f, pathUnif) ||
                    EstaEnCarpeta(f, pathOkLegacy) || EstaEnCarpeta(f, pathUnLegacy))
                {
                    continue;
                }
                boletines.Add(f);
            }
            boletines.Sort(StringComparer.OrdinalIgnoreCase);

            // Inventario de soportes en la raíz y pre-parseo único (M4: evita re-parsear N veces por hoja)
            string[] soportesRaiz = Directory.GetFiles(localPath, "*.pdf", SearchOption.TopDirectoryOnly);
            Dictionary<string, OrdenPagoNombreSoportePdf> soportesParsed = new Dictionary<string, OrdenPagoNombreSoportePdf>(StringComparer.OrdinalIgnoreCase);

            foreach (string rutaSoporte in soportesRaiz)
            {
                OrdenPagoNombreSoportePdf datosSoporte;
                if (OrdenPagoNombreSoportePdfParser.TryParse(Path.GetFileName(rutaSoporte), out datosSoporte))
                {
                    soportesParsed[rutaSoporte] = datosSoporte;
                }
                else
                {
                    LoggerHelper.Warning($"Soporte no reconocido por el patrón: {Path.GetFileName(rutaSoporte)}");
                }
            }

            // Hallazgo C5: Guarda sobre los elementos activos no archivados
            if (boletines.Count == 0 && soportesParsed.Count == 0)
            {
                LoggerHelper.Warning("No se encontraron boletines ni soportes activos en la estructura de directorios.");
                return -2;
            }

            LoggerHelper.Info($"UnificarOrdenPago: ruta={localPath} boletines={boletines.Count} soportesValidosRaiz={soportesParsed.Count} docType={docType}");

            dynamic excelApp = activeWorkbook.Application;
            int filesCount = 0;

            using (ExcelOptimizeHelper.Optimize(excelApp))
            {
                // 4. Identificar hojas de tipo Orden de Pago
                List<dynamic> hojas = new List<dynamic>();
                foreach (dynamic ws in activeWorkbook.Worksheets)
                {
                    if (OrdenPagoSheetReader.EsHojaOrdenPago(ws)) hojas.Add(ws);
                }

                int total = hojas.Count;
                int i = 0;

                foreach (dynamic ws in hojas)
                {
                    i++;
                    if (progressCallback != null) progressCallback(i, total);

                    OrdenPagoDatosHoja datos = OrdenPagoSheetReader.LeerHojaOp(ws);
                    if (!datos.EsValida)
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': sin nº de Orden de Pago en B15. Se omite.");
                        continue;
                    }

                    // Fase 12: Validación estricta de Obligación cuando docType="OB"
                    if (string.Equals(docType, "OB", StringComparison.OrdinalIgnoreCase)
                        && string.IsNullOrWhiteSpace(datos.Obligacion))
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': sin nº de Obligación (docType=OB). Se omite.");
                        continue;
                    }

                    // 4.1 Clave de match (Decisión D-6)
                    string clave = string.Equals(docType, "OB", StringComparison.OrdinalIgnoreCase)
                        ? datos.Obligacion : datos.Orden;

                    // 4.2 Selección de Boletín y Soportes (usando diccionario pre-parseado)
                    string boletin = BuscarBoletin(boletines, datos, docType, clave);

                    // Fase 10: Verificar si el boletín ya fue movido por otra hoja anterior en este ciclo
                    if (boletin != null && !File.Exists(boletin))
                    {
                        LoggerHelper.Warning($"Boletín ya utilizado/archivado: {Path.GetFileName(boletin)}. Se omite en hoja '{datos.NombreHoja}'.");
                        boletin = null;
                    }

                    List<string> soportesHoja = SeleccionarSoportes(soportesParsed, docType, datos.Orden, datos.Obligacion);

                    // 4.3 Resolución de categoría y nombre determinista (Fases 20-57)
                    // Nota (Fase 57): ResolverCategoria evalúa primero soportes en memoria; si prefijo >= 19, ConstruirNombreSalida omitirá la categoría por presupuesto de 30 chars.
                    string categoria = ResolverCategoria(soportesHoja, soportesParsed, ws);

                    if (string.IsNullOrWhiteSpace(datos.Obligacion))
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': obligación ilegible (X16/U16). Se usa la OP '{datos.Orden}' como clave.");
                    }
                    string claveNombre = string.IsNullOrWhiteSpace(datos.Obligacion) ? datos.Orden : datos.Obligacion;

                    // Fase 44: Presupuesto de longitud de ruta para evitar PathTooLongException en Win32 / Excel
                    const int MAX_PATH_SEGURO = 240;   // margen para ".pdf" + ".tmp" y rutas UNC
                    int presupuesto = MAX_PATH_SEGURO - pathUnif.Length - 1 - ".pdf".Length - ".tmp".Length;
                    if (presupuesto < 25)
                    {
                        LoggerHelper.Error($"Hoja '{datos.NombreHoja}': ruta de salida demasiado larga ({pathUnif.Length} chars). Se omite para evitar truncado de nombre.");
                        continue;
                    }

                    string nombreSalida = ConstruirNombreSalida(datos.FechaPago, claveNombre, categoria, datos.Tercero, Math.Min(MAX_LARGO_NOMBRE, presupuesto));
                    string rutaSalida = Path.Combine(pathUnif, nombreSalida + ".pdf");

                    // 4.4 Idempotencia (D-7) y tolerancias de proceso (D-3)
                    if (boletin == null && File.Exists(rutaSalida) && !forzarSobrescritura)
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': boletín ya archivado y salida existente. Se omite (idempotente). Use forzarSobrescritura=true para regenerar.");
                        continue;
                    }
                    if (boletin == null && soportesHoja.Count == 0)
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': sin boletín y sin soportes. Se omite.");
                        continue;
                    }
                    if (boletin == null)
                    {
                        LoggerHelper.Warning($"Hoja '{datos.NombreHoja}': no se halló boletín para {docType}={clave}. Se unifica solo con soportes.");
                    }

                    // 4.5 Construcción de lista ordenada (boletín primero, luego soportes alfabéticamente)
                    List<string> lista = new List<string>();
                    if (boletin != null) lista.Add(boletin);
                    lista.AddRange(soportesHoja);

                    string rutaTmp = null;
                    try
                    {
                        // Recomendación M3: Generación atómica usando archivo temporal
                        rutaTmp = rutaSalida + ".tmp";
                        if (File.Exists(rutaTmp)) File.Delete(rutaTmp);

                        util.CombinarArchivos(lista.ToArray(), rutaTmp);

                        if (File.Exists(rutaSalida))
                        {
                            LoggerHelper.Warning($"Se sobrescribe unificado existente: {Path.GetFileName(rutaSalida)}");
                            File.Replace(rutaTmp, rutaSalida, null, true);   // Fase 86: sustitucion atomica, sin ventana de perdida
                        }
                        else
                        {
                            File.Move(rutaTmp, rutaSalida);
                        }
                        filesCount++;

                        // Fase 15: Mover boletín procesado a la carpeta histórica tras éxito y remover de la lista en memoria solo si el traslado fue efectivo
                        if (boletin != null)
                        {
                            MoverAHistorico(boletin, Path.Combine(pathOk, Path.GetFileName(boletin)));
                            if (!File.Exists(boletin))
                            {
                                boletines.Remove(boletin);
                            }
                        }

                        LoggerHelper.Info($"Unificado OK: hoja={datos.NombreHoja} {docType}={clave} " +
                                          $"boletin={(boletin == null ? "-" : Path.GetFileName(boletin))} " +
                                          $"soportes={soportesHoja.Count} salida={Path.GetFileName(rutaSalida)}");
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error($"Error unificando hoja '{datos.NombreHoja}' ({docType}={clave})", ex);
                        
                        // Fase 17: Limpieza del archivo temporal huérfano si el merge falla
                        try
                        {
                            if (!string.IsNullOrEmpty(rutaTmp) && File.Exists(rutaTmp))
                            {
                                File.Delete(rutaTmp);
                            }
                        }
                        catch (Exception exDel)
                        {
                            LoggerHelper.Warning($"No se pudo limpiar temporal '{rutaTmp}': {exDel.Message}");
                        }
                    }
                }
            }

            return filesCount;
        }

        /// <summary>
        /// Selecciona y ordena los soportes PDF que coinciden con los criterios de la hoja desde el diccionario pre-parseado.
        /// </summary>
        internal static List<string> SeleccionarSoportes(
            Dictionary<string, OrdenPagoNombreSoportePdf> soportesParsed,
            string docType, string orden, string obligacion)
        {
            List<string> resultado = new List<string>();
            if (soportesParsed == null) return resultado;

            foreach (KeyValuePair<string, OrdenPagoNombreSoportePdf> kvp in soportesParsed)
            {
                if (Coincide(kvp.Value, docType, orden, obligacion))
                {
                    resultado.Add(kvp.Key);
                }
            }

            resultado.Sort(StringComparer.OrdinalIgnoreCase);
            return resultado;
        }

        /// <summary>
        /// Compara dos cadenas numéricas ignorando ceros a la izquierda y espacios (M5 / Severity Logic).
        /// </summary>
        internal static bool NumerosIguales(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            string na = a.Trim().TrimStart('0');
            string nb = b.Trim().TrimStart('0');
            if (na.Length == 0) na = "0";
            if (nb.Length == 0) nb = "0";
            return string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Evalúa si los datos del soporte coinciden con la Orden u Obligación según el tipo de documento (Decisión D-6)
        /// aplicando la comparación numéricamente tolerante.
        /// </summary>
        internal static bool Coincide(OrdenPagoNombreSoportePdf datos, string docType, string orden, string obligacion)
        {
            if (datos == null) return false;
            if (string.Equals(docType, "OB", StringComparison.OrdinalIgnoreCase))
            {
                return NumerosIguales(datos.Obligacion, obligacion);
            }
            return NumerosIguales(datos.Orden, orden);
        }

        internal const int MAX_LARGO_NOMBRE = 30;
        private const int MIN_TERCERO_VISIBLE = 8;

        /// <summary>
        /// Normaliza la fecha de pago a formato yyyyMMdd de forma determinista (Fases 22, 24, 29, 33).
        /// Devuelve "SINFECHA" si la fecha es nula, vacía o no válida para mantener la idempotencia (Decisión D-7).
        /// </summary>
        internal static string NormalizarFecha(string fechaPago)
        {
            if (!string.IsNullOrWhiteSpace(fechaPago))
            {
                string fp = fechaPago.Trim();
                DateTime fd;
                if (fp.Length == 8 && DateTime.TryParseExact(fp, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fd))
                {
                    return fp;
                }

                string norm = OrdenPagoSheetReader.LeerFechaPagoDesdeTexto(fp);
                if (norm.Length == 8)
                {
                    return norm;
                }
            }

            LoggerHelper.Warning($"FechaPago vacía o no reconocible ('{fechaPago}'). Se usa 'SINFECHA'.");
            return "SINFECHA";
        }

        /// <summary>
        /// Formatea el código de categoría a TitleCase (ej: "VIA" -> "Via") de forma independiente del locale (Fase 35)
        /// o devuelve cadena vacía para omitirla si la categoría es "XXX", nula o vacía (Fases 20, 23, 28).
        /// </summary>
        internal static string FormatearCategoria(string categoria)
        {
            if (string.IsNullOrWhiteSpace(categoria)) return "";
            string c = categoria.Trim();
            if (string.Equals(c, "XXX", StringComparison.OrdinalIgnoreCase)) return "";

            return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(c.ToLowerInvariant());
        }

        /// <summary>
        /// Extrae únicamente los caracteres numéricos de una cadena (Fase 31).
        /// </summary>
        private static string SoloDigitos(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return "";
            StringBuilder sb = new StringBuilder(valor.Length);
            foreach (char c in valor)
            {
                if (char.IsDigit(c)) sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Construye el nombre del PDF unificado utilizando la taxonomía {FechaPago}_{NroOB}_{Categoria}_{TerceroSanitizado} (Fases 20-55, 77).
        /// Si la categoría es "XXX" o vacía, o si reduce el presupuesto del tercero a menos de MIN_TERCERO_VISIBLE (8 chars), la omite del nombre (Fases 52, 55, 77).
        /// Garantiza una cota máxima estricta de 30 caracteres (MAX_LARGO_NOMBRE) en el nombre base del archivo de salida.
        /// Nota de arquitectura (Fase 77): Se elimina la prevención de colisiones mediante sufijos hash deterministas.
        /// La cadena completa de salida se concatena primero ({FechaPago}_{NroOB}_{Categoria}_{Tercero}) y posteriormente se aplica
        /// un truncamiento directo mediante String.Substring(0, maxLargo) si la longitud global supera los 30 caracteres.
        /// </summary>
        /// <param name="fechaPago">Fecha de pago extraída u obtenida del documento.</param>
        /// <param name="obligacion">Número o clave de la obligación (NroOB).</param>
        /// <param name="categoria">Categoría del pago (3 letras) o vacía/XXX.</param>
        /// <param name="tercero">Nombre original del tercero beneficiario.</param>
        /// <param name="maxLargo">Límite estricto global de caracteres (por defecto 30).</param>
        /// <returns>Nombre sanitizado del archivo de salida sin extensión, recortado a un máximo de 30 caracteres sin sufijos hash.</returns>
        internal static string ConstruirNombreSalida(string fechaPago, string obligacion, string categoria,
                                                     string tercero, int maxLargo = MAX_LARGO_NOMBRE)
        {
            int limite = maxLargo < 20 ? 20 : maxLargo;

            string fecha = NormalizarFecha(fechaPago);
            string ob = SoloDigitos(obligacion);
            if (ob.Length == 0) ob = "0";

            string cat = FormatearCategoria(categoria);
            string terCompleto = string.IsNullOrWhiteSpace(tercero)
                ? "Sintercero"
                : FirmaPDFsHelper.SanitizarNombreCompleto(tercero);

            string prefijoConCat = string.Concat(fecha, "_", ob, "_", string.IsNullOrEmpty(cat) ? "" : cat + "_");
            string prefijoSinCat = string.Concat(fecha, "_", ob, "_");

            int dispConCat = limite - prefijoConCat.Length;
            int dispSinCat = limite - prefijoSinCat.Length;

            bool omitirCat = !string.IsNullOrEmpty(cat) && (
                dispConCat < MIN_TERCERO_VISIBLE ||
                (terCompleto.Length > dispConCat && terCompleto.Length <= dispSinCat)
            );

            string prefijo = omitirCat ? prefijoSinCat : prefijoConCat;
            string textoCompleto = prefijo + terCompleto;

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                textoCompleto = textoCompleto.Replace(c.ToString(), "_");
            }

            // Fase 77: Recorte directo sobre la cadena concatenada completa a maxLargo (30 caracteres) sin sufijo hash.
            string textoRecortado = textoCompleto.Length > limite
                ? textoCompleto.Substring(0, limite)
                : textoCompleto;

            if (textoRecortado.EndsWith(".", StringComparison.Ordinal))
            {
                textoRecortado = textoRecortado.TrimEnd('_', '.', ' ');
            }

            return textoRecortado;
        }


        /// <summary>
        /// Resuelve la categoría de 3 letras: intenta primero desde el 1er soporte de la hoja con categoría válida;
        /// si no hay, busca el rótulo "Objeto:" en la hoja (Decisión D-4); si tampoco, asigna "XXX".
        /// </summary>
        internal static string ResolverCategoria(
            List<string> soportesHoja,
            Dictionary<string, OrdenPagoNombreSoportePdf> soportesParsed,
            dynamic ws)
        {
            if (soportesHoja != null && soportesParsed != null)
            {
                foreach (string ruta in soportesHoja)
                {
                    OrdenPagoNombreSoportePdf datos;
                    if (soportesParsed.TryGetValue(ruta, out datos) && datos != null && datos.TieneCategoria)
                    {
                        return datos.Categoria;
                    }
                }
            }

            string objeto = OrdenPagoSheetReader.BuscarObjeto(ws);
            if (string.IsNullOrEmpty(objeto)) return "XXX";

            string cat = FirmaPDFsHelper.GetCategoria(objeto);
            return string.IsNullOrEmpty(cat) ? "XXX" : cat;
        }

        /// <summary>
        /// Localiza el boletín en la lista disponible aplicando la estrategia de correspondencia (Decisión D-8).
        /// Incluye logs de observabilidad (Hallazgo C4) y tokenización numérica exacta (Fase 13).
        /// </summary>
        private static string BuscarBoletin(List<string> boletines, OrdenPagoDatosHoja datos, string docType, string clave)
        {
            if (string.IsNullOrEmpty(clave)) return null;

            List<string> candidatos = new List<string>();
            foreach (string b in boletines)
            {
                string nombreBase = Path.GetFileNameWithoutExtension(b);
                if (ContieneOEsIgualNumero(nombreBase, clave))
                {
                    candidatos.Add(b);
                }
            }

            if (candidatos.Count == 0)
            {
                LoggerHelper.Warning($"No hay boletín que contenga {docType}={clave} en Boletines\\ (hoja '{datos.NombreHoja}').");
                return null;
            }
            if (candidatos.Count == 1) return candidatos[0];

            // Desempate por la otra clave
            string otra = string.Equals(docType, "OB", StringComparison.OrdinalIgnoreCase) ? datos.Orden : datos.Obligacion;
            if (!string.IsNullOrEmpty(otra))
            {
                foreach (string c in candidatos)
                {
                    string nombreBase = Path.GetFileNameWithoutExtension(c);
                    if (ContieneOEsIgualNumero(nombreBase, otra))
                    {
                        LoggerHelper.Info($"Boletín desempatado por {otra}: {Path.GetFileName(c)}");
                        return c;
                    }
                }
            }

            LoggerHelper.Warning($"Ambigüedad de boletín para {docType}={clave}: se usa el primero ({Path.GetFileName(candidatos[0])}). Candidatos={candidatos.Count}");
            return candidatos[0];
        }

        /// <summary>
        /// Fase 13: Evalúa si el texto contiene un token numérico exacto coincidente con el número buscado.
        /// </summary>
        private static bool ContieneOEsIgualNumero(string texto, string numero)
        {
            if (string.IsNullOrEmpty(texto) || string.IsNullOrEmpty(numero)) return false;

            foreach (Match m in Regex.Matches(texto, @"\d+"))
            {
                if (NumerosIguales(m.Value, numero)) return true;
            }
            return false;
        }

        /// <summary>
        /// Traslada un boletín unificado a la carpeta de histórico de Ordenes de Pago.
        /// </summary>
        private static void MoverAHistorico(string origen, string destino)
        {
            try
            {
                if (File.Exists(destino))
                {
                    LoggerHelper.Warning($"Ya existía en OrdenesDePago, se reemplaza: {Path.GetFileName(destino)}");
                    File.Delete(destino);
                }
                File.Move(origen, destino);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"No se pudo mover a OrdenesDePago: {Path.GetFileName(origen)}", ex);
            }
        }

        /// <summary>
        /// Determina si un archivo se encuentra dentro de un directorio específico.
        /// </summary>
        private static bool EstaEnCarpeta(string rutaArchivo, string carpeta)
        {
            try
            {
                string normal = Path.GetFullPath(rutaArchivo).ToUpperInvariant();
                string raiz = Path.GetFullPath(carpeta).TrimEnd('\\').ToUpperInvariant();
                return normal.StartsWith(raiz + "\\", StringComparison.Ordinal);
            }
            catch { return false; }
        }
    }
}
