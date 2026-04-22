using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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
        public void RenombrarSolicitudes(string folder)
        {
            try
            {
                if (!Directory.Exists(folder)) return;
                
                // Obtiene la lista de comisiones desde la hoja CONSULTAS para tener los metadatos de referencia
                var solicitudes = GetSolicitudes();
                if (solicitudes.Count == 0) return;

                // Normaliza los nombres de archivos actuales (quita espacios, guiones bajos, etc.)
                Estandarizar(folder);
                
                int count = 0;
                // Procesa cada PDF en la carpeta para intentar identificarlo y renombrarlo
                foreach (string file in Directory.GetFiles(folder, "*.pdf"))
                {
                    if (RenombrarFila(file, solicitudes, folder)) count++;
                }
                
                MessageBox.Show($"Finalizado. {count} archivos renombrados.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { LoggerHelper.Error("RenombrarSolicitudes", ex); }
        }

        /// <inheritdoc/>
        public int UnificarSolicitudes(string folder)
        {
            try
            {
                if (!Directory.Exists(folder)) return 0;
                
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
                    
                    // Lógica de unificación por hoja
                    if (Unificar(ws, pSoli, pAuto, pRP, pCDP, pVeri, pUnif)) count++;
                }
                
                // Abre la carpeta de resultados si hubo éxito
                if (count > 0) System.Diagnostics.Process.Start("explorer.exe", pUnif);
                return count;
            }
            catch (Exception ex) { LoggerHelper.Error("UnificarSolicitudes", ex); return 0; }
        }

        /// <summary>
        /// Obtiene la información básica de comisiones (Número, CC, Nombre) desde la hoja CONSULTAS.
        /// </summary>
        private List<SInfo> GetSolicitudes()
        {
            var list = new List<SInfo>();
            object[,] data = _excelRepository.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS);
            if (data == null) return list;
            
            int rows = data.GetLength(0);
            for (int i = 2; i <= rows; i++)
            {
                string num = data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(num))
                {
                    list.Add(new SInfo { 
                        Num = num, 
                        CC = data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString() ?? "",
                        Nom = Clean(data[i, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO]?.ToString() ?? "")
                    });
                }
            }
            return list;
        }

        /// <summary>
        /// Limpia el nombre del funcionario para usarlo en nombres de archivos (PascalCase sin espacios).
        /// </summary>
        private string Clean(string s) => Regex.Replace(App.CapitalizarPalabras(s), @"\s+", "").Substring(0, Math.Min(s.Length, 18));

        /// <summary>
        /// Normaliza los nombres de archivos físicos reemplazando espacios y guiones bajos por un guion único.
        /// </summary>
        /// <summary>
        /// Normaliza los nombres de todos los archivos en un directorio reemplazando espacios y guiones bajos por un guion único.
        /// Previne problemas de búsqueda por discrepancias mínimas en el nombre.
        /// </summary>
        /// <param name="f">Ruta del directorio a estandarizar.</param>
        private void Estandarizar(string f)
        {
            foreach (string file in Directory.GetFiles(f))
            {
                string old = Path.GetFileName(file);
                string newN = Regex.Replace(old.Replace("_", "-").Replace(" ", "-"), @"-+", "-");
                if (old != newN) try { File.Move(file, Path.Combine(f, newN)); } catch { }
            }
        }

        /// <summary>
        /// Intenta renombrar un archivo individual basándose en el número de comisión detectado.
        /// Si hay duplicidad, recurre al OCR para extraer la cédula y confirmar la identidad.
        /// </summary>
        /// <summary>
        /// Intenta identificar y renombrar un archivo PDF basándose en el número de comisión.
        /// Utiliza una estrategia de tres niveles: identificación única, lectura de metadatos o OCR (FindCC).
        /// </summary>
        /// <param name="file">Ruta completa del archivo.</param>
        /// <param name="sols">Lista de solicitudes de referencia.</param>
        /// <param name="f">Carpeta contenedora.</param>
        /// <returns>True si el archivo fue renombrado.</returns>
        private bool RenombrarFila(string file, List<SInfo> sols, string f)
        {
            string oldN = Path.GetFileNameWithoutExtension(file);
            string ext = Path.GetExtension(file);
            
            // Se asume que el nombre del archivo empieza con el número de comisión (ej: 123-...)
            string num = oldN.Split('-')[0];
            var matches = sols.Where(s => s.Num == num).ToList();
            if (matches.Count == 0) return false;

            // OPTIMIZACIÓN: Intentar identificar por metadatos antes que usar OCR (FindCC)
            SInfo m = null;
            try
            {
                var meta = _metadataService.LeerMetadatos(file);
                if (meta != null && meta.MetaDocumento == TipoMetadata.VIA_SOLICITUD && !string.IsNullOrEmpty(meta.PalabrasClave))
                {
                    // Obtiene la CC que es el primer término separado por comas
                    string metaCC = meta.PalabrasClave.Split(',')[0].Trim();
                    m = matches.FirstOrDefault(s => s.CC == metaCC);
                }
            }
            catch { /* Ignorar errores de lectura de metadatos y seguir con el flujo normal */ }

            // Si no se identificó por metadatos, proceder con la lógica estándar (conteo único o OCR)
            if (m == null)
            {
                m = matches.Count == 1 ? matches[0] : matches.FirstOrDefault(s => s.CC == FindCC(file));
            }
            
            if (m != null)
            {
                string newN = $"{m.Num}-{m.CC}-{m.Nom}{ext}";
                if (oldN + ext != newN)
                {
                    string p = Path.Combine(f, newN);
                    if (!File.Exists(p)) { File.Move(file, p); return true; }
                }
            }
            return false;
        }

        /// <summary>
        /// Extrae el texto de la primera página del PDF y busca patrones de cédula de ciudadanía.
        /// </summary>
        /// <summary>
        /// Extrae el texto de la primera página del PDF y busca patrones de cédula de ciudadanía mediante Regex.
        /// </summary>
        /// <param name="f">Ruta del archivo PDF.</param>
        /// <returns>La cédula encontrada o una cadena vacía.</returns>
        private string FindCC(string f)
        {
            string t = _pdfToTexto.ExtraerTextoPagina(f, 1);
            
            // Patrón 1: Cédula seguida de un correo electrónico (común en formatos institucionales)
            var m = Regex.Match(t, @"(\d[\d\.,\s]*)\s+[a-zA-Z0-9._%+-]+@");
            if (m.Success) return Regex.Replace(m.Groups[1].Value, @"[\.,\s]", "");
            
            // Patrón 2: Fallback para cualquier secuencia numérica de 7 a 11 dígitos
            var m2 = Regex.Match(t, @"\b\d{7,11}\b");
            return m2.Success ? m2.Value : "";
        }

        /// <summary>
        /// Lógica central de unificación: recolecta las rutas de todos los componentes del expediente (Soli, Auto, RP, CDP, Veri)
        /// y genera un PDF unificado usando PdfSharp.
        /// </summary>
        /// <summary>
        /// Orquesta la unión de múltiples PDFs para formar un expediente unificado.
        /// Recolecta Solicitud, Resolución, RP, CDP y Cumplimiento basándose en los datos de la hoja.
        /// </summary>
        /// <param name="ws">Hoja de cálculo individual del comisionado.</param>
        /// <param name="s">Carpeta de Solicitudes.</param>
        /// <param name="a">Carpeta de Autorizaciones.</param>
        /// <param name="r">Carpeta de RPs.</param>
        /// <param name="c">Carpeta de CDPs.</param>
        /// <param name="v">Carpeta de Verificados.</param>
        /// <param name="outD">Carpeta de salida para el PDF unificado.</param>
        /// <returns>True si el expediente se unificó correctamente.</returns>
        private bool Unificar(dynamic ws, string s, string a, string r, string c, string v, string outD)
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
        /// </summary>
        /// <summary>
        /// Busca el primer archivo PDF que contenga el patrón especificado dentro de una carpeta.
        /// Útil para encontrar archivos por ID (ej. número de comisión o CC).
        /// </summary>
        /// <param name="d">Directorio de búsqueda.</param>
        /// <param name="p">Patrón o prefijo a buscar.</param>
        /// <returns>La ruta completa del primer archivo encontrado o vacío.</returns>
        private string Buscar(string d, string p) => (!string.IsNullOrEmpty(p) && Directory.Exists(d)) ? Directory.GetFiles(d, "*.pdf").FirstOrDefault(f => f.Contains(p)) ?? "" : "";

        /// <summary>
        /// DTO interno para manejar la información de solicitudes detectadas.
        /// </summary>
        /// <summary>
        /// DTO interno para manejar la información de identificación de solicitudes.
        /// </summary>
        private class SInfo { public string Num, CC, Nom; }
    }
}