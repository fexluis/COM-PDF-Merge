using System;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio para la extracción de datos de documentos para firma.
    /// </summary>
    public class FirmaExtractService : IFirmaExtractService
    {
        /// <inheritdoc/>
        public ExtraerDatos ObligacionData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xDep = ws.Range["B32"].Value?.ToString();

            if (ws.Range["F16"].Value?.ToString() == "Estado:")
            {
                extract.xNumero = ws.Range["C15"].Value?.ToString() ?? "";
                extract.xFechaStr = ws.Range["I15"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["I22"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["C22"].Value?.ToString() ?? "";
                extract.xRP = ws.Range["AE16"].Value?.ToString() ?? "";
                extract.xObjeto = ws.Range["C38"].Value?.ToString() ?? "";
                extract.xCategoria = GetCategoria(extract.xObjeto);
            }
            else
            {
                extract.xNumero = ws.Range["D15"].Value?.ToString() ?? "";
                extract.xFechaStr = ws.Range["K15"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["K22"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["D22"].Value?.ToString() ?? "";
                extract.xRP = ws.Range["AN16"].Value?.ToString() ?? "";

                dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                if (rngObjeto != null)
                {
                    extract.xObjeto = ws.Range["D" + rngObjeto.Row].Value?.ToString() ?? "";
                    extract.xCategoria = GetCategoria(extract.xObjeto);
                }
                else
                {
                    extract.xObjeto = "DESCONOCIDO";
                    extract.xCategoria = "XXX";
                }
            }

            extract.xUnifiedFileName = $"OB-{extract.xNumero}-{extract.xCategoria}-{SanitizarNombre(extract.xTercero)}";
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos OrdeDePagoData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            string prefijo = "";
            string nombreOrden = "";

            if (ws.Range["X16"].Value != null)
            {
                extract.xFechaStr = ws.Range["J15"].Value?.ToString() ?? "";
                extract.xOB = ws.Range["X16"].Value?.ToString() ?? "";
                extract.xOP = ws.Range["B15"].Value?.ToString() ?? "";
                extract.xEstado = ws.Range["J16"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["J30"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["B30"].Value?.ToString() ?? "";
                
                extract.xFechaStr = extract.xFechaStr.Replace("-", "");
                extract.xTercero = (extract.xEstado == "Anulada") ? "Anulada" : extract.xTercero;
                extract.xUnifiedFileName = extract.xFechaStr + "_" + extract.xOP + "_" + extract.xOB + "_";
            }
            else
            {
                extract.xFechaStr = ws.Range["I15"].Value?.ToString() ?? "";
                extract.xOB = ws.Range["U16"].Value?.ToString() ?? "";
                extract.xOP = ws.Range["B15"].Value?.ToString() ?? "";
                extract.xEstado = ws.Range["I16"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["B30"].Value?.ToString() ?? "";
                
                extract.xFechaStr = extract.xFechaStr.Replace("-", "");

                extract.xTercero = string.IsNullOrEmpty(ws.Range["I30"].Value?.ToString()) ? ws.Range["I31"].Value?.ToString() : ws.Range["I30"].Value?.ToString();
                extract.xTercero = (extract.xEstado == "Anulada") ? "Anulada" : extract.xTercero;
                extract.xUnifiedFileName = extract.xFechaStr + "_" + extract.xOP + "_" + extract.xOB + "_";
            }

            extract.xUnifiedFileName = SanitizarNombre(extract.xUnifiedFileName);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos CdpData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["E18"].Value?.ToString() ?? "";
            extract.xCDP = ws.Range["B18"].Value?.ToString() ?? "";
            extract.xNumero = extract.xCDP;
            
            string prefijo = extract.xFechaStr.Replace("-", "") + "_CDP_" + extract.xCDP + "_";
            extract.xUnifiedFileName = SanitizarNombre(prefijo);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos RpData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["I16"].Value?.ToString() ?? "";
            extract.xRP = ws.Range["F16"].Value?.ToString() ?? "";
            extract.xTercero = ws.Range["I20"].Value?.ToString() ?? "";
            extract.xNumero = extract.xRP;

            string prefijo = extract.xFechaStr.Replace("-", "") + "_RP_" + extract.xRP + "_";
            extract.xUnifiedFileName = SanitizarNombre(prefijo + extract.xTercero);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos VerificadosData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["N15"].Value?.ToString() ?? "";
            extract.xNumero = ws.Range["C15"].Value?.ToString() ?? "";
            extract.xTercero = ws.Range["A23"].Value?.ToString() ?? "";

            string prefijo = extract.xFechaStr.Replace("-", "") + "_V_" + extract.xNumero + "_";
            extract.xUnifiedFileName = SanitizarNombre(prefijo + extract.xTercero);
            return extract;
        }

        /// <summary>
        /// Sanitiza el nombre de un archivo para que sea válido y siga un formato estandarizado. (Ported from FirmaPDFsHelper)
        /// </summary>
        private string SanitizarNombre(string nombre)
        {
            if (string.IsNullOrEmpty(nombre)) return "Archivo_" + DateTime.Now.Ticks;
            string n = nombre.Replace("  ", " ").Replace("-", " ");
            n = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(n.ToLower()).Replace(" ", "");
            var partes = n.Split('_');
            for (int i = 0; i < partes.Length; i++) if (partes[i] == "") partes[i] = "DESCONOCIDO";
            n = string.Join("_", partes);
            return n.Length > 30 ? n.Substring(0, 30) : n;
        }

        /// <summary>
        /// Determina la categoría del documento basándose en palabras clave del objeto.
        /// </summary>
        /// <param name="objeto">Texto del objeto de la obligación.</param>
        /// <returns>Código corto de categoría (VIA, HON, SSO, NOM, etc.).</returns>
        private string GetCategoria(string objeto)
        {
            if (string.IsNullOrEmpty(objeto)) return "XXX";
            string t = objeto.ToUpperInvariant();
            if (t.Contains("VIAT")) return "VIA";
            if (t.Contains("ARL") || t.Contains("RIESGO")) return "ARL";
            if (t.Contains("HONOR")) return "HON";
            if (t.Contains("SEGURIDA") || t.Contains("SOCIAL")) return "SSO";
            if (t.Contains("NOMINA") || t.Contains("NÓMINA")) return "NOM";
            if (t.Contains("CESANT")) return "FNA";
            if (t.Contains("ARREND") || t.Contains("ARRIEND")) return "ARR";
            if (t.Contains("SERVICIO") || t.Contains("PUBLICO") || t.Contains("PÚBLICO")) return "SPU";
            if (t.Contains("ESTUDIANTE") || t.Contains("MONITOR") || t.Contains("PRACTICA") || t.Contains("ESTIMULO")) return "EST";
            if (t.Contains("4X") || t.Contains("4 X") || t.Contains("X1000") || t.Contains("X 1000") || t.Contains("POR MIL")) return "4PM";
            return "XXX";
        }
    }
}
