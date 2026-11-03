using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Interpretación del texto extraído de un PDF: tipo de informe y verificación de la
    /// cédula del comisionado con distintos niveles de confianza.
    /// </summary>
    public static class InformeTextoHelper
    {
        private const string PrefijoErrorPdf = "ERROR:";

        /// <summary>
        /// Etiquetas habituales que anteceden a la cédula en los formatos de la entidad.
        /// Se evalúan sobre texto normalizado (sin tildes, minúsculas).
        /// </summary>
        private const string EtiquetasCedula =
            @"(no\.?\s*identificacion(\s*persona)?|identificacion(\s*persona)?|" +
            @"cedula(\s*de\s*ciudadania)?|c\.?\s*c\.?|documento(\s*de\s*identidad)?|" +
            @"nro\.?\s*identificacion|numero\s*de\s*identificacion|nro\.?|numero)\s*[:#\-]?";

        private static readonly Regex RegexCedulaCorreo =
            new Regex(@"(?<!\d)(\d[\d\.,\u00A0\t ]{5,})\s+[a-z0-9._%+-]+@", RegexOptions.Compiled);

        private static readonly Regex RegexEtiqueta =
            new Regex(EtiquetasCedula, RegexOptions.Compiled);

        private static readonly Regex RegexNumero =
            new Regex(@"(?<![\d\.,])(\d[\d\.,\u00A0\t ]{3,18}\d)(?![\d\.,])", RegexOptions.Compiled);

        private static readonly Regex RegexCertificadoCumplimiento =
            new Regex(@"certificado\s+de\s+cumplimiento", RegexOptions.Compiled);

        private static readonly Regex RegexInformeCumplimiento =
            new Regex(@"informe\s+de\s+cumplimiento", RegexOptions.Compiled);

        /// <summary>
        /// Normaliza un texto para comparaciones: minúsculas, sin tildes y con espacios colapsados.
        /// </summary>
        /// <param name="texto">Texto original.</param>
        /// <returns>Texto normalizado.</returns>
        public static string Normalizar(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return string.Empty;

            string sinAcentos = InformeNombreHelper.QuitarDiacriticos(texto).ToLowerInvariant();
            sinAcentos = sinAcentos.Replace('\u00A0', ' ');
            return Regex.Replace(sinAcentos, @"[^\S\r\n]+", " ");
        }

        /// <summary>
        /// Determina el tipo de informe a partir del texto del PDF.
        /// Devuelve <see cref="InformeTipo.Err"/> cuando no puede determinarse o es ambiguo
        /// (fail-closed: nunca devuelve XXX).
        /// </summary>
        /// <param name="texto">Texto extraído del PDF.</param>
        /// <returns>INF, CUM o ERR.</returns>
        public static string DetectarTipo(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return InformeTipo.Err;
            if (texto.TrimStart().StartsWith(PrefijoErrorPdf, StringComparison.OrdinalIgnoreCase)) return InformeTipo.Err;

            string t = Normalizar(texto);

            bool certificado = t.Contains("certificado");
            bool informe = t.Contains("informe");
            bool cumplimiento = t.Contains("cumplimiento");

            if (certificado && cumplimiento && !informe) return InformeTipo.Cum;
            if (informe && cumplimiento && !certificado) return InformeTipo.Inf;

            if (certificado && informe && cumplimiento)
            {
                // Ambas palabras presentes: se resuelve por frase exacta; si no, es ambiguo.
                if (RegexCertificadoCumplimiento.IsMatch(t)) return InformeTipo.Cum;
                if (RegexInformeCumplimiento.IsMatch(t)) return InformeTipo.Inf;
            }

            return InformeTipo.Err;
        }

        /// <summary>
        /// Busca y verifica la cédula del comisionado dentro del texto del PDF.
        /// Estrategias en orden de precisión:
        /// E1 cédula + correo (formato GF-FO-023), E2 etiqueta en la misma línea,
        /// E3 coincidencia exacta y única con las cédulas de CONSULTAS y
        /// E4 heurística libre (confianza baja, nunca valida por sí sola).
        /// </summary>
        /// <param name="texto">Texto extraído del PDF.</param>
        /// <param name="cedulasCandidatas">Cédulas de CONSULTAS para la solicitud del archivo.</param>
        /// <returns>Resultado con cédula, confianza y evidencia.</returns>
        public static ResultadoCedulaPdf BuscarCedula(string texto, IEnumerable<string> cedulasCandidatas)
        {
            if (string.IsNullOrWhiteSpace(texto)) return ResultadoCedulaPdf.SinHallazgo(MotivoInforme.SinTexto);
            if (texto.TrimStart().StartsWith(PrefijoErrorPdf, StringComparison.OrdinalIgnoreCase))
                return ResultadoCedulaPdf.SinHallazgo(MotivoInforme.SinTexto);

            var candidatas = new HashSet<string>();
            if (cedulasCandidatas != null)
            {
                foreach (string c in cedulasCandidatas)
                {
                    string normalizada = InformeNombreHelper.NormalizarDocumento(c);
                    if (normalizada.Length >= 5) candidatas.Add(normalizada);
                }
            }

            string normalizado = Normalizar(texto);

            // E1: cédula seguida de correo electrónico.
            foreach (Match m in RegexCedulaCorreo.Matches(normalizado))
            {
                string hallada = InformeNombreHelper.NormalizarDocumento(m.Groups[1].Value);
                if (candidatas.Contains(hallada))
                    return ResultadoCedulaPdf.Hallazgo(hallada, ConfianzaCedula.Alta, "cedula+correo");
            }

            // E2: etiqueta y número en la misma línea.
            var porEtiqueta = new HashSet<string>();
            foreach (string linea in normalizado.Split('\n'))
            {
                if (linea.Length == 0 || !RegexEtiqueta.IsMatch(linea)) continue;

                foreach (string numero in ExtraerNumeros(linea))
                {
                    if (candidatas.Count == 0 || candidatas.Contains(numero)) porEtiqueta.Add(numero);
                }
            }

            if (porEtiqueta.Count == 1)
                return ResultadoCedulaPdf.Hallazgo(porEtiqueta.First(), ConfianzaCedula.Alta, "etiqueta identificacion");
            if (porEtiqueta.Count > 1)
                return ResultadoCedulaPdf.SinHallazgo(MotivoInforme.CedulaAmbigua);

            // E3: coincidencia exacta y única contra CONSULTAS.
            var exactas = new HashSet<string>();
            foreach (string numero in ExtraerNumeros(normalizado))
            {
                if (candidatas.Contains(numero)) exactas.Add(numero);
            }

            if (exactas.Count == 1)
                return ResultadoCedulaPdf.Hallazgo(exactas.First(), ConfianzaCedula.Media, "coincidencia exacta CONSULTAS");
            if (exactas.Count > 1)
                return ResultadoCedulaPdf.SinHallazgo(MotivoInforme.CedulaAmbigua);

            // E4: heurística libre, solo si no hay candidatas contra las que comparar.
            if (candidatas.Count == 0)
            {
                foreach (string numero in ExtraerNumeros(normalizado))
                {
                    if (EsCedulaPlausible(numero))
                        return ResultadoCedulaPdf.Hallazgo(numero, ConfianzaCedula.Baja, "heuristica libre");
                }
            }

            return ResultadoCedulaPdf.SinHallazgo(
                candidatas.Count > 0 ? MotivoInforme.CedulaNoHallada : MotivoInforme.SinTexto);
        }

        /// <summary>
        /// Extrae los números candidatos (6-11 dígitos) de un texto ya normalizado.
        /// </summary>
        /// <param name="texto">Texto normalizado.</param>
        /// <returns>Números en solo dígitos.</returns>
        internal static IEnumerable<string> ExtraerNumeros(string texto)
        {
            if (string.IsNullOrEmpty(texto)) yield break;

            foreach (Match m in RegexNumero.Matches(texto))
            {
                string numero = InformeNombreHelper.NormalizarDocumento(m.Groups[1].Value);
                if (numero.Length >= 6 && numero.Length <= 11) yield return numero;
            }
        }

        /// <summary>
        /// Descarta números que no pueden ser cédulas (años, celulares, ceros).
        /// </summary>
        /// <param name="cedula">Número en solo dígitos.</param>
        /// <returns>True si el número es una cédula plausible.</returns>
        internal static bool EsCedulaPlausible(string cedula)
        {
            if (string.IsNullOrEmpty(cedula)) return false;
            if (cedula.Length < 6 || cedula.Length > 10) return false;
            if (cedula[0] == '3') return false;
            if (cedula.StartsWith("19", StringComparison.Ordinal) || cedula.StartsWith("20", StringComparison.Ordinal)) return false;

            bool todosCeros = true;
            for (int i = 0; i < cedula.Length; i++)
            {
                if (cedula[i] != '0') { todosCeros = false; break; }
            }
            return !todosCeros;
        }
    }
}
