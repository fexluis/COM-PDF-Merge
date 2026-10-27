using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Analizador sintáctico (Parser) basado en expresiones regulares para extraer información
    /// estructurada de los nombres de archivos PDF de soporte de Orden de Pago.
    /// Paridad funcional con la lógica de coincidencia del VBA original.
    /// </summary>
    internal static class OrdenPagoNombreSoportePdfParser
    {
        /// <summary>
        /// Expresión regular para el patrón completo: [prefijos_opcionales]orden_OB-obligacion-categoria-tercero
        /// </summary>
        private static readonly Regex PatronConCategoria = new Regex(
            @"^(?:[A-Za-z0-9]+[-_])*(?<orden>\d{4,})[-_]OB[-_](?<obligacion>\d+)[-_](?<categoria>[A-Za-z]{3})[-_](?<tercero>.+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Expresión regular para el patrón fallback: [prefijos_opcionales]orden_OB-obligacion-tercero
        /// </summary>
        private static readonly Regex PatronSinCategoria = new Regex(
            @"^(?:[A-Za-z0-9]+[-_])*(?<orden>\d{4,})[-_]OB[-_](?<obligacion>\d+)[-_](?<tercero>.+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Catálogo oficial de 3 letras de categorías válidas (alineado con FirmaPDFsHelper.GetCategoria).
        /// </summary>
        private static readonly HashSet<string> CatalogoCategorias =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "XXX", "VIA", "ARL", "HON", "SSO", "NOM", "FNA", "ARR", "SPU", "EST", "4PM" };

        /// <summary>
        /// Intenta parsear el nombre de un archivo PDF de soporte y extraer sus componentes.
        /// </summary>
        /// <param name="nombreArchivo">Nombre o ruta del archivo de soporte.</param>
        /// <param name="datos">Objeto <see cref="OrdenPagoNombreSoportePdf"/> resultante si el parseo fue exitoso; null en caso contrario.</param>
        /// <returns>True si el nombre coincide con la estructura de un soporte reconocido; False en caso contrario.</returns>
        internal static bool TryParse(string nombreArchivo, out OrdenPagoNombreSoportePdf datos)
        {
            datos = null;
            if (string.IsNullOrWhiteSpace(nombreArchivo))
            {
                return false;
            }

            string baseSinExt = Path.GetFileNameWithoutExtension(nombreArchivo.Trim());
            if (string.IsNullOrEmpty(baseSinExt))
            {
                return false;
            }

            Match m = PatronConCategoria.Match(baseSinExt);
            if (m.Success && CatalogoCategorias.Contains(m.Groups["categoria"].Value))
            {
                datos = new OrdenPagoNombreSoportePdf
                {
                    Orden = m.Groups["orden"].Value,
                    Obligacion = m.Groups["obligacion"].Value,
                    Categoria = m.Groups["categoria"].Value.ToUpperInvariant(),
                    Tercero = m.Groups["tercero"].Value
                };
                return true;
            }

            m = PatronSinCategoria.Match(baseSinExt);
            if (m.Success)
            {
                datos = new OrdenPagoNombreSoportePdf
                {
                    Orden = m.Groups["orden"].Value,
                    Obligacion = m.Groups["obligacion"].Value,
                    Categoria = null,
                    Tercero = m.Groups["tercero"].Value
                };
                return true;
            }

            return false;
        }
    }
}
