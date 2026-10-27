using System;
using System.Globalization;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Utilidad para la inspección y extracción de datos desde las hojas de cálculo de Excel
    /// correspondientes a Ordenes de Pago Presupuestales.
    /// Maneja de forma tolerante la localización dinámica de celdas por rótulo (Q16/X16 vs O16/U16).
    /// </summary>
    internal static class OrdenPagoSheetReader
    {
        internal const string TITULO_HOJA_OP = "ORDEN DE PAGO PRESUPUESTAL";
        private const string ROTULO_OBLIGACION = "Nro Obligación:";
        private const string ROTULO_TERCERO = "Razón Social:";
        private static readonly string[] ROTULOS_FECHA =
            { "Fecha Registro:", "Fecha:", "Fecha de Pago:", "Fecha Pago:", "Fecha Radicación:" };

        private static readonly string[] FormatosFecha =
            { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd-MM-yyyy", "d-M-yyyy", "yyyyMMdd" };

        private const int SERIAL_OADATE_MIN = 36526;   // 2000-01-01
        private const int SERIAL_OADATE_MAX = 2958466; // 9999-12-31

        private static bool EsRotuloFecha(string valorCelda)
        {
            if (string.IsNullOrEmpty(valorCelda)) return false;
            foreach (string r in ROTULOS_FECHA)
            {
                if (RotuloCoincide(valorCelda, r)) return true;
            }
            return false;
        }

        /// <summary>
        /// Lee una fecha en formato texto plano tolerando seriales OADate enteros y los formatos de fecha comunes (Fases 33, 37, 42).
        /// </summary>
        internal static string LeerFechaPagoDesdeTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return "";
            string s = texto.Trim();

            int oaParsed;
            if (int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out oaParsed) &&
                oaParsed >= SERIAL_OADATE_MIN && oaParsed < SERIAL_OADATE_MAX)
            {
                return DateTime.FromOADate(oaParsed).ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            }

            DateTime f;
            if (DateTime.TryParseExact(s, FormatosFecha, CultureInfo.GetCultureInfo("es-CO"),
                                       DateTimeStyles.None, out f))
            {
                return f.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            }

            if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out f))
            {
                return f.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            }

            return "";
        }

        /// <summary>
        /// Lee una fecha de celda COM tolerando DateTime, serial OADate (rango 36526..2958466) y texto (Fases 21, 24, 30, 33, 37, 42).
        /// </summary>
        internal static string LeerFechaPago(dynamic celda)
        {
            try
            {
                if (celda == null) return "";
                object v = celda.Value;
                if (v == null) return "";

                if (v is DateTime)
                {
                    DateTime dtValue = (DateTime)v;
                    return dtValue.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                }

                if (v is double)
                {
                    double oa = (double)v;
                    if (oa >= SERIAL_OADATE_MIN && oa < SERIAL_OADATE_MAX)
                    {
                        return DateTime.FromOADate(oa).ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                    }
                }

                string s = Convert.ToString(v).Trim();
                if (string.IsNullOrEmpty(s)) return "";

                return LeerFechaPagoDesdeTexto(s);
            }
            catch { return ""; }
        }

        /// <summary>
        /// Convierte el valor de una celda COM de Excel a cadena limpia sin espacios laterales.
        /// Evita excepciones nulas.
        /// </summary>
        internal static string Texto(dynamic celda)
        {
            try
            {
                if (celda == null) return "";
                object v = celda.Value;
                if (v == null) return "";
                string s = Convert.ToString(v);
                return s == null ? "" : s.Trim();
            }
            catch { return ""; }
        }

        /// <summary>
        /// Compara de forma tolerante un texto de celda contra un rótulo esperado, eliminando ':' finales y espacios.
        /// </summary>
        internal static bool RotuloCoincide(string valorCelda, string rotulo)
        {
            if (string.IsNullOrEmpty(valorCelda) || string.IsNullOrEmpty(rotulo)) return false;
            string a = valorCelda.Trim().TrimEnd(':').Trim();
            string b = rotulo.Trim().TrimEnd(':').Trim();
            return a.Equals(b, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Evalúa si la hoja de cálculo activa es una hoja válida de Orden de Pago verificando la celda A14.
        /// </summary>
        internal static bool EsHojaOrdenPago(dynamic ws)
        {
            try
            {
                string a14 = Texto(ws.Range["A14"]);
                return a14.Equals(TITULO_HOJA_OP, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>
        /// Extrae la información estructurada de la hoja de cálculo de Orden de Pago.
        /// </summary>
        /// <param name="ws">Objeto COM Worksheet de Excel.</param>
        /// <returns>Instancia de <see cref="OrdenPagoDatosHoja"/> con los datos leídos.</returns>
        internal static OrdenPagoDatosHoja LeerHojaOp(dynamic ws)
        {
            OrdenPagoDatosHoja d = new OrdenPagoDatosHoja();
            try { d.NombreHoja = Convert.ToString(ws.Name); } catch { d.NombreHoja = "?"; }

            d.Orden = Texto(ws.Range["B15"]);

            if (EsRotuloFecha(Texto(ws.Range["E15"])))
                d.FechaPago = LeerFechaPago(ws.Range["I15"]);
            else if (EsRotuloFecha(Texto(ws.Range["F15"])))
                d.FechaPago = LeerFechaPago(ws.Range["J15"]);
            else
            {
                LoggerHelper.Warning($"Hoja '{d.NombreHoja}': rótulo de fecha no reconocido en E15/F15. FechaPago quedará vacía.");
                d.FechaPago = "";
            }

            if (RotuloCoincide(Texto(ws.Range["Q16"]), ROTULO_OBLIGACION))
                d.Obligacion = Texto(ws.Range["X16"]);
            else if (RotuloCoincide(Texto(ws.Range["O16"]), ROTULO_OBLIGACION))
                d.Obligacion = Texto(ws.Range["U16"]);
            else
                d.Obligacion = "";

            if (RotuloCoincide(Texto(ws.Range["F30"]), ROTULO_TERCERO))
                d.Tercero = Texto(ws.Range["J30"]);
            else if (RotuloCoincide(Texto(ws.Range["E30"]), ROTULO_TERCERO))
                d.Tercero = Texto(ws.Range["I30"]);
            else
                d.Tercero = "";

            return d;
        }

        /// <summary>
        /// Busca la celda con el rótulo "Objeto:" dentro del rango A1:BH300 y obtiene su valor adyacente
        /// para resolver la categoría fallback (Decisión D-4).
        /// </summary>
        internal static string BuscarObjeto(dynamic ws)
        {
            try
            {
                dynamic rng = ws.Range["A1:BH300"].Find("Objeto:");
                if (rng == null) return "";
                string valor = Texto(rng.Offset[0, 1]);
                if (valor == "") valor = Texto(rng.Offset[0, 2]);
                return valor;
            }
            catch { return ""; }
        }
    }
}
