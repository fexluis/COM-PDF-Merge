using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Construcción, interpretación y normalización de los nombres de archivo de informes
    /// bajo el estándar <c>{solicitud}-{cedula}-{tipo}-{nombre}</c>.
    /// </summary>
    public static class InformeNombreHelper
    {
        /// <summary> Longitud máxima del segmento de nombre del comisionado. </summary>
        public const int LongitudNombreComisionado = 12;

        /// <summary> Nombre usado cuando CONSULTAS no aporta el nombre del comisionado. </summary>
        public const string NombreDesconocido = "DESCONOCIDO";

        private static readonly char[] CaracteresInvalidos =
            { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '.', ' ', ',', '-', '_', '\t', '\r', '\n' };

        /// <summary>
        /// Quita tildes y diacríticos de un texto (á→a, ñ→n, ü→u).
        /// </summary>
        /// <param name="texto">Texto original.</param>
        /// <returns>Texto sin diacríticos.</returns>
        public static string QuitarDiacriticos(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return string.Empty;

            var descompuesto = texto.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(descompuesto.Length);

            foreach (char c in descompuesto)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        /// <summary>
        /// Deja solo dígitos de un texto.
        /// </summary>
        /// <param name="valor">Texto de entrada.</param>
        /// <returns>Cadena con dígitos únicamente.</returns>
        public static string SoloDigitos(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return string.Empty;

            var sb = new StringBuilder(valor.Length);
            foreach (char c in valor)
            {
                if (c >= '0' && c <= '9') sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Normaliza un número de documento quitando espacios, puntos y comas
        /// (12.345.678 → 12345678).
        /// </summary>
        /// <param name="valor">Número a normalizar.</param>
        /// <returns>Número en solo dígitos.</returns>
        public static string NormalizarDocumento(string valor)
        {
            return Regex.Replace(valor ?? string.Empty, @"[\s\.,\u00A0]", string.Empty);
        }

        /// <summary>
        /// Limpia el nombre del comisionado para usarlo en el nombre del archivo:
        /// sin tildes, sin separadores y en mayúsculas.
        /// </summary>
        /// <param name="nombre">Nombre original.</param>
        /// <param name="max">Longitud máxima (por defecto 12).</param>
        /// <returns>Nombre limpio, o vacío si no queda ningún carácter válido.</returns>
        public static string LimpiarNombre(string nombre, int max = LongitudNombreComisionado)
        {
            if (string.IsNullOrWhiteSpace(nombre)) return string.Empty;

            string sinAcentos = QuitarDiacriticos(nombre);
            var sb = new StringBuilder(sinAcentos.Length);

            foreach (char c in sinAcentos)
            {
                if (Array.IndexOf(CaracteresInvalidos, c) >= 0) continue;
                if (!char.IsLetterOrDigit(c)) continue;
                sb.Append(c);
            }

            string limpio = sb.ToString().ToUpperInvariant();
            if (max > 0 && limpio.Length > max) limpio = limpio.Substring(0, max);
            return limpio;
        }

        /// <summary>
        /// Construye el nombre estandarizado (sin extensión).
        /// </summary>
        /// <param name="solicitud">Número de solicitud.</param>
        /// <param name="cedula">Cédula del comisionado.</param>
        /// <param name="tipo">Tipo (INF, CUM, UNI, ERR).</param>
        /// <param name="nombreComisionado">Nombre del comisionado.</param>
        /// <returns>Nombre del archivo sin extensión.</returns>
        public static string Construir(string solicitud, string cedula, string tipo, string nombreComisionado)
        {
            string num = SoloDigitos(solicitud);
            string cc = SoloDigitos(cedula);
            string t = (tipo ?? string.Empty).Trim().ToUpperInvariant();
            string nombre = LimpiarNombre(nombreComisionado);

            if (string.IsNullOrEmpty(num)) num = "0";
            if (string.IsNullOrEmpty(cc)) cc = "000000";
            if (!InformeTipo.EsTipoValido(t)) t = InformeTipo.Err;
            if (string.IsNullOrEmpty(nombre)) nombre = NombreDesconocido;

            return num + "-" + cc + "-" + t + "-" + nombre;
        }

        /// <summary>
        /// Interpreta un nombre de archivo con el estándar completo (tipo válido incluido).
        /// </summary>
        /// <param name="nombreArchivo">Nombre del archivo (con o sin extensión).</param>
        /// <param name="info">Datos interpretados.</param>
        /// <returns>True si el nombre cumple el estándar.</returns>
        public static bool TryParse(string nombreArchivo, out InformePdf info)
        {
            info = null;

            string solicitud, cedula, tipo, nombre;
            if (!TryParseLoose(nombreArchivo, out solicitud, out cedula, out tipo, out nombre)) return false;
            if (!InformeTipo.EsTipoValido(tipo)) return false;

            info = new InformePdf
            {
                NumeroSolicitud = solicitud,
                Cedula = cedula,
                Tipo = tipo,
                NombreComisionado = nombre,
                NombreArchivo = Path.GetFileName(nombreArchivo)
            };
            return true;
        }

        /// <summary>
        /// Interpreta un nombre con la estructura <c>{solicitud}-{cedula}-{tipo}-{nombre}</c>
        /// sin exigir que el tipo sea uno de los admitidos (útil para nombres heredados con XXX).
        /// </summary>
        /// <param name="nombreArchivo">Nombre del archivo.</param>
        /// <param name="solicitud">Número de solicitud.</param>
        /// <param name="cedula">Cédula.</param>
        /// <param name="tipo">Tipo tal como aparece (mayúsculas).</param>
        /// <param name="nombre">Resto del nombre.</param>
        /// <returns>True si la estructura es válida.</returns>
        public static bool TryParseLoose(string nombreArchivo, out string solicitud, out string cedula, out string tipo, out string nombre)
        {
            solicitud = string.Empty;
            cedula = string.Empty;
            tipo = string.Empty;
            nombre = string.Empty;

            if (string.IsNullOrWhiteSpace(nombreArchivo)) return false;

            string baseNombre = Path.GetFileNameWithoutExtension(nombreArchivo).Replace('_', '-');
            string[] partes = baseNombre.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (partes.Length < 4) return false;
            if (!EsNumero(partes[0]) || !EsNumero(partes[1])) return false;

            solicitud = partes[0];
            cedula = partes[1];
            tipo = partes[2].ToUpperInvariant();
            nombre = string.Join("-", partes.Skip(3));
            return true;
        }

        /// <summary>
        /// Extrae el número de solicitud del inicio del nombre; si no lo encuentra al inicio,
        /// devuelve el primer bloque numérico del nombre.
        /// </summary>
        /// <param name="nombreArchivo">Nombre del archivo.</param>
        /// <returns>Número de solicitud o cadena vacía.</returns>
        public static string ExtraerNumeroSolicitud(string nombreArchivo)
        {
            string baseNombre = Path.GetFileNameWithoutExtension(nombreArchivo ?? string.Empty);
            if (string.IsNullOrWhiteSpace(baseNombre)) return string.Empty;

            Match alInicio = Regex.Match(baseNombre, @"^(\d+)");
            if (alInicio.Success) return alInicio.Groups[1].Value;

            Match cualquiera = Regex.Match(baseNombre, @"(\d+)");
            return cualquiera.Success ? cualquiera.Groups[1].Value : string.Empty;
        }

        /// <summary>
        /// Genera una ruta libre de colisión dentro de la carpeta añadiendo sufijos _1, _2…
        /// </summary>
        /// <param name="carpeta">Carpeta destino.</param>
        /// <param name="nombreArchivo">Nombre con extensión.</param>
        /// <returns>Ruta completa que no existe en disco.</returns>
        public static string NombreUnico(string carpeta, string nombreArchivo)
        {
            string carpetaFinal = string.IsNullOrWhiteSpace(carpeta) ? "." : carpeta;
            string baseNombre = Path.GetFileNameWithoutExtension(nombreArchivo);
            string extension = Path.GetExtension(nombreArchivo);

            string candidato = Path.Combine(carpetaFinal, baseNombre + extension);
            int contador = 1;

            while (File.Exists(candidato) && contador <= 999)
            {
                candidato = Path.Combine(carpetaFinal, baseNombre + "_" + contador + extension);
                contador++;
            }

            return candidato;
        }

        private static bool EsNumero(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            for (int i = 0; i < valor.Length; i++)
            {
                if (valor[i] < '0' || valor[i] > '9') return false;
            }
            return true;
        }
    }
}
