using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Catálogo canónico de modalidades de tutoría académica.
    /// Código = valor persistido en base de datos; Etiqueta = texto visible en interfaz de usuario.
    /// Acepta valores legados (Virtual, Semipresencial, A Distancia) para retrocompatibilidad con bases de datos sin migrar.
    /// </summary>
    public static class ModalidadTutoriaCatalog
    {
        public const string Presencial = "Presencial";
        public const string Plataforma = "Plataforma";
        public const string Mixta      = "Mixta";

        private static readonly string[] _opciones = new string[] { Presencial, Plataforma, Mixta };

        /// <summary>
        /// Opciones canónicas en el orden oficial de presentación.
        /// </summary>
        public static IReadOnlyList<string> Opciones { get; } = Array.AsReadOnly(_opciones);

        /// <summary>
        /// Pares tipados (Código persistido, Etiqueta visible) para enlace a controles UI.
        /// </summary>
        public static IReadOnlyList<(string Codigo, string Etiqueta)> Pares { get; } = new List<(string Codigo, string Etiqueta)>
        {
            (Presencial, Presencial),
            (Plataforma, Plataforma),
            (Mixta, Mixta)
        }.AsReadOnly();

        /// <summary>
        /// Remueve acentos y diacríticos normalizando a caracteres ASCII en mayúsculas.
        /// </summary>
        private static string LimpiarTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return string.Empty;
            string formD = texto.Trim().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            for (int i = 0; i < formD.Length; i++)
            {
                char ch = formD[i];
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(ch);
                }
            }
            return sb.ToString().ToUpperInvariant();
        }

        /// <summary>
        /// Convierte cualquier variante textual (incluidas modalidades legadas) al código canónico.
        /// Null o cadena en blanco retorna 'Presencial' por defecto.
        /// </summary>
        /// <remarks>Renombrado desde 'Normalizar' para evitar colision con nombres reservados/UDF de Excel (mismo criterio que AgFranjaJornada.NormalizarJornada).</remarks>
        public static string NormalizarTutoria(string modalidad)
        {
            if (string.IsNullOrWhiteSpace(modalidad)) return Presencial;

            string limpio = LimpiarTexto(modalidad);

            // 1. Plataforma / Virtual / En línea / E-learning
            if (limpio.Contains("PLATAFORMA") || limpio.Contains("VIRTUAL") || limpio.Contains("EN LINEA") || limpio.Contains("E-LEARNING"))
            {
                return Plataforma;
            }

            // 2. Mixta / Semipresencial / A Distancia (evaluado ANTES de Presencial porque 'Semipresencial' contiene 'Presencial')
            if (limpio.Contains("MIXTA") || limpio.Contains("SEMI") || limpio.Contains("DISTANCIA"))
            {
                return Mixta;
            }

            // 3. Presencial
            if (limpio.Contains("PRESENCIAL"))
            {
                return Presencial;
            }

            // 4. Desconocido: retornar el valor original recortado
            return modalidad.Trim();
        }

        /// <summary>
        /// Retorna la etiqueta visible correspondiente al código o valor de modalidad.
        /// Retorna cadena vacía ante null/blanco.
        /// </summary>
        public static string AEtiqueta(string codigo)
        {
            if (string.IsNullOrWhiteSpace(codigo)) return string.Empty;

            string normalizado = NormalizarTutoria(codigo);
            if (Opciones.Contains(normalizado, StringComparer.Ordinal))
            {
                return normalizado;
            }

            return codigo.Trim();
        }

        /// <summary>
        /// Indica si la modalidad especificada corresponde a Presencial.
        /// </summary>
        public static bool EsPresencial(string modalidad)
        {
            return string.Equals(NormalizarTutoria(modalidad), Presencial, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Indica si la modalidad especificada corresponde a Plataforma (o legada Virtual).
        /// </summary>
        public static bool EsPlataforma(string modalidad)
        {
            return string.Equals(NormalizarTutoria(modalidad), Plataforma, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Indica si la modalidad especificada corresponde a Mixta (o legada Semipresencial / A Distancia).
        /// </summary>
        public static bool EsMixta(string modalidad)
        {
            return string.Equals(NormalizarTutoria(modalidad), Mixta, StringComparison.OrdinalIgnoreCase);
        }
    }
}
