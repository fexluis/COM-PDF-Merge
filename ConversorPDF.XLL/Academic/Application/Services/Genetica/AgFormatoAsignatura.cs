using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Proveedor de formateo estándar y legible para asignaturas y semestres en mensajes de diagnóstico,
    /// pre-vuelo y bitácoras del motor genético (Fase 325).
    /// </summary>
    public static class AgFormatoAsignatura
    {
        /// <summary>
        /// Símbolo separador configurable. Por defecto utiliza flecha Unicode U+2192 (→) con soporte UTF-8.
        /// </summary>
        public static string Flecha { get; set; } = "→";

        /// <summary>
        /// Normaliza un identificador o texto de semestre a su notación en números romanos estándar.
        /// Convierte formatos numéricos arábigos (ej. "2" → "II"), variantes en minúsculas (ej. "ii" → "II"),
        /// y retorna "SIN ASIGNAR" ante valores nulos o vacíos.
        /// </summary>
        /// <param name="semestre">Texto o número del semestre a normalizar.</param>
        /// <returns>Semestre en números romanos o texto de contingencia.</returns>
        public static string SemestreRomano(string semestre)
        {
            if (string.IsNullOrWhiteSpace(semestre))
            {
                return "SIN ASIGNAR";
            }

            string s = semestre.Trim();
            switch (s.ToUpperInvariant())
            {
                case "1":
                case "I":
                    return "I";
                case "2":
                case "II":
                    return "II";
                case "3":
                case "III":
                    return "III";
                case "4":
                case "IV":
                    return "IV";
                case "5":
                case "V":
                    return "V";
                case "6":
                case "VI":
                    return "VI";
                case "7":
                case "VII":
                    return "VII";
                case "8":
                case "VIII":
                    return "VIII";
                case "9":
                case "IX":
                    return "IX";
                case "10":
                case "X":
                    return "X";
                case "11":
                case "XI":
                    return "XI";
                case "12":
                case "XII":
                    return "XII";
                default:
                    return s.ToUpperInvariant();
            }
        }

        /// <summary>
        /// Mapeo canónico de equivalencias entre números arábigos y romanos (1..12).
        /// </summary>
        private static readonly Tuple<string, string>[] MapaSemestres = new Tuple<string, string>[]
        {
            Tuple.Create("1", "I"),
            Tuple.Create("2", "II"),
            Tuple.Create("3", "III"),
            Tuple.Create("4", "IV"),
            Tuple.Create("5", "V"),
            Tuple.Create("6", "VI"),
            Tuple.Create("7", "VII"),
            Tuple.Create("8", "VIII"),
            Tuple.Create("9", "IX"),
            Tuple.Create("10", "X"),
            Tuple.Create("11", "XI"),
            Tuple.Create("12", "XII")
        };

        /// <summary>
        /// Normaliza un identificador o texto de semestre a su notación arábiga estándar (ej. "IV" → "4", "iv" → "4", "4" → "4").
        /// Si el valor es nulo, vacío o "SIN ASIGNAR", retorna una cadena vacía (Fase 378).
        /// </summary>
        /// <param name="semestre">Texto o número del semestre a normalizar.</param>
        /// <returns>Semestre en números arábigos ("1".."12") o cadena vacía si no tiene correspondencia.</returns>
        public static string SemestreArabigo(string semestre)
        {
            if (string.IsNullOrWhiteSpace(semestre))
            {
                return string.Empty;
            }

            string r = SemestreRomano(semestre);
            if (r == "SIN ASIGNAR")
            {
                return string.Empty;
            }

            for (int i = 0; i < MapaSemestres.Length; i++)
            {
                if (string.Equals(MapaSemestres[i].Item2, r, StringComparison.OrdinalIgnoreCase))
                {
                    return MapaSemestres[i].Item1;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Separa y normaliza una cadena combinada de semestre y grupo (ej. "II-A" → ("II", "A"), "VIII" → ("VIII", ""), "8-b" → ("VIII", "B")).
        /// Compatible con C# 7.3 utilizando array explícito en Split (Fase 378 / Fase 393).
        /// </summary>
        /// <param name="valor">Cadena combinada de semestre y opcionalmente grupo separada por guión.</param>
        /// <returns>Tupla conteniendo el semestre normalizado en romano y el grupo en mayúsculas.</returns>
        public static (string Semestre, string Grupo) ParsearSemestreGrupo(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return (string.Empty, string.Empty);
            }

            string[] partes = valor.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            string sem = SemestreRomano(partes[0]);
            string grp = partes.Length >= 2 ? partes[1].Trim().ToUpperInvariant() : string.Empty;
            return (sem, grp);
        }

        /// <summary>
        /// Formatea una línea descriptiva unificada de asignatura: "Id: Semestre → Nombre".
        /// </summary>
        /// <param name="a">Asignatura a formatear.</param>
        /// <returns>Cadena formateada o vacía si el objeto es nulo.</returns>
        public static string Linea(Asignatura a)
        {
            if (a == null) return string.Empty;
            return $"{a.Id}: {SemestreRomano(a.Semestre)} {Flecha} {a.Nombre}";
        }

        /// <summary>
        /// Genera una lista formateada con viñetas para un conjunto de asignaturas.
        /// Cada elemento está separado por "\n• ".
        /// </summary>
        /// <param name="asignaturas">Colección de asignaturas.</param>
        /// <returns>Texto compuesto listo para presentación en cuadros de diálogo o bitácoras.</returns>
        public static string BulletList(IEnumerable<Asignatura> asignaturas)
        {
            if (asignaturas == null) return string.Empty;
            return string.Join("\n• ", asignaturas.Where(a => a != null).Select(Linea));
        }
    }
}
