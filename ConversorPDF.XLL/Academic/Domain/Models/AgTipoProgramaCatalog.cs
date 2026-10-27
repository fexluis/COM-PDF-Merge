using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Catálogo y filtro canónico de tipo/nivel académico de programa (Fase 1311 / 1406).
    /// </summary>
    public static class AgTipoProgramaCatalog
    {
        public const string Todos = "TODOS";
        public const string Pregrado = "PREGRADO";
        public const string Posgrado = "POSGRADO";

        /// <summary>
        /// Determina el nivel académico a partir de la modalidad del programa (PREGRADO / POSGRADO).
        /// </summary>
        public static AgNivelAcademico Nivel(string modalidadPrograma)
        {
            return AgNivelAcademicoCatalogo.Desde(modalidadPrograma);
        }

        /// <summary>
        /// Alias para obtener el nivel académico.
        /// </summary>
        public static AgNivelAcademico Desde(string modalidadPrograma)
        {
            return Nivel(modalidadPrograma);
        }

        /// <summary>
        /// Retorna la etiqueta textual en mayúsculas del nivel académico.
        /// </summary>
        public static string Etiqueta(AgNivelAcademico nivel)
        {
            return nivel == AgNivelAcademico.Posgrado ? Posgrado : Pregrado;
        }

        /// <summary>
        /// Determina las horas por día sugeridas para la modalidad según nivel académico (Pregrado 4h / Posgrado 8h).
        /// </summary>
        public static int HorasDiaSugeridas(AgNivelAcademico nivel)
        {
            return nivel == AgNivelAcademico.Posgrado ? 8 : 4;
        }

        public static int HorasDiaSugeridas(string modalidadPrograma)
        {
            return HorasDiaSugeridas(Nivel(modalidadPrograma));
        }

        public static bool Coincide(string filtro, string programaModalidad)
        {
            if (string.IsNullOrWhiteSpace(filtro) || filtro.StartsWith(Todos, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var nivelObjetivo = filtro.StartsWith(Posgrado, StringComparison.OrdinalIgnoreCase)
                ? AgNivelAcademico.Posgrado
                : AgNivelAcademico.Pregrado;

            return Nivel(programaModalidad) == nivelObjetivo;
        }

        /// <summary>
        /// Verifica si un programa coincide con el tipo de programa seleccionado.
        /// </summary>
        public static bool Coincide(Programa p, string tipo)
        {
            if (p == null) return false;
            return Coincide(tipo, p.Modalidad);
        }

        /// <summary>
        /// Filtra una lista de programas según el tipo seleccionado.
        /// </summary>
        public static List<Programa> Filtrar(IEnumerable<Programa> programas, string tipo)
        {
            if (programas == null) return new List<Programa>();
            if (string.IsNullOrWhiteSpace(tipo) || tipo.StartsWith(Todos, StringComparison.OrdinalIgnoreCase))
            {
                return programas.ToList();
            }

            return programas.Where(p => Coincide(p, tipo)).ToList();
        }

        /// <summary>
        /// Opciones formateadas para poblar el ComboBox de tipos de programa.
        /// </summary>
        public static IReadOnlyList<string> Opciones(int nPregrado, int nPosgrado)
        {
            return new[]
            {
                Todos,
                $"{Pregrado} ({nPregrado})",
                $"{Posgrado} ({nPosgrado})"
            };
        }

        /// <summary>
        /// Retorna la etiqueta compacta del programa (Abreviatura si existe, de lo contrario NombrePrograma).
        /// Fase 1448.
        /// </summary>
        public static string EtiquetaPrograma(Programa p)
        {
            if (p == null) return string.Empty;
            if (!string.IsNullOrWhiteSpace(p.Abreviatura)) return p.Abreviatura.Trim();
            return (p.NombrePrograma ?? string.Empty).Trim();
        }
    }
}
