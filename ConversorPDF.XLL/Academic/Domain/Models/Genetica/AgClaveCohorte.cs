using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Generador canónico y unificado de claves de cohorte (Fase 953).
    /// Asegura paridad entre AgFabricaIndividuos y AgValidadorReglas.
    /// </summary>
    public static class AgClaveCohorte
    {
        // Fase 1033: De(...) -> DesdeCohorte / DesdeAula / DesdeGen para evitar
        // "Repeated function name" (Excel-DNA registra por nombre de método).
        // Fase 1036: key-builder de dominio; internal evita registrarlo como UDF
        // (Excel-DNA 1.8.0 ignora el valor por defecto y exigiría 6 argumentos).
        internal static string DesdeCohorte(int idSede, int idPrograma, int semestre, string grupo, string aulaRef = null)
        {
            if (!string.IsNullOrWhiteSpace(aulaRef))
            {
                return aulaRef.Trim();
            }

            return $"{idSede}|{idPrograma}|{semestre}|{(grupo ?? string.Empty).Trim()}";
        }

        public static string DesdeAula(AgAula aula)
        {
            if (aula == null) return "A";
            if (!string.IsNullOrWhiteSpace(aula.AulaRef))
            {
                return aula.AulaRef.Trim();
            }

            return $"{aula.IdSede}|{aula.IdPrograma}|{aula.Semestre}|{(aula.Grupo ?? string.Empty).Trim()}";
        }

        public static string DesdeGen(AgGen gen)
        {
            if (gen == null) return "G";
            return $"{gen.IdSede}|{gen.IdPrograma}|{gen.Semestre}|{(gen.Grupo ?? string.Empty).Trim()}";
        }
    }
}
