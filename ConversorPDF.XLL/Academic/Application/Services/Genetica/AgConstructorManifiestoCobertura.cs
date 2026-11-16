using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Construye demanda desde selección explícita y relaciones requeridas, sin índices con First ni genes.</summary>
    public sealed class AgConstructorManifiestoCobertura
    {
        /// <summary>Valida todas las aulas seleccionadas y devuelve un snapshot sin mutar el contexto.
        /// La territorial/subunidad deben proceder de la selección confirmada, nunca de configuración global implícita.</summary>
        public Result<AgManifiestoCobertura> Construir(AgContexto contexto, IEnumerable<int> idsAulasSeleccionadas,
            int idTerritorial, string subunidad, DateTime inicio, DateTime fin)
        {
            if (contexto == null || contexto.Asignaturas == null || idsAulasSeleccionadas == null || idTerritorial <= 0 || string.IsNullOrWhiteSpace(subunidad)
                || fin.Date < inicio.Date) return Result.Failure<AgManifiestoCobertura>("MANIFIESTO_ENTRADA_INVALIDA");
            var ids = idsAulasSeleccionadas.ToList();
            if (ids.Count == 0 || ids.Any(i => i <= 0) || ids.Distinct().Count() != ids.Count)
                return Result.Failure<AgManifiestoCobertura>("SELECCION_VACIA_O_DUPLICADA");
            var items = new List<AgRelacionCobertura>();
            foreach (int id in ids.OrderBy(i => i))
            {
                var aulas = (contexto.Aulas ?? new List<AgAula>()).Where(a => a.Id == id).ToList();
                if (aulas.Count != 1) return Fallo("AULA_AUSENTE_O_AMBIGUA", id);
                var aula = aulas[0];
                if (aula.Estado != AgEstadoAula.Disponible && aula.Estado != AgEstadoAula.Aprobado)
                    return Fallo("AULA_NO_PROGRAMABLE", id);
                if (AgClaveRelacionCobertura.Texto(aula.Subunidad) != AgClaveRelacionCobertura.Texto(subunidad))
                    return Fallo("SUBUNIDAD_NO_COINCIDE", id);
                var relaciones = (contexto.AulaAsignaturas ?? new List<AgAulaAsignatura>())
                    .Where(r => r.IdAula == id && r.Requerida == 1).ToList();
                if (relaciones.Count == 0 || relaciones.GroupBy(r => r.IdAsignatura).Any(g => g.Count() != 1))
                    return Fallo("RELACIONES_AUSENTES_O_AMBIGUAS", id);
                foreach (var relacion in relaciones.OrderBy(r => r.IdAsignatura))
                {
                    // Un vacío histórico no se trata como permiso para cruzar particiones.
                    if (AgClaveRelacionCobertura.Texto(relacion.Subunidad) != AgClaveRelacionCobertura.Texto(subunidad))
                        return Fallo("SUBUNIDAD_RELACION_NO_COINCIDE", id);
                    var asignaturas = contexto.Asignaturas.Where(a => a.Id == relacion.IdAsignatura).ToList();
                    if (asignaturas.Count != 1) return Fallo("ASIGNATURA_AUSENTE_O_AMBIGUA", id);
                    var asignatura = asignaturas[0];
                    string semestre = AgFormatoAsignatura.SemestreRomano(aula.Semestre);
                    if (semestre == "SIN ASIGNAR" || asignatura.IdPrograma != aula.IdPrograma
                        || (!string.IsNullOrWhiteSpace(asignatura.Semestre)
                            && AgFormatoAsignatura.SemestreRomano(asignatura.Semestre) != semestre)
                        || (!string.IsNullOrWhiteSpace(asignatura.Subunidad)
                            && AgClaveRelacionCobertura.Texto(asignatura.Subunidad) != AgClaveRelacionCobertura.Texto(subunidad)))
                        return Fallo("IDENTIDAD_ASIGNATURA_NO_COINCIDE", id);
                    // Demanda original antes del redondeo; la demanda efectiva reutiliza la autoridad vigente.
                    double original = relacion.HorasObjetivo > 0 ? relacion.HorasObjetivo
                        : asignatura.IntHoraria > 0 ? asignatura.IntHoraria : asignatura.Horas > 0 ? asignatura.Horas
                        : asignatura.Creditos > 0 ? asignatura.Creditos * 16.0 : 64.0;
                    try
                    {
                        var clave = new AgClaveRelacionCobertura(idTerritorial, subunidad, aula.IdPeriodo, aula.IdSede,
                            aula.IdPrograma, aula.Id, asignatura.Id, semestre, aula.Grupo, aula.ModalidadVinculacion);
                        items.Add(new AgRelacionCobertura(clave, original,
                            AgEvaluadorDemandaVentana.HorasObjetivo(asignatura, relacion), (int)aula.Estado));
                    }
                    catch (ArgumentException ex) { return Fallo(ex.Message, id); }
                }
            }
            return Result.Success(new AgManifiestoCobertura(inicio, fin, items));
        }

        private static Result<AgManifiestoCobertura> Fallo(string codigo, int id)
            => Result.Failure<AgManifiestoCobertura>(codigo + ":" + id);
    }
}
