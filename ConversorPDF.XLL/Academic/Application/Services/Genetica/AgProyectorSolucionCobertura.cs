using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Proyecta la unión de sesiones canónicas conservadas y staging nuevo sin escribir ninguna de ellas.</summary>
    public sealed class AgProyectorSolucionCobertura
    {
        /// <summary>Exige una decisión publicada por cada aula aprobada; no interpreta el estado AG como aprobación productiva.</summary>
        public Result<IReadOnlyCollection<AgSesionCobertura>> Conservar(AgContexto contexto)
        {
            var m = contexto?.ManifiestoCobertura;
            if (m == null) return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("MANIFIESTO_AUSENTE");
            var aprobadas = m.Relaciones.Where(r => r.EstadoAula == (int)AgEstadoAula.Aprobado)
                .Select(r => r.Clave.IdAula).Distinct().ToList();
            if (aprobadas.Count == 0)
                return Result.Success<IReadOnlyCollection<AgSesionCobertura>>(Array.Empty<AgSesionCobertura>());
            var snapshot = contexto.PlanAprobadoSnapshot;
            var validacion = contexto.ValidacionPlanAprobado;
            if (snapshot == null || validacion == null || snapshot.Manifiesto.Huella != m.Huella ||
                validacion.HuellaOrigen != snapshot.HuellaOrigen || validacion.Diagnosticos.Count != 0 ||
                aprobadas.Any(id => validacion.Aulas.Single(a => a.IdAula == id).Estado != AgEstadoValidacionPlan.RequiereRatificacion))
                return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("PLAN_NO_VERIFICABLE");
            var candidatos = contexto.CandidatosCobertura ?? Array.Empty<AgCandidatoPlanAprobado>();
            if (aprobadas.Any(id => candidatos.Count(c => c.IdAula == id && c.HuellaManifiesto == m.Huella &&
                c.HuellaOrigen == snapshot.HuellaOrigen && c.Inicio == m.Inicio && c.Fin == m.Fin) != 1))
                return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RATIFICACION_NO_ACREDITADA");
            var conservar = new HashSet<int>(aprobadas.Where(id => candidatos.Single(c => c.IdAula == id).Desde ==
                candidatos.Single(c => c.IdAula == id).Hacia));
            var salida = new List<AgSesionCobertura>();
            foreach (var relacion in snapshot.Relaciones.Where(r => r.ClavesCandidatas.Count == 1 &&
                conservar.Contains(r.ClavesCandidatas[0].IdAula)))
            {
                var clave = relacion.ClavesCandidatas[0];
                if (!m.Relaciones.Any(r => r.Clave.Valor == clave.Valor) ||
                    snapshot.Relaciones.Count(r => r.ClavesCandidatas.Count == 1 && r.ClavesCandidatas[0].Valor == clave.Valor) != 1)
                    return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RELACION_NO_RECONCILIADA");
                foreach (var s in relacion.Sesiones)
                {
                    if (!s.Fecha.HasValue || !s.HoraInicio.HasValue || !s.HoraFin.HasValue || s.IdTercero > int.MaxValue)
                        return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("SESION_NO_VERIFICABLE");
                    salida.Add(new AgSesionCobertura(clave, "CONSERVADA", s.IdSesion, (int)s.IdTercero,
                        s.Fecha.Value, s.HoraInicio.Value, s.HoraFin.Value, s.HorasLectivas));
                }
            }
            return Extras(snapshot, m, aprobadas, salida);
        }

        /// <summary>Proyecta las extras declaradas como conservadas inmutables: ocupan aula y docente y nunca
        /// se insertan en staging ni producción. Decisión humana 2026-11-15 «B: extras como ocupación fija».</summary>
        private static Result<IReadOnlyCollection<AgSesionCobertura>> Extras(AgPlanAprobadoSnapshot snapshot,
            AgManifiestoCobertura m, IReadOnlyCollection<int> aprobadas, List<AgSesionCobertura> salida)
        {
            foreach (var extra in snapshot.RelacionesExtra)
            {
                if (extra.ClavesCandidatas.Count != 1)
                    return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RELACION_NO_RECONCILIADA");
                var clave = extra.ClavesCandidatas[0];
                if (!m.EsExtra(clave))
                    return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RELACION_NO_RECONCILIADA");
                if (!aprobadas.Contains(clave.IdAula)) continue;
                foreach (var s in extra.Sesiones)
                {
                    if (!s.Fecha.HasValue || !s.HoraInicio.HasValue || !s.HoraFin.HasValue || s.IdTercero > int.MaxValue)
                        return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("SESION_NO_VERIFICABLE");
                    salida.Add(new AgSesionCobertura(clave, "CONSERVADA", s.IdSesion, (int)s.IdTercero,
                        s.Fecha.Value, s.HoraInicio.Value, s.HoraFin.Value, s.HorasLectivas));
                }
            }
            return Result.Success<IReadOnlyCollection<AgSesionCobertura>>(salida.AsReadOnly());
        }

        /// <summary>Resuelve cada sesión nueva por gen y relación completa; jamás infiere una relación por asignatura sola.</summary>
        public Result<IReadOnlyCollection<AgSesionCobertura>> Nuevas(AgContexto contexto, AgIndividuo individuo,
            IReadOnlyCollection<AgSesion> staging)
        {
            if (contexto?.ManifiestoCobertura == null || individuo?.Genes == null || staging == null)
                return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("COBERTURA_ENTRADA_INVALIDA");
            var genes = individuo.Genes.GroupBy(g => g.Id).ToDictionary(g => g.Key, g => g.ToList());
            var salida = new List<AgSesionCobertura>();
            foreach (var s in staging)
            {
                if (!genes.TryGetValue(s.IdGen, out var encontrados) || encontrados.Count != 1 ||
                    s.IdAula != encontrados[0].IdAula || !s.FechaSesion.HasValue)
                    return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RELACION_INESPERADA");
                var g = encontrados[0];
                var claves = contexto.ManifiestoCobertura.Relaciones.Where(r => r.Clave.IdAula == g.IdAula &&
                    r.Clave.IdAsignatura == g.IdAsignatura && r.Clave.IdPrograma == g.IdPrograma &&
                    r.Clave.IdSede == g.IdSede && r.Clave.IdPeriodo == contexto.IdPeriodo &&
                    r.Clave.Subunidad == AgClaveRelacionCobertura.Texto(g.Subunidad) &&
                    r.Clave.Semestre == AgClaveRelacionCobertura.Texto(g.Semestre) &&
                    r.Clave.Grupo == AgClaveRelacionCobertura.Texto(g.Grupo)).ToList();
                if (claves.Count != 1 || !contexto.PuedeGenerarAula(g.IdAula))
                    return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("RELACION_INESPERADA");
                var horario = HorarioSesionParser.TryParse(s.HorarioSesion);
                if (horario.IsFailure) return Result.Failure<IReadOnlyCollection<AgSesionCobertura>>("SESION_NO_VERIFICABLE");
                salida.Add(new AgSesionCobertura(claves[0].Clave, "NUEVA", s.Id, s.IdTercero,
                    s.FechaSesion.Value, horario.Value.Inicio, horario.Value.Fin, s.Horas, s.EsVacante == 1));
            }
            return Result.Success<IReadOnlyCollection<AgSesionCobertura>>(salida.AsReadOnly());
        }
    }
}
