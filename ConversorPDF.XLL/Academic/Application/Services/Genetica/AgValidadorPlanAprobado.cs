using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Valida origen y cobertura temporal sin mutar contexto, fechas, aprobación ni producción.</summary>
    public sealed class AgValidadorPlanAprobado
    {
        /// <summary>Valida todas las relaciones del snapshot contra su propia ventana y los festivos capturados.
        /// Las horas lectivas no pueden exceder el intervalo ocupado; no se inventa conversión académica.
        /// La reconciliación histórica/manual siempre espera ratificación T4.</summary>
        public AgValidacionPlanAprobado Validar(AgPlanAprobadoSnapshot snapshot, AgContexto contexto)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));
            var m = snapshot.Manifiesto;
            var globales = snapshot.Diagnosticos.ToList();
            var requeridas = m.Relaciones.ToDictionary(r => r.Clave.Valor, StringComparer.Ordinal);
            var reconciliadas = snapshot.Relaciones.Where(r => r.ClavesCandidatas.Count == 1
                && requeridas.ContainsKey(r.ClavesCandidatas[0].Valor)).ToList();
            foreach (var r in snapshot.Relaciones.Except(reconciliadas))
                globales.Add("RELACION_NO_RECONCILIADA:" + r.IdVinculacion);
            foreach (var g in snapshot.Relaciones.GroupBy(r => r.IdVinculacion).Where(g => g.Count() > 1))
                globales.Add("VINCULACION_DUPLICADA:" + g.Key);
            var sesiones = snapshot.Relaciones.SelectMany(r => r.Sesiones).ToList();
            var repetidas = new HashSet<long>(sesiones.GroupBy(s => s.IdSesion).Where(g => g.Count() > 1).Select(g => g.Key));
            var errores = new Dictionary<AgSesionCanonica, List<string>>();
            foreach (var r in snapshot.Relaciones)
            foreach (var s in r.Sesiones)
            {
                if (!errores.ContainsKey(s)) errores[s] = new List<string>();
                var e = errores[s];
                if (repetidas.Contains(s.IdSesion)) e.Add("SESION_DUPLICADA:" + s.IdSesion);
                if (s.IdSesion <= 0 || s.IdVinculacion != r.IdVinculacion || r.IdVinculacion <= 0
                    || (r.ClavesCandidatas.Count == 1 && !string.Equals(s.Subunidad?.Trim(),
                        r.ClavesCandidatas[0].Subunidad, StringComparison.OrdinalIgnoreCase)))
                    e.Add("IDENTIDAD_SESION_INVALIDA:" + s.IdSesion);
                if (!IntervaloValido(s) || !HorasValidas(s.HorasLectivas) || s.IdTercero <= 0
                    || s.IdTercero == contexto.IdTerceroVacante)
                    e.Add("SESION_NO_VERIFICABLE:" + s.IdSesion);
                if (IntervaloValido(s) && HorasValidas(s.HorasLectivas)
                    && s.HorasLectivas > (s.HoraFin.Value - s.HoraInicio.Value).TotalHours + 0.001)
                    e.Add("DURACION_INVALIDA:" + s.IdSesion);
                if (s.Fecha.HasValue)
                {
                    if (s.Fecha.Value.Date < m.Inicio || s.Fecha.Value.Date > m.Fin)
                        e.Add("SESION_FUERA_DE_VENTANA:" + s.IdSesion);
                    if (contexto.Festivos.Any(f => s.Fecha.Value.Date >= f.Inicio.Date && s.Fecha.Value.Date <= f.Fin.Date))
                        e.Add("SESION_EN_FESTIVO:" + s.IdSesion);
                }
            }
            // Intervalos reales, extremos contiguos permitidos; no compara una sesión consigo misma.
            // Incluye otras aulas reconciliadas (también Disponible) como ocupación docente.
            var conocidas = reconciliadas.SelectMany(r => r.Sesiones.Select(s => new
                { Clave = r.ClavesCandidatas[0], Sesion = s })).Where(x => IntervaloValido(x.Sesion)).ToList();
            foreach (var dia in conocidas.GroupBy(x => x.Sesion.Fecha.Value.Date))
            {
                var lista = dia.OrderBy(x => x.Sesion.HoraInicio.Value).ToList();
                for (int i = 0; i < lista.Count; i++)
                for (int j = i + 1; j < lista.Count && lista[j].Sesion.HoraInicio.Value < lista[i].Sesion.HoraFin.Value; j++)
                {
                    var a = lista[i]; var b = lista[j];
                    if (a.Sesion.IdSesion == b.Sesion.IdSesion) continue;
                    string tipo = a.Clave.IdAula == b.Clave.IdAula ? "CRUCE_AULA:"
                        : a.Sesion.IdTercero > 0 && a.Sesion.IdTercero == b.Sesion.IdTercero ? "CRUCE_DOCENTE:" : null;
                    if (tipo == null) continue;
                    errores[a.Sesion].Add(tipo + a.Sesion.IdSesion + ":" + b.Sesion.IdSesion);
                    errores[b.Sesion].Add(tipo + b.Sesion.IdSesion + ":" + a.Sesion.IdSesion);
                }
            }
            var aulas = new List<AgValidacionAulaPlan>();
            foreach (var aula in m.Relaciones.GroupBy(r => r.Clave.IdAula))
            {
                bool aprobada = aula.Any(r => r.EstadoAula == (int)AgEstadoAula.Aprobado);
                var resultados = new List<AgValidacionRelacionPlan>();
                foreach (var requerida in aula)
                {
                    var relaciones = reconciliadas.Where(r => r.ClavesCandidatas[0].Valor == requerida.Clave.Valor).ToList();
                    var e = new List<string>();
                    if (relaciones.Count == 0) e.Add("RELACION_AUSENTE:" + requerida.Clave.Valor);
                    if (relaciones.Count > 1) e.Add("VINCULACIONES_MULTIPLES:" + requerida.Clave.Valor);
                    var leidas = relaciones.SelectMany(r => r.Sesiones).ToList();
                    if (leidas.Count == 0) e.Add("SIN_SESIONES:" + requerida.Clave.Valor);
                    e.AddRange(leidas.SelectMany(s => errores[s]));
                    double validas = relaciones.Count == 1
                        ? leidas.Where(s => errores[s].Count == 0).Sum(s => s.HorasLectivas) : 0;
                    double fuera = leidas.Where(s => s.Fecha.HasValue && (s.Fecha.Value.Date < m.Inicio || s.Fecha.Value.Date > m.Fin)
                        && HorasValidas(s.HorasLectivas)).GroupBy(s => s.IdSesion).Sum(g => g.First().HorasLectivas);
                    resultados.Add(new AgValidacionRelacionPlan(requerida, validas, fuera, e));
                }
                var estado = !aprobada ? AgEstadoValidacionPlan.NoRequerida
                    : globales.Count > 0 || resultados.Any(r => r.Diagnosticos.Count > 0 || r.Exceso > 0.001)
                        ? AgEstadoValidacionPlan.PlanNoVerificable : AgEstadoValidacionPlan.RequiereRatificacion;
                aulas.Add(new AgValidacionAulaPlan(aula.Key, estado, resultados));
            }
            return new AgValidacionPlanAprobado(snapshot.HuellaOrigen, aulas, globales);
        }

        private static bool HorasValidas(double horas) => horas > 0 && !double.IsNaN(horas) && !double.IsInfinity(horas);
        private static bool IntervaloValido(AgSesionCanonica s) => s.Fecha.HasValue && s.HoraInicio.HasValue && s.HoraFin.HasValue
            && s.HoraInicio.Value >= TimeSpan.Zero && s.HoraFin.Value <= TimeSpan.FromDays(1)
            && s.HoraFin.Value > s.HoraInicio.Value;
    }
}
