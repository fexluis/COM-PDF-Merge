using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Comprueba la demanda seleccionada contra sesiones nuevas y conservadas, sin alterar staging.</summary>
    public sealed class AgValidadorCoberturaSeleccion
    {
        /// <summary>Valida identidad, cardinalidad, horas y ocupación de la unión disjunta. La autorización
        /// y la verificación del origen deben haberse comprobado antes de proporcionar conservadas.</summary>
        public Result Validar(AgManifiestoCobertura manifiesto,
            IReadOnlyCollection<AgSesionCobertura> conservadas, IReadOnlyCollection<AgSesionCobertura> nuevas,
            int idTerceroVacante = 0, IReadOnlyCollection<(DateTime Inicio, DateTime Fin)> festivos = null,
            IReadOnlyCollection<int> asignaturasVacantesConsentidas = null,
            ICollection<AgDetalleVacanteCobertura> diagnosticos = null)
        {
            if (manifiesto == null || conservadas == null || nuevas == null)
                return Result.Failure("COBERTURA_ENTRADA_INVALIDA");

            var consentidas = asignaturasVacantesConsentidas as ISet<int>
                              ?? new HashSet<int>(asignaturasVacantesConsentidas ?? Enumerable.Empty<int>());

            // El límite es la ventana semestral seleccionada, no una fecha fija. El manifiesto ya garantiza
            // Inicio <= Fin en su constructor, así que aquí sólo se contrastan las sesiones contra él.
            var requeridas = manifiesto.Relaciones.ToDictionary(r => r.Clave.Valor, StringComparer.Ordinal);
            var extras = new HashSet<string>(manifiesto.RelacionesExtra.Select(k => k.Valor), StringComparer.Ordinal);
            var sesiones = conservadas.Concat(nuevas).ToList();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var posiciones = new HashSet<string>(StringComparer.Ordinal);
            bool hayInconsistencia = false;
            foreach (var s in sesiones)
            {
                if (s == null || s.Clave == null) return Result.Failure("RELACION_INESPERADA");
                // Decisión humana 2026-11-15 «B: extras como ocupación fija»: una extra declarada sólo se
                // acepta conservada; nunca se replantea ni suma demanda requerida.
                if (!requeridas.ContainsKey(s.Clave.Valor)
                    && !(extras.Contains(s.Clave.Valor) && s.Origen == "CONSERVADA"))
                    return Result.Failure("RELACION_INESPERADA");
                if ((s.Origen == "CONSERVADA") != conservadas.Contains(s) ||
                    (s.Origen == "NUEVA") != nuevas.Contains(s))
                    return Result.Failure("ORIGEN_SESION_INVALIDO");
                if (s.Id <= 0 || !ids.Add(s.Origen + ":" + s.Id) ||
                    !posiciones.Add(AgOrigenCanonico.Campos(s.Clave.Valor, s.Fecha.Date, s.Inicio, s.Fin)))
                    return Result.Failure("SESION_DUPLICADA");
                if (s.Fecha.Date < manifiesto.Inicio || s.Fecha.Date > manifiesto.Fin)
                    return Result.Failure("SESION_FUERA_DE_VENTANA");
                if (festivos != null && festivos.Any(f => s.Fecha.Date >= f.Inicio.Date && s.Fecha.Date <= f.Fin.Date))
                    return Result.Failure("SESION_EN_FESTIVO");
                if (s.Inicio < TimeSpan.Zero || s.Fin > TimeSpan.FromDays(1) || s.Fin <= s.Inicio ||
                    s.Horas <= 0 || double.IsNaN(s.Horas) || double.IsInfinity(s.Horas) ||
                    s.Horas > (s.Fin - s.Inicio).TotalHours + 0.001)
                    return Result.Failure("SESION_NO_VERIFICABLE");
                // FASE 1920 / FASE 1925 / FASE 1926: Diagnóstico completo de vacantes y detección de inconsistencias
                // Una vacante declarada es resultado válido de dominio; solo falla si falta docente sin declarar vacante.
                bool consentida = consentidas.Contains(s.Clave.IdAsignatura);
                bool esVacanteDeclarada = s.EsVacante || (idTerceroVacante > 0 && s.IdTercero == idTerceroVacante);
                if (s.IdTercero <= 0 && !s.EsVacante)
                {
                    hayInconsistencia = true;
                    diagnosticos?.Add(new AgDetalleVacanteCobertura
                    {
                        ClaveValor = s.Clave.Valor,
                        IdTerritorial = s.Clave.IdTerritorial,
                        Subunidad = s.Clave.Subunidad,
                        IdPeriodo = s.Clave.IdPeriodo,
                        IdSede = s.Clave.IdSede,
                        IdPrograma = s.Clave.IdPrograma,
                        IdAula = s.Clave.IdAula,
                        IdAsignatura = s.Clave.IdAsignatura,
                        Semestre = s.Clave.Semestre,
                        Grupo = s.Clave.Grupo,
                        MotivoVacante = string.IsNullOrWhiteSpace(s.MotivoVacante) ? "SIN_MARCA_VACANTE" : s.MotivoVacante,
                        Fecha = s.Fecha,
                        Horas = s.Horas,
                        Consentida = false
                    });
                }
                else if (esVacanteDeclarada)
                {
                    diagnosticos?.Add(new AgDetalleVacanteCobertura
                    {
                        ClaveValor = s.Clave.Valor,
                        IdTerritorial = s.Clave.IdTerritorial,
                        Subunidad = s.Clave.Subunidad,
                        IdPeriodo = s.Clave.IdPeriodo,
                        IdSede = s.Clave.IdSede,
                        IdPrograma = s.Clave.IdPrograma,
                        IdAula = s.Clave.IdAula,
                        IdAsignatura = s.Clave.IdAsignatura,
                        Semestre = s.Clave.Semestre,
                        Grupo = s.Clave.Grupo,
                        MotivoVacante = s.MotivoVacante,
                        Fecha = s.Fecha,
                        Horas = s.Horas,
                        Consentida = consentida
                    });
                }
            }
            for (int i = 0; i < sesiones.Count; i++)
                for (int j = i + 1; j < sesiones.Count; j++)
                {
                    var a = sesiones[i]; var b = sesiones[j];
                    if (a.Fecha.Date != b.Fecha.Date || a.Inicio >= b.Fin || b.Inicio >= a.Fin) continue;
                    if (a.Clave.IdAula == b.Clave.IdAula) return Result.Failure("CRUCE_AULA");
                    bool aVacante = a.EsVacante || (idTerceroVacante > 0 && a.IdTercero == idTerceroVacante) || a.IdTercero <= 0;
                    bool bVacante = b.EsVacante || (idTerceroVacante > 0 && b.IdTercero == idTerceroVacante) || b.IdTercero <= 0;
                    if (!aVacante && !bVacante && a.IdTercero == b.IdTercero) return Result.Failure("CRUCE_DOCENTE");
                }
            foreach (var r in manifiesto.Relaciones)
            {
                double horas = sesiones.Where(s => s.Clave.Valor == r.Clave.Valor).Sum(s => s.Horas);
                if (horas < r.HorasEfectivas - 0.001) return Result.Failure("COBERTURA_INCOMPLETA:" + r.Clave.Valor);
                if (horas > r.HorasEfectivas + 0.001) return Result.Failure("COBERTURA_EXCESIVA:" + r.Clave.Valor);
            }
            if (hayInconsistencia) return Result.Failure("SESION_SIN_DOCENTE");
            return Result.Success();
        }
    }

    /// <summary>Proyección inmutable identificada por origen; no es una fila para insertar en staging.</summary>
    public sealed class AgSesionCobertura
    {
        /// <summary>Identidad requerida.</summary>
        public AgClaveRelacionCobertura Clave { get; }
        /// <summary>CONSERVADA o NUEVA.</summary>
        public string Origen { get; }
        /// <summary>ID canónico para conservadas, ID staging para nuevas.</summary>
        public long Id { get; }
        /// <summary>Docente real, o centinela de vacante.</summary>
        public int IdTercero { get; }
        /// <summary>Fecha real.</summary>
        public DateTime Fecha { get; }
        /// <summary>Inicio ocupado.</summary>
        public TimeSpan Inicio { get; }
        /// <summary>Fin ocupado.</summary>
        public TimeSpan Fin { get; }
        /// <summary>Horas lectivas.</summary>
        public double Horas { get; }
        /// <summary>La falta de docente nunca exonera la demanda.</summary>
        public bool EsVacante { get; }
        /// <summary>Motivo de vacante diagnosticado por el AG.</summary>
        public string MotivoVacante { get; }
        /// <summary>Copia los valores sin mutar el origen.</summary>
        public AgSesionCobertura(AgClaveRelacionCobertura clave, string origen, long id, int idTercero,
            DateTime fecha, TimeSpan inicio, TimeSpan fin, double horas, bool esVacante = false, string motivoVacante = null)
        {
            Clave = clave; Origen = origen; Id = id; IdTercero = idTercero;
            Fecha = fecha; Inicio = inicio; Fin = fin; Horas = horas; EsVacante = esVacante;
            MotivoVacante = motivoVacante ?? string.Empty;
        }
    }
}
