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
            int idTerceroVacante = 0, IReadOnlyCollection<(DateTime Inicio, DateTime Fin)> festivos = null)
        {
            if (manifiesto == null || conservadas == null || nuevas == null)
                return Result.Failure("COBERTURA_ENTRADA_INVALIDA");
            if (manifiesto.Inicio < new DateTime(2027, 2, 5) || manifiesto.Fin > new DateTime(2027, 5, 30))
                return Result.Failure("VENTANA_EXCEDE_LIMITE_30_05_2027");
            var requeridas = manifiesto.Relaciones.ToDictionary(r => r.Clave.Valor, StringComparer.Ordinal);
            var sesiones = conservadas.Concat(nuevas).ToList();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var posiciones = new HashSet<string>(StringComparer.Ordinal);
            foreach (var s in sesiones)
            {
                if (s == null || s.Clave == null || !requeridas.ContainsKey(s.Clave.Valor))
                    return Result.Failure("RELACION_INESPERADA");
                if ((s.Origen == "CONSERVADA") != conservadas.Contains(s) ||
                    (s.Origen == "NUEVA") != nuevas.Contains(s))
                    return Result.Failure("ORIGEN_SESION_INVALIDO");
                if (s.Id <= 0 || !ids.Add(s.Origen + ":" + s.Id) ||
                    !posiciones.Add(AgOrigenCanonico.Campos(s.Clave.Valor, s.Fecha.Date, s.Inicio, s.Fin)))
                    return Result.Failure("SESION_DUPLICADA");
                if (s.Fecha.Date < manifiesto.Inicio || s.Fecha.Date > manifiesto.Fin ||
                    s.Fecha.Date < new DateTime(2027, 2, 5) || s.Fecha.Date > new DateTime(2027, 5, 30))
                    return Result.Failure("SESION_FUERA_DE_VENTANA");
                if (festivos != null && festivos.Any(f => s.Fecha.Date >= f.Inicio.Date && s.Fecha.Date <= f.Fin.Date))
                    return Result.Failure("SESION_EN_FESTIVO");
                if (s.Inicio < TimeSpan.Zero || s.Fin > TimeSpan.FromDays(1) || s.Fin <= s.Inicio ||
                    s.Horas <= 0 || double.IsNaN(s.Horas) || double.IsInfinity(s.Horas) ||
                    s.Horas > (s.Fin - s.Inicio).TotalHours + 0.001)
                    return Result.Failure("SESION_NO_VERIFICABLE");
                // VACANTE no equivale a demanda ausente: la cobertura horaria se cuenta y el docente falta.
                if (s.IdTercero <= 0 || s.IdTercero == idTerceroVacante || s.EsVacante)
                    return Result.Failure("DOCENTE_VACANTE");
            }
            for (int i = 0; i < sesiones.Count; i++)
            for (int j = i + 1; j < sesiones.Count; j++)
            {
                var a = sesiones[i]; var b = sesiones[j];
                if (a.Fecha.Date != b.Fecha.Date || a.Inicio >= b.Fin || b.Inicio >= a.Fin) continue;
                if (a.Clave.IdAula == b.Clave.IdAula) return Result.Failure("CRUCE_AULA");
                if (a.IdTercero == b.IdTercero) return Result.Failure("CRUCE_DOCENTE");
            }
            foreach (var r in manifiesto.Relaciones)
            {
                double horas = sesiones.Where(s => s.Clave.Valor == r.Clave.Valor).Sum(s => s.Horas);
                if (horas < r.HorasEfectivas - 0.001) return Result.Failure("COBERTURA_INCOMPLETA:" + r.Clave.Valor);
                if (horas > r.HorasEfectivas + 0.001) return Result.Failure("COBERTURA_EXCESIVA:" + r.Clave.Valor);
            }
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
        /// <summary>Copia los valores sin mutar el origen.</summary>
        public AgSesionCobertura(AgClaveRelacionCobertura clave, string origen, long id, int idTercero,
            DateTime fecha, TimeSpan inicio, TimeSpan fin, double horas, bool esVacante = false)
        {
            Clave = clave; Origen = origen; Id = id; IdTercero = idTercero;
            Fecha = fecha; Inicio = inicio; Fin = fin; Horas = horas; EsVacante = esVacante;
        }
    }
}
