using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    internal static class AgCalendarioConflictos
    {
        internal sealed class Ocupacion
        {
            internal AgGen Gen;
            internal DateTime Fecha;
            internal TimeSpan Inicio;
            internal TimeSpan Fin;
            internal string Franja;
        }

        internal static Result<List<Ocupacion>> Crear(IEnumerable<AgGen> genes, AgContexto contexto)
        {
            return Crear(genes, contexto, new AgExpansorSesiones());
        }

        internal static Result<List<Ocupacion>> Crear(IEnumerable<AgGen> genes, AgContexto contexto, IAgExpansorSesiones expansor)
        {
            if (genes == null)
            {
                throw new ArgumentNullException(nameof(genes));
            }

            if (contexto == null)
            {
                throw new ArgumentNullException(nameof(contexto));
            }

            if (expansor == null)
            {
                expansor = new AgExpansorSesiones();
            }

            var resultado = new List<Ocupacion>();

            foreach (var gen in genes)
            {
                if (gen == null || gen.HorasGen <= 0)
                {
                    continue;
                }

                if (contexto.BloquesPorId == null ||
                    !contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bloque) ||
                    bloque == null ||
                    bloque.Activo != 1 ||
                    !TimeSpan.TryParse(bloque.HoraInicio, CultureInfo.InvariantCulture, out var ini) ||
                    !TimeSpan.TryParse(bloque.HoraFin, CultureInfo.InvariantCulture, out var fin) ||
                    ini >= fin)
                {
                    return Result.Failure<List<Ocupacion>>($"I6_BLOQUE_INVALIDO: aula={gen.IdAula}, asig={gen.IdAsignatura}");
                }

                var sesiones = expansor.Expandir(gen, 0, contexto);
                if (sesiones.IsFailure)
                {
                    return Result.Failure<List<Ocupacion>>(sesiones.Error);
                }

                foreach (var sesion in sesiones.Value)
                {
                    if (!sesion.FechaSesion.HasValue)
                    {
                        return Result.Failure<List<Ocupacion>>("I5_SESION_SIN_FECHA");
                    }

                    resultado.Add(new Ocupacion
                    {
                        Gen = gen,
                        Fecha = sesion.FechaSesion.Value.Date,
                        Inicio = ini,
                        Fin = fin,
                        Franja = bloque.Franja ?? string.Empty
                    });
                }
            }

            return Result.Success(resultado);
        }

        internal static bool SeCruzan(Ocupacion a, Ocupacion b)
        {
            if (a == null || b == null)
            {
                return false;
            }

            return a.Fecha == b.Fecha && a.Inicio < b.Fin && b.Inicio < a.Fin;
        }

        internal static string Cohorte(AgGen g)
        {
            if (g == null)
            {
                return string.Empty;
            }

            string semestre = AgFormatoAsignatura.SemestreRomano(g.Semestre);
            if (string.IsNullOrWhiteSpace(semestre) || semestre == "SIN ASIGNAR")
            {
                semestre = (g.Semestre ?? string.Empty).Trim();
            }

            return $"{g.IdSede}|{g.IdPrograma}|{semestre}|{(g.Grupo ?? string.Empty).Trim()}";
        }

        internal static bool DocenteReal(AgGen g, AgContexto c)
        {
            if (g == null || c == null)
            {
                return false;
            }

            return g.EsVacante == 0 && g.IdTercero > 0 && g.IdTercero != c.IdTerceroVacante;
        }
    }
}
