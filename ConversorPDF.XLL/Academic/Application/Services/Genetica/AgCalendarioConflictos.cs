using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
            throw new NotImplementedException();
        }

        internal static bool SeCruzan(Ocupacion a, Ocupacion b)
        {
            throw new NotImplementedException();
        }

        internal static string Cohorte(AgGen g)
        {
            throw new NotImplementedException();
        }

        internal static bool DocenteReal(AgGen g, AgContexto c)
        {
            throw new NotImplementedException();
        }
    }
}
