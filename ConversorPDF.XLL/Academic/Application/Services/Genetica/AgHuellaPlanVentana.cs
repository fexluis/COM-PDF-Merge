using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Huella versionada de entradas por contenido para invalidar planes obsoletos.</summary>
    internal static class AgHuellaPlanVentana
    {
        internal static string Calcular(AgContexto contexto, AgAula aula,
            string ventana = null, IReadOnlyCollection<int> patronesLiberados = null,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificaciones = null,
            IEnumerable<Dictionary<string, object>> planFechado = null)
        {
            var partes = new List<string> { "plan-ventana-v3", Entidad(contexto.Ventana),
                Entidad(aula, new Dictionary<string, object> { { "DiasVentana", ventana ?? aula.DiasVentana } }) };
            partes.AddRange(contexto.Programas.Select(p => "programa:" + Entidad(p)).OrderBy(x => x, StringComparer.Ordinal));
            partes.AddRange(contexto.Asignaturas.Select(a => "asignatura:" + Entidad(a)).OrderBy(x => x, StringComparer.Ordinal));
            partes.AddRange(contexto.Bloques.Select(b => "bloque:" + Entidad(b)).OrderBy(x => x, StringComparer.Ordinal));
            partes.AddRange(contexto.AulaAsignaturaPorAulaAsig.Values.Where(a => a.IdAula == aula.Id)
                .Select(a =>
                {
                    Dictionary<string, object> dict = null;
                    if (patronesLiberados != null && patronesLiberados.Contains(a.IdAsignatura)
                        && string.Equals(a.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                    {
                        dict = new Dictionary<string, object> { { "DiasPatron", string.Empty } };
                    }
                    if (reclasificaciones != null)
                    {
                        var rec = reclasificaciones.FirstOrDefault(r => r.IdAsignatura == a.IdAsignatura
                            && r.Autorizado && r.IdAula == aula.Id
                            && string.Equals(r.AulaRef, aula.AulaRef, StringComparison.Ordinal)
                            && string.Equals(r.Subunidad ?? string.Empty, aula.Subunidad ?? string.Empty, StringComparison.Ordinal));
                        if (rec != null)
                        {
                            if (dict == null) { dict = new Dictionary<string, object>(); }
                            dict["OrigenPatron"] = rec.OrigenPropuesto;
                            dict["DiasPatron"] = rec.DiasPatronPropuesto ?? string.Empty;
                        }
                    }
                    return "relacion:" + Entidad(a, dict);
                })
                .OrderBy(x => x, StringComparer.Ordinal));
            partes.AddRange(contexto.Festivos.OrderBy(f => f.Inicio).ThenBy(f => f.Fin).Select(f =>
                "festivo:" + f.Inicio.ToString("O", CultureInfo.InvariantCulture) + "/" + f.Fin.ToString("O", CultureInfo.InvariantCulture)));
            if (contexto.Ventana != null)
            {
                for (var fecha = contexto.Ventana.FechaInicio.Date; fecha <= contexto.Ventana.FechaFin.Date;)
                {
                    partes.Add(fecha.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ":" + contexto.EsFestivo(fecha));
                    if (fecha == contexto.Ventana.FechaFin.Date) { break; }
                    fecha = fecha.AddDays(1);
                }
            }

            if (planFechado != null)
            {
                foreach (var pf in planFechado)
                {
                    if (pf == null) continue;
                    long idGen = pf.ContainsKey("idGen") ? Convert.ToInt64(pf["idGen"]) : 0;
                    int idAsig = pf.ContainsKey("idAsignatura") ? Convert.ToInt32(pf["idAsignatura"]) : 0;
                    int idAula = pf.ContainsKey("idAula") ? Convert.ToInt32(pf["idAula"]) : 0;
                    int idBloque = pf.ContainsKey("idBloqueHorario") ? Convert.ToInt32(pf["idBloqueHorario"]) : 0;
                    var fechasList = pf.ContainsKey("fechas") ? pf["fechas"] as System.Collections.IEnumerable : null;
                    var fechasStr = fechasList != null
                        ? string.Join(",", fechasList.Cast<object>().Select(x => x?.ToString()).Where(s => !string.IsNullOrEmpty(s)).OrderBy(s => s, StringComparer.Ordinal))
                        : string.Empty;
                    partes.Add(string.Format(CultureInfo.InvariantCulture, "plan:{0}:{1}:{2}:{3}:{4}", idGen, idAsig, idAula, idBloque, fechasStr));
                }
            }

            return Hash(partes);
        }

        internal static string Hash(IEnumerable<string> partes)
        {
            var texto = new StringBuilder();
            foreach (string parte in partes)
            {
                if (parte == null) { texto.Append("-1:"); continue; }
                texto.Append(parte.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(parte);
            }
            using (var sha = SHA256.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(texto.ToString())));
            }
        }

        private static string Entidad(object entidad, IDictionary<string, object> reemplazos = null)
        {
            if (entidad == null) { return Hash(new string[] { null }); }
            var valores = new List<string>();
            foreach (var propiedad in entidad.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0).OrderBy(p => p.Name, StringComparer.Ordinal))
            {
                var tipo = Nullable.GetUnderlyingType(propiedad.PropertyType) ?? propiedad.PropertyType;
                if (!(tipo.IsPrimitive || tipo.IsEnum || tipo == typeof(string) || tipo == typeof(decimal)
                    || tipo == typeof(DateTime) || tipo == typeof(TimeSpan))) { continue; }
                object valor;
                if (reemplazos == null || !reemplazos.TryGetValue(propiedad.Name, out valor))
                {
                    valor = propiedad.GetValue(entidad, null);
                }
                valores.Add(propiedad.Name);
                valores.Add(valor == null ? null : valor is DateTime
                    ? ((DateTime)valor).ToString("O", CultureInfo.InvariantCulture)
                    : valor is double ? ((double)valor).ToString("R", CultureInfo.InvariantCulture)
                    : valor is float ? ((float)valor).ToString("R", CultureInfo.InvariantCulture)
                    : Convert.ToString(valor, CultureInfo.InvariantCulture));
            }
            return Hash(valores);
        }
    }
}
