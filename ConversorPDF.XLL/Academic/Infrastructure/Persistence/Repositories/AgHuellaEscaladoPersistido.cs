using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>Instantánea SQL por contenido. La lectura y la comparación se realizan en la transacción del escalado.</summary>
    internal static class AgHuellaEscaladoPersistido
    {
        private static readonly ConcurrentDictionary<string, HashSet<string>> _cacheColumnas =
            new ConcurrentDictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        private static HashSet<string> ObtenerColumnas(IDbConnection db, IDbTransaction tx, string tabla)
        {
            try
            {
                int version = db.ExecuteScalar<int>("PRAGMA schema_version;", transaction: tx);
                string key = $"{db.ConnectionString}|{version}|{tabla}";
                return _cacheColumnas.GetOrAdd(key, _ =>
                    new HashSet<string>(db.Query<string>($"SELECT name FROM pragma_table_info('{tabla}');", transaction: tx),
                        StringComparer.OrdinalIgnoreCase));
            }
            catch
            {
                return new HashSet<string>(db.Query<string>($"SELECT name FROM pragma_table_info('{tabla}');", transaction: tx),
                    StringComparer.OrdinalIgnoreCase);
            }
        }

        internal static string Calcular(IDbConnection db, IDbTransaction tx, string referencia,
            string subunidad, int perfil, string destino = null, IReadOnlyCollection<int> liberar = null,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificaciones = null)
        {
            var aulas = LeerAulas(db, tx, referencia, subunidad);
            if (aulas.Count != 1) { throw new InvalidOperationException("Configuración modificada: aula ausente o ambigua."); }
            long id = Convert.ToInt64(((IDictionary<string, object>)aulas[0])["id"], CultureInfo.InvariantCulture);
            var partes = new List<string> { "escalado-sql-v1" };
            Action<string, IEnumerable<dynamic>> agregar = (tabla, filas) =>
            {
                var registros = new List<string>();
                foreach (var fila in filas)
                {
                    var valores = new Dictionary<string, object>((IDictionary<string, object>)fila, StringComparer.OrdinalIgnoreCase);
                    if (destino != null && tabla == "ag_aula") { valores["dias_ventana"] = destino; }
                    if (destino != null && tabla == "ag_arbol_nodo" && perfil > 0
                        && Convert.ToInt32(valores["idarbol"], CultureInfo.InvariantCulture) == perfil
                        && SubunidadCoincide(valores, subunidad)) { valores["dias_ventana"] = destino; }
                    if (tabla == "ag_aula_asignatura" && liberar != null
                        && liberar.Contains(Convert.ToInt32(valores["id_asignatura"], CultureInfo.InvariantCulture))
                        && valores.ContainsKey("origen_patron") && Normalizar(valores["origen_patron"]) == "AUTOMATICO")
                    {
                        valores["dias_patron"] = string.Empty;
                    }
                    if (tabla == "ag_aula_asignatura" && reclasificaciones != null)
                    {
                        int idAsigVal = Convert.ToInt32(valores["id_asignatura"], CultureInfo.InvariantCulture);
                        var rec = reclasificaciones.FirstOrDefault(r => r.IdAsignatura == idAsigVal
                            && r.Autorizado && r.IdAula == id
                            && string.Equals(r.AulaRef, referencia, StringComparison.Ordinal)
                            && string.Equals(r.Subunidad ?? string.Empty, subunidad ?? string.Empty, StringComparison.Ordinal));
                        if (rec != null)
                        {
                            valores["origen_patron"] = rec.OrigenPropuesto;
                            valores["dias_patron"] = rec.DiasPatronPropuesto ?? string.Empty;
                        }
                    }
                    registros.Add(AgHuellaPlanVentana.Hash(valores.OrderBy(v => v.Key, StringComparer.Ordinal)
                        .SelectMany(v => new[] { v.Key, Texto(v.Value) })));
                }
                partes.Add(tabla);
                partes.AddRange(registros.OrderBy(x => x, StringComparer.Ordinal));
            };
            var colsRel = ObtenerColumnas(db, tx, "ag_aula_asignatura");
            var colsRelDeseadas = new[] { "id_aula", "id_asignatura", "exclusivo", "prioridad", "horas_objetivo", "orden_curricular", "requerida", "origen_patron", "dias_patron", "subunidad" };
            var colsRelPresentes = colsRelDeseadas.Where(c => colsRel.Contains(c));
            string sqlRel = $"SELECT {string.Join(",", colsRelPresentes)} FROM ag_aula_asignatura WHERE id_aula=@Id;";

            var colsNodo = ObtenerColumnas(db, tx, "ag_arbol_nodo");
            var colsNodoDeseadas = new[] { "idarbol", "aula_ref", "subunidad", "jornada", "horas_dia", "dias_ventana", "habilitado", "capacidad", "tipo_aula", "es_virtual", "caracteristicas_json", "codigo", "nombre", "idperiodo", "idsede", "idprograma", "semestre", "grupo", "modalidad" };
            var colsNodoPresentes = colsNodoDeseadas.Where(c => colsNodo.Contains(c));
            string sqlNodo = $"SELECT {string.Join(",", colsNodoPresentes)} FROM ag_arbol_nodo WHERE aula_ref=@Ref;";

            agregar("ag_aula", aulas);
            agregar("ag_aula_asignatura", db.Query(sqlRel, new { Id = id }, tx));
            var tablas = new HashSet<string>(db.Query<string>("SELECT name FROM sqlite_master WHERE type='table';", transaction: tx), StringComparer.Ordinal);
            if (tablas.Contains("ag_arbol_nodo"))
            {
                agregar("ag_arbol_nodo", db.Query(sqlNodo, new { Ref = referencia }, tx));
            }

            var catalogoCols = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "asignaturas", new[] { "id", "idprograma", "semestre", "num_nivel",
                    "cod_unidad", "cod_materia", "uni_teorica", "metodologia", "abreviatura",
                    "nucleo", "nombre", "creditos", "int_horaria", "tipoasignatura",
                    "tipocredito", "prerequisito", "subunidad" } },
                { "programas", new[] { "id", "modalidad", "nombre" } },
                { "ag_bloque_horario", new[] { "id", "activo", "dia_semana", "franja",
                    "duracion_horas", "hora_inicio", "hora_fin", "subunidad" } },
                { "festivos", new[] { "id", "fecha" } },
                { "presenciales", new[] { "id", "fecha" } },
                { "periodos", new[] { "id", "fecha_inicio", "fecha_fin", "nombre" } },
                { "ag_arbol", new[] { "id", "idperiodo", "nombre", "fecha_inicio", "fecha_fin",
                    "ventana_origen", "manual_ventana", "poblacion", "generaciones",
                    "tasa_cruce", "tasa_mutacion", "elitismo", "estancamiento_max",
                    "semilla", "subunidad", "activo" } }
            };

            foreach (var kvp in catalogoCols)
            {
                string tabla = kvp.Key;
                partes.Add(tabla + ":" + tablas.Contains(tabla));
                // Nombres constantes, nunca datos procedentes de la propuesta.
                if (tablas.Contains(tabla))
                {
                    var colsInfo = ObtenerColumnas(db, tx, tabla);
                    var colsPresentes = kvp.Value.Where(c => colsInfo.Contains(c));
                    string sqlCatalogo = colsPresentes.Any() ? $"SELECT {string.Join(",", colsPresentes)} FROM {tabla};" : $"SELECT * FROM {tabla};";
                    agregar(tabla, db.Query(sqlCatalogo, transaction: tx));
                }
            }
            return AgHuellaPlanVentana.Hash(partes);
        }

        internal static bool CoincideOrigen(IDbConnection db, IDbTransaction tx, AgContexto contexto, AgAula aula)
        {
            var aulas = LeerAulas(db, tx, aula.AulaRef, aula.Subunidad);
            if (aulas.Count != 1) { return false; }
            var fila = (IDictionary<string, object>)aulas[0];
            if (Convert.ToInt32(fila["id"], CultureInfo.InvariantCulture) != aula.Id
                || Normalizar(fila["dias_ventana"]) != Normalizar(aula.DiasVentana)
                || Normalizar(fila["jornada"]) != Normalizar(aula.Jornada)
                || Convert.ToInt32(fila["horas_dia"], CultureInfo.InvariantCulture) != aula.HorasDia) { return false; }
            if (fila.TryGetValue("estado", out var estadoVal) && estadoVal != null && !(estadoVal is DBNull))
            {
                if (Convert.ToInt32(estadoVal, CultureInfo.InvariantCulture) != (int)aula.Estado) { return false; }
            }
            var colsRel = ObtenerColumnas(db, tx, "ag_aula_asignatura");
            var colsRelDeseadas = new[] { "id_aula", "id_asignatura", "exclusivo", "prioridad", "horas_objetivo", "orden_curricular", "requerida", "origen_patron", "dias_patron", "subunidad" };
            var colsRelPresentes = colsRelDeseadas.Where(c => colsRel.Contains(c));
            string sqlRel = $"SELECT {string.Join(",", colsRelPresentes)} FROM ag_aula_asignatura WHERE id_aula=@Id;";

            var actuales = db.Query(sqlRel, new { Id = aula.Id }, tx)
                .Select(x => (IDictionary<string, object>)x).ToList();
            var esperadas = contexto.AulaAsignaturaPorAulaAsig.Values.Where(x => x.IdAula == aula.Id).ToList();
            if (actuales.Count != esperadas.Count || actuales.GroupBy(x => Convert.ToInt32(x["id_asignatura"], CultureInfo.InvariantCulture)).Any(g => g.Count() != 1)) { return false; }
            return esperadas.All(e => actuales.Any(x =>
                Convert.ToInt32(x["id_asignatura"], CultureInfo.InvariantCulture) == e.IdAsignatura
                && (!x.TryGetValue("exclusivo", out var exVal) || exVal == null || exVal is DBNull || Convert.ToInt32(exVal, CultureInfo.InvariantCulture) == e.Exclusivo)
                && (!x.TryGetValue("prioridad", out var prioVal) || prioVal == null || prioVal is DBNull || Convert.ToInt32(prioVal, CultureInfo.InvariantCulture) == e.Prioridad)
                && (!x.TryGetValue("requerida", out var reqVal) || reqVal == null || reqVal is DBNull || Convert.ToInt32(reqVal, CultureInfo.InvariantCulture) == e.Requerida)
                && Convert.ToDouble(x["horas_objetivo"], CultureInfo.InvariantCulture) == e.HorasObjetivo
                && Convert.ToInt32(x["orden_curricular"], CultureInfo.InvariantCulture) == e.OrdenCurricular
                && Normalizar(x["origen_patron"]) == Normalizar(e.OrigenPatron)
                && Normalizar(x["dias_patron"]) == Normalizar(e.DiasPatron)
                && (!x.TryGetValue("subunidad", out var subVal) || subVal == null || subVal is DBNull || Normalizar(subVal) == Normalizar(e.Subunidad))));
        }

        internal static bool EstadoCompleto(IDbConnection db, IDbTransaction tx, string referencia,
            string subunidad, int perfil, string destino, IReadOnlyCollection<int> liberar,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificaciones = null)
        {
            var aulas = LeerAulas(db, tx, referencia, subunidad);
            if (aulas.Count != 1) { return false; }
            if (destino != null)
            {
                if (!string.Equals(Convert.ToString(((IDictionary<string, object>)aulas[0])["dias_ventana"]), destino, StringComparison.Ordinal)) { return false; }
                if (perfil > 0)
                {
                    var nodos = db.Query<string>(@"SELECT dias_ventana FROM ag_arbol_nodo
WHERE idarbol=@Perfil AND aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub;",
                        new { Perfil = perfil, Ref = referencia, Sub = subunidad ?? string.Empty }, tx).ToList();
                    if (nodos.Count != 1 || !string.Equals(nodos[0], destino, StringComparison.Ordinal)) { return false; }
                }
            }
            long id = Convert.ToInt64(((IDictionary<string, object>)aulas[0])["id"], CultureInfo.InvariantCulture);
            if (liberar != null && liberar.Count > 0)
            {
                int pendientes = db.QuerySingle<int>(@"SELECT COUNT(*) FROM ag_aula_asignatura
WHERE id_aula=@Id AND id_asignatura IN @Ids AND UPPER(COALESCE(origen_patron,''))='AUTOMATICO'
AND COALESCE(dias_patron,'')<>'';", new { Id = id, Ids = liberar }, tx);
                if (pendientes > 0) { return false; }
            }
            if (reclasificaciones != null && reclasificaciones.Count > 0)
            {
                foreach (var rec in reclasificaciones)
                {
                    if (rec == null || !rec.Autorizado || rec.IdAula != id
                        || !string.Equals(rec.AulaRef, referencia, StringComparison.Ordinal)
                        || !string.Equals(rec.Subunidad ?? string.Empty, subunidad ?? string.Empty, StringComparison.Ordinal))
                    {
                        return false;
                    }
                    int ok = db.QuerySingle<int>(@"SELECT COUNT(*) FROM ag_aula_asignatura
WHERE id_aula=@Id AND id_asignatura=@IdAsig
AND UPPER(COALESCE(origen_patron,''))=@OrigenPropuesto
AND COALESCE(dias_patron,'')=@DiasPropuesto;",
                        new
                        {
                            Id = id,
                            IdAsig = rec.IdAsignatura,
                            OrigenPropuesto = (rec.OrigenPropuesto ?? string.Empty).Trim().ToUpperInvariant(),
                            DiasPropuesto = rec.DiasPatronPropuesto ?? string.Empty
                        }, tx);
                    if (ok != 1) { return false; }
                }
            }
            return true;
        }

        private static List<dynamic> LeerAulas(IDbConnection db, IDbTransaction tx, string referencia, string subunidad)
        {
            var colsAula = ObtenerColumnas(db, tx, "ag_aula");
            var colsDeseadas = new[] { "id", "aula_ref", "subunidad", "jornada", "horas_dia", "dias_ventana", "estado", "capacidad", "tipo_aula", "es_virtual", "caracteristicas_json", "fecha_inicio", "fecha_fin", "semestre_inicio", "semestre_fin", "ventana_origen" };
            var colsPresentes = colsDeseadas.Where(c => colsAula.Contains(c));
            string sqlAula = $"SELECT {string.Join(",", colsPresentes)} FROM ag_aula WHERE aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub;";
            return db.Query(sqlAula, new { Ref = referencia, Sub = subunidad ?? string.Empty }, tx).ToList();
        }

        private static bool SubunidadCoincide(IDictionary<string, object> fila, string subunidad)
        {
            object valor;
            if (!fila.TryGetValue("subunidad", out valor) || string.IsNullOrWhiteSpace(Convert.ToString(valor))) { return true; }
            return string.Equals(Convert.ToString(valor).Trim(), subunidad ?? string.Empty, StringComparison.Ordinal);
        }

        private static string Normalizar(object valor)
        {
            return valor == null || valor is DBNull ? string.Empty : Convert.ToString(valor, CultureInfo.InvariantCulture).Trim().ToUpperInvariant();
        }

        private static string Texto(object valor)
        {
            if (valor == null || valor is DBNull) { return null; }
            string dato = valor is byte[]? Convert.ToBase64String((byte[])valor)
                : valor is DateTime ? ((DateTime)valor).ToString("O", CultureInfo.InvariantCulture)
                : valor is double ? ((double)valor).ToString("R", CultureInfo.InvariantCulture)
                : valor is float ? ((float)valor).ToString("R", CultureInfo.InvariantCulture)
                : Convert.ToString(valor, CultureInfo.InvariantCulture);
            return valor.GetType().FullName + ":" + dato;
        }
    }
}
