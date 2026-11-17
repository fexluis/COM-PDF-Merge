using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>Lectura canónica de producción. Aprobado y ag_promocion no certifican consentimiento ni revisión.</summary>
    public sealed class AgPlanAprobadoRepository : IAgPlanAprobadoRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        /// <summary>Usa la misma fábrica de conexión que AgScheduleRepository; no abre bases por rutas propias.</summary>
        public AgPlanAprobadoRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        /// <inheritdoc />
        public async Task<Result<AgPlanAprobadoSnapshot>> ObtenerAsync(AgManifiestoCobertura manifiesto, CancellationToken ct)
        {
            if (manifiesto == null) return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_MANIFIESTO_REQUERIDO");
            try
            {
                ct.ThrowIfCancellationRequested();
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    // Un único snapshot de lectura: no mezclar catálogo, vinculaciones y sesiones de revisiones distintas.
                    using (var tx = db.BeginTransaction())
                        return await ObtenerEnTransaccionAsync(db, tx, manifiesto, ct);
                }
            }
            catch (OperationCanceledException) { return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_CANCELADA"); }
            catch (Exception ex) { return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_ERROR:" + ex.GetType().Name); }
        }

        /// <summary>Lee el mismo origen dentro de la transacción de escritura del confirmador; no inicia otra transacción.</summary>
        internal static async Task<Result<AgPlanAprobadoSnapshot>> ObtenerEnTransaccionAsync(IDbConnection db,
            IDbTransaction tx, AgManifiestoCobertura manifiesto, CancellationToken ct)
        {
            if (db == null || tx == null || manifiesto == null)
                return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_MANIFIESTO_REQUERIDO");
            try
            {
                        ct.ThrowIfCancellationRequested();
                        var claves = manifiesto.Relaciones.Select(r => r.Clave).ToList();
                        var relaciones = new List<AgRelacionPlanLeida>();
                        var extras = new List<AgRelacionPlanLeida>();
                        var descubiertas = new List<AgClaveRelacionCobertura>();
                        var diagnosticos = new List<string>();
                        var catalogoOrigen = new List<string>();
                        var territoriales = await Filas(db, tx, "SELECT * FROM territoriales WHERE id IN @Ids ORDER BY id",
                            new { Ids = claves.Select(k => k.IdTerritorial).Distinct().ToArray() }, ct);
                        catalogoOrigen.AddRange(territoriales.Select(Origen));
                        foreach (var k in claves)
                            if (territoriales.Count(t => Entero(t, "id") == k.IdTerritorial && Texto(t, "subunidad") == k.Subunidad) != 1)
                                diagnosticos.Add("TERRITORIAL_NO_COINCIDE:" + k.Valor);
                        // Una huérfana no tiene período/cohorte atribuible. Se consulta una vez por selección,
                        // nunca se adjunta a un aula. Sólo la partición seleccionada permite leer su fila completa.
                        var particiones = new { Subunidades = claves.Select(k => k.Subunidad).Distinct().ToArray() };
                        var huerfanas = await Filas(db, tx, @"SELECT s.* FROM sesiones s
WHERE UPPER(TRIM(COALESCE(s.subunidad,''))) IN @Subunidades
AND NOT EXISTS (SELECT 1 FROM vinculaciones v WHERE v.id=s.idvinculacion) ORDER BY s.id", particiones, ct);
                        foreach (var s in huerfanas)
                        {
                            diagnosticos.Add("SESION_HUERFANA_SUBUNIDAD_SELECCIONADA:" + Entero(s, "id"));
                            catalogoOrigen.Add(AgOrigenCanonico.Campos("SESION_HUERFANA_SUBUNIDAD_SELECCIONADA",
                                Entero(s, "id"), Entero(s, "idvinculacion"), AgOrigenCanonico.Huella(Origen(s))));
                        }
                        var huerfanasDesconocidas = await Filas(db, tx, @"SELECT s.id, s.idvinculacion FROM sesiones s
WHERE UPPER(TRIM(COALESCE(s.subunidad,''))) NOT IN @Subunidades
AND (TRIM(COALESCE(s.subunidad,''))='' OR NOT EXISTS (SELECT 1 FROM territoriales t
 WHERE UPPER(TRIM(t.subunidad))=UPPER(TRIM(s.subunidad))))
AND NOT EXISTS (SELECT 1 FROM vinculaciones v WHERE v.id=s.idvinculacion) ORDER BY s.id", particiones, ct);
                        RegistrarIncertidumbres(huerfanasDesconocidas, "SESION_HUERFANA_ORIGEN_DESCONOCIDO", diagnosticos, catalogoOrigen);
                        foreach (var alcance in claves.GroupBy(k => new { k.IdPeriodo, k.IdSede, k.IdPrograma, k.Subunidad })
                            .OrderBy(g => g.Key.IdPeriodo).ThenBy(g => g.Key.IdSede).ThenBy(g => g.Key.IdPrograma)
                            .ThenBy(g => g.Key.Subunidad, StringComparer.Ordinal))
                        {
                            ct.ThrowIfCancellationRequested();
                            var parametros = alcance.Key;
                            // Payload completo exclusivamente del tenant seleccionado; no filtrar asignatura
                            // para preservar la detección de extras dentro de esa partición.
                            const string dimensiones = "idperiodo=@IdPeriodo AND idsede=@IdSede AND idprograma=@IdPrograma";
                            const string particion = "UPPER(TRIM(COALESCE(subunidad,'')))=@Subunidad";
                            const string desconocida = @"UPPER(TRIM(COALESCE(subunidad,'')))<>@Subunidad
AND (TRIM(COALESCE(subunidad,''))='' OR NOT EXISTS (SELECT 1 FROM territoriales t
 WHERE UPPER(TRIM(t.subunidad))=UPPER(TRIM(subunidadOrigen.subunidad))))";
                            var aulas = await Filas(db, tx, "SELECT * FROM ag_aula WHERE " + dimensiones + " AND " + particion + " ORDER BY id", parametros, ct);
                            var vinculaciones = await Filas(db, tx, "SELECT * FROM vinculaciones WHERE " + dimensiones + " AND " + particion + " ORDER BY id", parametros, ct);
                            var requeridas = alcance.ToList();
                            // Origen desconocido: sólo identidad mínima para clasificar, nunca horarios/docentes/payload ajeno.
                            const string identidadMinima = "id, semestre, grupo, modalidadVinculacion";
                            var aulasDesconocidas = await Filas(db, tx, "SELECT " + identidadMinima + " FROM ag_aula subunidadOrigen WHERE "
                                + dimensiones + " AND " + desconocida + " ORDER BY id", parametros, ct);
                            var vinculacionesDesconocidas = await Filas(db, tx, "SELECT " + identidadMinima + " FROM vinculaciones subunidadOrigen WHERE "
                                + dimensiones + " AND " + desconocida + " ORDER BY id", parametros, ct);
                            RegistrarIncertidumbres(aulasDesconocidas.Where(a => requeridas.Any(k => Cohorte(a, k))),
                                "AULA_SUBUNIDAD_DESCONOCIDA", diagnosticos, catalogoOrigen);
                            RegistrarIncertidumbres(vinculacionesDesconocidas.Where(v => requeridas.Any(k => Cohorte(v, k))),
                                "VINCULACION_SUBUNIDAD_DESCONOCIDA", diagnosticos, catalogoOrigen);
                            // EXISTS no multiplica sesiones. Las dos identidades deben pertenecer al tenant.
                            var sesiones = await Filas(db, tx, @"SELECT s.* FROM sesiones s
WHERE UPPER(TRIM(COALESCE(s.subunidad,'')))=@Subunidad AND EXISTS (
SELECT 1 FROM vinculaciones v WHERE v.id=s.idvinculacion AND v.idperiodo=@IdPeriodo
AND v.idsede=@IdSede AND v.idprograma=@IdPrograma
AND UPPER(TRIM(COALESCE(v.subunidad,'')))=@Subunidad) ORDER BY s.id", parametros, ct);
                            // Cruces de partición no importan la fila completa de ninguna de las dos partes.
                            var inconsistentes = await Filas(db, tx, @"SELECT s.id, s.idvinculacion,
v.semestre, v.grupo, v.modalidadVinculacion FROM sesiones s
JOIN vinculaciones v ON v.id=s.idvinculacion
WHERE v.idperiodo=@IdPeriodo AND v.idsede=@IdSede AND v.idprograma=@IdPrograma
AND ((UPPER(TRIM(COALESCE(v.subunidad,'')))=@Subunidad AND UPPER(TRIM(COALESCE(s.subunidad,'')))<>@Subunidad)
 OR (UPPER(TRIM(COALESCE(s.subunidad,'')))=@Subunidad AND UPPER(TRIM(COALESCE(v.subunidad,'')))<>@Subunidad))
ORDER BY s.id", parametros, ct);
                            RegistrarIncertidumbres(inconsistentes.Where(s => requeridas.Any(k => Cohorte(s, k))),
                                "SESION_SUBUNIDAD_INCONSISTENTE", diagnosticos, catalogoOrigen);
                            var porVinculacion = sesiones.ToLookup(s => Entero(s, "idvinculacion"));
                            foreach (var a in aulas.Where(a => requeridas.Any(k => Cohorte(a, k))))
                                catalogoOrigen.Add(Origen(a));
                            foreach (var v in vinculaciones)
                            {
                                ct.ThrowIfCancellationRequested();
                                var cohorte = requeridas.Where(k => Texto(v, "subunidad") == k.Subunidad && Cohorte(v, k)).ToList();
                                if (cohorte.Count == 0) continue; // Otra cohorte conocida no es demanda seleccionada.
                                long id = Entero(v, "id");
                                var candidatas = cohorte.Where(k => Exacta(v, k) && Entero(v, "idasignatura") == k.IdAsignatura).ToList();
                                // Decisión humana 2026-11-15 «B: extras como ocupación fija»: una relación productiva
                                // declarada como extra no requerida se conserva inmutable y ocupa aula/docente; no bloquea.
                                var extrasDeclaradas = manifiesto.RelacionesExtra.Where(k => Texto(v, "subunidad") == k.Subunidad
                                    && Exacta(v, k) && Entero(v, "idasignatura") == k.IdAsignatura).ToList();
                                // Sólo vincular materia exclusivamente productiva con identidad completa y aula
                                // única de la cohorte seleccionada. Nunca atribuir a la primera aula compatible.
                                var aulasCompatibles = aulas.Where(a => cohorte.Any(k => Exacta(a, k)))
                                    .Select(a => Entero(a, "id")).Distinct().ToList();
                                if (candidatas.Count == 0 && extrasDeclaradas.Count == 0 && Completa(v)
                                    && Entero(v, "idasignatura") > 0 && Entero(v, "idasignatura") <= int.MaxValue
                                    && aulasCompatibles.Count == 1)
                                {
                                    var aulaOrigen = cohorte.Where(k => k.IdAula == aulasCompatibles[0] && Exacta(v, k)).ToList();
                                    if (aulaOrigen.Count > 0)
                                    {
                                        var baseClave = aulaOrigen[0];
                                        var clave = new AgClaveRelacionCobertura(baseClave.IdTerritorial, baseClave.Subunidad,
                                            baseClave.IdPeriodo, baseClave.IdSede, baseClave.IdPrograma, baseClave.IdAula,
                                            (int)Entero(v, "idasignatura"), baseClave.Semestre, baseClave.Grupo, baseClave.Modalidad);
                                        extrasDeclaradas.Add(clave);
                                        descubiertas.Add(clave);
                                    }
                                }
                                bool esExtra = candidatas.Count == 0 && extrasDeclaradas.Count == 1;
                                bool identidadCompleta = Completa(v);
                                if (!identidadCompleta) diagnosticos.Add("IDENTIDAD_INCOMPLETA:" + id);
                                if (!cohorte.Any(k => Texto(v, "subunidad") == k.Subunidad))
                                    diagnosticos.Add("SUBUNIDAD_NO_COINCIDE:" + id);
                                if (!cohorte.Any(k => k.IdAsignatura == Entero(v, "idasignatura")) && !esExtra)
                                    diagnosticos.Add("RELACION_EXTRA:" + id);
                                if (extrasDeclaradas.Count > 1) diagnosticos.Add("IDENTIDAD_AMBIGUA:" + id);
                                if (aulasCompatibles.Count != 1 || candidatas.Select(k => k.IdAula).Distinct().Count() > 1
                                    || candidatas.Any(k => !aulasCompatibles.Contains(k.IdAula))
                                    || (esExtra && !aulasCompatibles.Contains(extrasDeclaradas[0].IdAula)))
                                    diagnosticos.Add("IDENTIDAD_AMBIGUA:" + id);
                                var leidas = new List<AgSesionCanonica>();
                                foreach (var s in porVinculacion[id])
                                {
                                    ct.ThrowIfCancellationRequested();
                                    DateTime fecha;
                                    DateTime? fechaLeida = DateTime.TryParseExact(Cadena(s, "fechasesion"),
                                        new[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.FFFFFFF" },
                                        CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha) ? fecha : (DateTime?)null;
                                    var franja = HorarioSesionParser.TryParse(Cadena(s, "horariosesion"));
                                    double horas = Numero(s, "horas");
                                    long idSesion = Entero(s, "id");
                                    if (!fechaLeida.HasValue || franja.IsFailure || horas <= 0 || double.IsNaN(horas)
                                        || double.IsInfinity(horas) || Entero(s, "idtercero") <= 0)
                                        diagnosticos.Add("SESION_NO_VERIFICABLE:" + idSesion);
                                    if (Texto(s, "subunidad") != Texto(v, "subunidad"))
                                        diagnosticos.Add("SUBUNIDAD_SESION_NO_COINCIDE:" + idSesion);
                                    leidas.Add(new AgSesionCanonica(idSesion, id, Entero(s, "idtercero"), fechaLeida,
                                        franja.IsSuccess ? (TimeSpan?)franja.Value.Inicio : null,
                                        franja.IsSuccess ? (TimeSpan?)franja.Value.Fin : null,
                                        horas, Cadena(s, "subunidad"), Origen(s)));
                                }
                                var leida = new AgRelacionPlanLeida(id, esExtra ? extrasDeclaradas : candidatas, leidas, Origen(v));
                                if (esExtra) extras.Add(leida); else relaciones.Add(leida);
                            }
                        }
                        foreach (var repetida in relaciones.Concat(extras).SelectMany(r => r.ClavesCandidatas.Select(k => k.Valor))
                            .GroupBy(k => k, StringComparer.Ordinal).Where(g => g.Count() > 1))
                            diagnosticos.Add("VINCULACIONES_MULTIPLES:" + repetida.Key);
                        ct.ThrowIfCancellationRequested();
                        // Sellar extras descubiertas y recibidas en el manifiesto ANTES de construir la
                        // huella del origen: cualquier cambio posterior invalida ratificaciones anteriores.
                        var clavesExtra = manifiesto.RelacionesExtra.Concat(descubiertas)
                            .GroupBy(k => k.Valor, StringComparer.Ordinal).Select(g => g.First()).ToList();
                        var ligado = clavesExtra.Count == manifiesto.RelacionesExtra.Count
                            ? manifiesto : new AgManifiestoCobertura(manifiesto.Inicio, manifiesto.Fin,
                                manifiesto.Relaciones, clavesExtra);
                        return Result.Success(new AgPlanAprobadoSnapshot(ligado, relaciones, diagnosticos,
                            string.Concat(catalogoOrigen.OrderBy(s => s, StringComparer.Ordinal).Select(s => AgOrigenCanonico.Campos(s))),
                            extras));
            }
            catch (OperationCanceledException) { return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_CANCELADA"); }
            catch (Exception ex) { return Result.Failure<AgPlanAprobadoSnapshot>("LECTURA_PLAN_ERROR:" + ex.GetType().Name); }
        }

        private static async Task<List<IDictionary<string, object>>> Filas(IDbConnection db, IDbTransaction tx,
            string sql, object parametros, CancellationToken ct)
        {
            var filas = await db.QueryAsync(new CommandDefinition(sql, parametros, transaction: tx, cancellationToken: ct));
            return filas.Select(f => (IDictionary<string, object>)f).ToList();
        }

        private static void RegistrarIncertidumbres(IEnumerable<IDictionary<string, object>> filas, string tipo,
            List<string> diagnosticos, List<string> origen)
        {
            foreach (var fila in filas)
            {
                diagnosticos.Add(tipo + ":" + Entero(fila, "id"));
                origen.Add(AgOrigenCanonico.Campos(tipo, Origen(fila)));
            }
        }

        private static object Valor(IDictionary<string, object> fila, string campo)
        {
            object valor;
            return fila.TryGetValue(campo, out valor) && valor != DBNull.Value ? valor : null;
        }
        private static string Cadena(IDictionary<string, object> f, string c) => Convert.ToString(Valor(f, c), CultureInfo.InvariantCulture);
        private static string Texto(IDictionary<string, object> f, string c) => AgClaveRelacionCobertura.Texto(Cadena(f, c));
        private static long Entero(IDictionary<string, object> f, string c) => Convert.ToInt64(Valor(f, c), CultureInfo.InvariantCulture);
        private static double Numero(IDictionary<string, object> f, string c) => Valor(f, c) == null ? double.NaN : Convert.ToDouble(Valor(f, c), CultureInfo.InvariantCulture);
        private static string Modalidad(IDictionary<string, object> f) => string.IsNullOrWhiteSpace(Cadena(f, "modalidadVinculacion"))
            ? string.Empty : AgClaveRelacionCobertura.Texto(ModalidadTutoriaCatalog.NormalizarTutoria(Cadena(f, "modalidadVinculacion")));
        private static bool Completa(IDictionary<string, object> f) => !string.IsNullOrWhiteSpace(Cadena(f, "semestre"))
            && !string.IsNullOrWhiteSpace(Cadena(f, "grupo")) && Modalidad(f) != string.Empty && Texto(f, "subunidad") != string.Empty;
        private static bool Cohorte(IDictionary<string, object> f, AgClaveRelacionCobertura k)
        {
            // Campos vacíos no se descartan: se devuelven como ambiguos/incompletos, no como una identidad inventada.
            string semestre = Cadena(f, "semestre"), grupo = Texto(f, "grupo"), modalidad = Modalidad(f);
            return (Texto(f, "subunidad") == string.Empty || Texto(f, "subunidad") == k.Subunidad)
                && (string.IsNullOrWhiteSpace(semestre) || AgFormatoAsignatura.SemestreRomano(semestre) == k.Semestre)
                && (grupo == string.Empty || grupo == k.Grupo) && (modalidad == string.Empty || modalidad == k.Modalidad);
        }
        private static bool Exacta(IDictionary<string, object> f, AgClaveRelacionCobertura k)
            => Completa(f) && Cohorte(f, k) && Texto(f, "subunidad") == k.Subunidad;
        private static string Origen(IDictionary<string, object> fila)
            => string.Concat(fila.OrderBy(p => p.Key, StringComparer.Ordinal)
                .Select(p => AgOrigenCanonico.Campos(p.Key, p.Value)));
    }
}
