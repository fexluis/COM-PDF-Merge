using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V66: Renombrado canónico de modalidades de tutoría en bases de datos y staging:
    /// 'Virtual' -> 'Plataforma' y 'Semipresencial' / 'A Distancia' -> 'Mixta'.
    /// Actualiza tablas vinculaciones, ag_vinculaciones, ag_aula y ag_arbol_nodo, incluyendo identificadores derivados (aula_ref, codigo, nombre).
    /// </summary>
    public class MigrationV66_RenombrarModalidadTutoria : IMigration, IMigrationTolerante
    {
        public int Version => 66;
        public string Name => "RenombrarModalidadTutoria";
        public bool Aditiva => true;

        private static readonly (string Token, string Destino)[] TokensModalidad =
        {
            ("|VIRTUAL",        "|PLATAFORMA"),
            ("|Virtual",        "|PLATAFORMA"),
            ("|SEMIPRESENCIAL", "|MIXTA"),
            ("|Semipresencial", "|MIXTA"),
            ("|A DISTANCIA",    "|MIXTA"),
            ("|A Distancia",    "|MIXTA")
        };

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V66] Iniciando migración de modalidades de tutoría (Plataforma y Mixta)...");

            int fallos = 0;
            var errores = new List<Exception>();

            try
            {
                MigrarTablaSimple(connection, transaction, "vinculaciones", "modalidadVinculacion");
            }
            catch (Exception ex)
            {
                fallos++;
                errores.Add(ex);
                LoggerHelper.Error("[V66] Bloque vinculaciones falló.", ex);
            }

            try
            {
                MigrarTablaSimple(connection, transaction, "ag_vinculaciones", "modalidadVinculacion");
            }
            catch (Exception ex)
            {
                fallos++;
                errores.Add(ex);
                LoggerHelper.Error("[V66] Bloque ag_vinculaciones falló.", ex);
            }

            try
            {
                MigrarAgAula(connection, transaction);
            }
            catch (Exception ex)
            {
                fallos++;
                errores.Add(ex);
                LoggerHelper.Error("[V66] Bloque ag_aula falló.", ex);
            }

            try
            {
                MigrarAgArbolNodo(connection, transaction);
            }
            catch (Exception ex)
            {
                fallos++;
                errores.Add(ex);
                LoggerHelper.Error("[V66] Bloque ag_arbol_nodo falló.", ex);
            }

            if (fallos > 0)
            {
                // Aditiva: SchemaMigrator no registrará v66 y se reintentará en el próximo arranque;
                // se prefiere reinicio completo antes que dejar AulaRef inconsistentes de forma permanente.
                Exception causa = errores.Count == 1 ? errores[0] : new AggregateException(errores);
                throw new InvalidOperationException($"[V66] {fallos} bloque(s) fallaron; no se registra la versión 66.", causa);
            }

            LoggerHelper.Info("[V66] Migración V66 finalizada exitosamente.");
        }

        private static bool ExisteTabla(IDbConnection connection, IDbTransaction transaction, string tabla)
        {
            int count = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@Tabla;",
                new { Tabla = tabla },
                transaction: transaction);
            return count > 0;
        }

        private static bool ExisteColumna(IDbConnection connection, IDbTransaction transaction, string tabla, string columna)
        {
            int count = connection.QueryFirstOrDefault<int>(
                $"SELECT COUNT(*) FROM pragma_table_info('{tabla}') WHERE name='{columna}';",
                transaction: transaction);
            return count > 0;
        }

        private static void MigrarTablaSimple(IDbConnection connection, IDbTransaction transaction, string tabla, string columna)
        {
            if (!ExisteTabla(connection, transaction, tabla))
            {
                LoggerHelper.Info($"[V66] Tabla '{tabla}' no existe en la base de datos actual. Omitiendo.");
                return;
            }

            if (!ExisteColumna(connection, transaction, tabla, columna))
            {
                LoggerHelper.Info($"[V66] Columna '{columna}' no existe en '{tabla}'. Omitiendo.");
                return;
            }

            int prevVirtual = connection.QueryFirstOrDefault<int>(
                $"SELECT COUNT(*) FROM {tabla} WHERE UPPER(TRIM(COALESCE({columna},''))) = 'VIRTUAL';",
                transaction: transaction);

            int prevMixta = connection.QueryFirstOrDefault<int>(
                $"SELECT COUNT(*) FROM {tabla} WHERE UPPER(TRIM(COALESCE({columna},''))) IN ('SEMIPRESENCIAL','SEMI PRESENCIAL','A DISTANCIA','DISTANCIA');",
                transaction: transaction);

            LoggerHelper.Info($"[V66] Tabla '{tabla}': Previo - Virtual={prevVirtual}, Semipresencial/Distancia={prevMixta}.");

            int actVirtual = connection.Execute(
                $"UPDATE {tabla} SET {columna} = 'Plataforma' WHERE UPPER(TRIM(COALESCE({columna},''))) = 'VIRTUAL';",
                transaction: transaction);

            int actMixta = connection.Execute(
                $"UPDATE {tabla} SET {columna} = 'Mixta' WHERE UPPER(TRIM(COALESCE({columna},''))) IN ('SEMIPRESENCIAL','SEMI PRESENCIAL','A DISTANCIA','DISTANCIA');",
                transaction: transaction);

            LoggerHelper.Info($"[V66] Tabla '{tabla}': Actualizados - Plataforma={actVirtual}, Mixta={actMixta}.");

            int residual = connection.QueryFirstOrDefault<int>(
                $"SELECT COUNT(*) FROM {tabla} WHERE UPPER(TRIM(COALESCE({columna},''))) IN ('VIRTUAL','SEMIPRESENCIAL','SEMI PRESENCIAL','A DISTANCIA','DISTANCIA');",
                transaction: transaction);

            if (residual > 0)
            {
                LoggerHelper.Warning($"[V66] Advertencia: Se detectaron {residual} filas con valores legados residuales en '{tabla}'.");
            }
        }

        private static void MigrarAgAula(IDbConnection connection, IDbTransaction transaction)
        {
            if (!ExisteTabla(connection, transaction, "ag_aula"))
            {
                LoggerHelper.Info("[V66] Tabla 'ag_aula' no existe. Omitiendo.");
                return;
            }

            // 1. ModalidadVinculacion simple
            MigrarTablaSimple(connection, transaction, "ag_aula", "modalidadVinculacion");

            // 2. Identificadores derivados dinámicos según columnas existentes
            bool tieneAulaRef = ExisteColumna(connection, transaction, "ag_aula", "aula_ref");
            bool tieneCodigo  = ExisteColumna(connection, transaction, "ag_aula", "codigo");
            bool tieneNombre  = ExisteColumna(connection, transaction, "ag_aula", "nombre");

            var cols = new List<string>();
            if (tieneAulaRef) cols.Add("aula_ref");
            if (tieneCodigo)  cols.Add("codigo");
            if (tieneNombre)  cols.Add("nombre");
            if (cols.Count == 0) return;

            var sets = string.Join(", ", cols.Select(c => $"{c} = " + TokensModalidad.Aggregate(c, (acc, tk) =>
                $"REPLACE({acc}, '{tk.Token}', '{tk.Destino}')")));
            var where = string.Join(" OR ", cols.SelectMany(c => TokensModalidad.Select(tk => $"{c} LIKE '%{tk.Token}%'")));

            int actIds = connection.Execute($"UPDATE ag_aula SET {sets} WHERE {where};", transaction: transaction);

            LoggerHelper.Info($"[V66] ag_aula: Referencias actualizadas ({string.Join("/", cols)})={actIds}.");
        }

        private static void MigrarAgArbolNodo(IDbConnection connection, IDbTransaction transaction)
        {
            if (!ExisteTabla(connection, transaction, "ag_arbol_nodo"))
            {
                LoggerHelper.Info("[V66] Tabla 'ag_arbol_nodo' no existe. Omitiendo.");
                return;
            }

            bool tieneAulaRef = ExisteColumna(connection, transaction, "ag_arbol_nodo", "aula_ref");
            bool tieneIdArbol = ExisteColumna(connection, transaction, "ag_arbol_nodo", "idarbol");
            bool tieneModalidad = ExisteColumna(connection, transaction, "ag_arbol_nodo", "modalidad");
            bool tieneCodigo = ExisteColumna(connection, transaction, "ag_arbol_nodo", "codigo");
            bool tieneNombre = ExisteColumna(connection, transaction, "ag_arbol_nodo", "nombre");

            // 1. Deduplicación defensiva en 2 pasos para proteger el índice único uq_ag_arbol_nodo(idarbol, aula_ref)
            if (tieneAulaRef && tieneIdArbol)
            {
                string legacyVirtual = "aula_ref LIKE '%|VIRTUAL%' OR aula_ref LIKE '%|Virtual%'"
                    + (tieneModalidad ? " OR UPPER(TRIM(COALESCE(modalidad,'')))='VIRTUAL'" : "");
                string legacyMixta = "aula_ref LIKE '%|SEMIPRESENCIAL%' OR aula_ref LIKE '%|Semipresencial%'"
                    + " OR aula_ref LIKE '%|A DISTANCIA%' OR aula_ref LIKE '%|A Distancia%'"
                    + (tieneModalidad ? " OR UPPER(TRIM(COALESCE(modalidad,''))) IN ('SEMIPRESENCIAL','SEMI PRESENCIAL','A DISTANCIA','DISTANCIA')" : "");

                // Paso 1: borrar legacy si ya existe equivalente canónica (normalizando ambos lados)
                int borradosVirtualPaso1 = connection.Execute($@"
DELETE FROM ag_arbol_nodo
 WHERE ({legacyVirtual})
   AND EXISTS (
       SELECT 1 FROM ag_arbol_nodo n2
        WHERE n2.idarbol = ag_arbol_nodo.idarbol
          AND n2.id <> ag_arbol_nodo.id
          AND REPLACE(REPLACE(n2.aula_ref,'|VIRTUAL','|PLATAFORMA'),'|Virtual','|PLATAFORMA')
            = REPLACE(REPLACE(ag_arbol_nodo.aula_ref,'|VIRTUAL','|PLATAFORMA'),'|Virtual','|PLATAFORMA')
          AND n2.aula_ref NOT LIKE '%|VIRTUAL%' AND n2.aula_ref NOT LIKE '%|Virtual%'
   );",
                    transaction: transaction);

                // Paso 2: legacy vs legacy equivalente -> conservar la de menor id
                int borradosVirtualPaso2 = connection.Execute($@"
DELETE FROM ag_arbol_nodo
 WHERE ({legacyVirtual})
   AND EXISTS (
       SELECT 1 FROM ag_arbol_nodo n2
        WHERE n2.idarbol = ag_arbol_nodo.idarbol
          AND n2.id < ag_arbol_nodo.id
          AND REPLACE(REPLACE(n2.aula_ref,'|VIRTUAL','|PLATAFORMA'),'|Virtual','|PLATAFORMA')
            = REPLACE(REPLACE(ag_arbol_nodo.aula_ref,'|VIRTUAL','|PLATAFORMA'),'|Virtual','|PLATAFORMA')
   );",
                    transaction: transaction);

                // Paso 1 Mixta: legacy vs canónica equivalente
                int borradosMixtaPaso1 = connection.Execute($@"
DELETE FROM ag_arbol_nodo
 WHERE ({legacyMixta})
   AND EXISTS (
       SELECT 1 FROM ag_arbol_nodo n2
        WHERE n2.idarbol = ag_arbol_nodo.idarbol
          AND n2.id <> ag_arbol_nodo.id
          AND REPLACE(REPLACE(REPLACE(REPLACE(n2.aula_ref, '|SEMIPRESENCIAL', '|MIXTA'), '|Semipresencial', '|MIXTA'), '|A DISTANCIA', '|MIXTA'), '|A Distancia', '|MIXTA')
            = REPLACE(REPLACE(REPLACE(REPLACE(ag_arbol_nodo.aula_ref, '|SEMIPRESENCIAL', '|MIXTA'), '|Semipresencial', '|MIXTA'), '|A DISTANCIA', '|MIXTA'), '|A Distancia', '|MIXTA')
          AND n2.aula_ref NOT LIKE '%|SEMIPRESENCIAL%' AND n2.aula_ref NOT LIKE '%|Semipresencial%'
          AND n2.aula_ref NOT LIKE '%|A DISTANCIA%' AND n2.aula_ref NOT LIKE '%|A Distancia%'
   );",
                    transaction: transaction);

                // Paso 2 Mixta: legacy vs legacy equivalente -> conservar la de menor id
                int borradosMixtaPaso2 = connection.Execute($@"
DELETE FROM ag_arbol_nodo
 WHERE ({legacyMixta})
   AND EXISTS (
       SELECT 1 FROM ag_arbol_nodo n2
        WHERE n2.idarbol = ag_arbol_nodo.idarbol
          AND n2.id < ag_arbol_nodo.id
          AND REPLACE(REPLACE(REPLACE(REPLACE(n2.aula_ref, '|SEMIPRESENCIAL', '|MIXTA'), '|Semipresencial', '|MIXTA'), '|A DISTANCIA', '|MIXTA'), '|A Distancia', '|MIXTA')
            = REPLACE(REPLACE(REPLACE(REPLACE(ag_arbol_nodo.aula_ref, '|SEMIPRESENCIAL', '|MIXTA'), '|Semipresencial', '|MIXTA'), '|A DISTANCIA', '|MIXTA'), '|A Distancia', '|MIXTA')
   );",
                    transaction: transaction);

                int totalBorrados = borradosVirtualPaso1 + borradosVirtualPaso2 + borradosMixtaPaso1 + borradosMixtaPaso2;
                if (totalBorrados > 0)
                {
                    LoggerHelper.Info($"[V66] ag_arbol_nodo: Filas legadas deduplicadas para evitar colisión única. Total={totalBorrados} (Virtual={borradosVirtualPaso1 + borradosVirtualPaso2}, Mixta={borradosMixtaPaso1 + borradosMixtaPaso2}).");
                }
            }
            else if (tieneAulaRef)
            {
                LoggerHelper.Warning("[V66] ag_arbol_nodo: falta columna idarbol; se omite deduplicación previa.");
            }

            // 2. Actualización de modalidad
            int actModVirtual = 0;
            int actModMixta = 0;
            if (tieneModalidad)
            {
                actModVirtual = connection.Execute(
                    "UPDATE ag_arbol_nodo SET modalidad = 'Plataforma' WHERE UPPER(TRIM(COALESCE(modalidad,''))) = 'VIRTUAL';",
                    transaction: transaction);

                actModMixta = connection.Execute(
                    "UPDATE ag_arbol_nodo SET modalidad = 'Mixta' WHERE UPPER(TRIM(COALESCE(modalidad,''))) IN ('SEMIPRESENCIAL','SEMI PRESENCIAL','A DISTANCIA','DISTANCIA');",
                    transaction: transaction);
            }

            // 3. Actualización de aula_ref, codigo y nombre dinámico según columnas existentes
            int actIds = 0;
            var colsNodo = new List<string>();
            if (tieneAulaRef) colsNodo.Add("aula_ref");
            if (tieneCodigo)  colsNodo.Add("codigo");
            if (tieneNombre)  colsNodo.Add("nombre");

            if (colsNodo.Count > 0)
            {
                var sets = string.Join(", ", colsNodo.Select(c => $"{c} = " + TokensModalidad.Aggregate(c, (acc, tk) =>
                    $"REPLACE({acc}, '{tk.Token}', '{tk.Destino}')")));
                var where = string.Join(" OR ", colsNodo.SelectMany(c => TokensModalidad.Select(tk => $"{c} LIKE '%{tk.Token}%'")));

                actIds = connection.Execute($"UPDATE ag_arbol_nodo SET {sets} WHERE {where};", transaction: transaction);
            }

            LoggerHelper.Info($"[V66] ag_arbol_nodo: Modalidades actualizadas (Plataforma={actModVirtual}, Mixta={actModMixta}), referencias={actIds}.");
        }
    }
}
