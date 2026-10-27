using System;
using System.Data;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V48: Crea la tabla física 'terceros_asignaturas' para el pool docente-asignatura (Capa 0)
    /// con clave primaria de 63 bits compatible con CRDT y ejecuta el backfill histórico desde 'vinculaciones'.
    /// </summary>
    public class MigrationV48_AddTercerosAsignaturas : IMigration
    {
        public int Version => 48;
        public string Name => "AddTercerosAsignaturas";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // 1. Crear tabla terceros_asignaturas (CRDT ready: PK NOT NULL sin AUTOINCREMENT, sin FKs físicas)
            string createTableSql = @"
CREATE TABLE IF NOT EXISTS terceros_asignaturas (
    id                      INTEGER PRIMARY KEY NOT NULL,
    idtercero               INTEGER NOT NULL DEFAULT 0,
    idprograma              INTEGER NOT NULL DEFAULT 0,
    idasignatura            INTEGER NOT NULL DEFAULT 0,
    semestre                TEXT    NOT NULL DEFAULT '',
    activo                  INTEGER NOT NULL DEFAULT 1,
    prioridad               INTEGER NOT NULL DEFAULT 0,
    veces_asignada          INTEGER NOT NULL DEFAULT 0,
    fecha_ultima_asignacion TEXT    NOT NULL DEFAULT '',
    origen                  TEXT    NOT NULL DEFAULT 'MANUAL',
    observaciones           TEXT    NOT NULL DEFAULT '',
    subunidad               TEXT    NOT NULL DEFAULT '',
    created_at              TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at              TEXT    NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_tas_tercero    ON terceros_asignaturas(idtercero);
CREATE INDEX IF NOT EXISTS idx_tas_asignatura ON terceros_asignaturas(idasignatura);
CREATE INDEX IF NOT EXISTS idx_tas_programa   ON terceros_asignaturas(idprograma, semestre);";

            connection.Execute(createTableSql, transaction: transaction);

            // 2. Backfill en C# calculando claves deterministas de 63 bits a partir de vinculaciones históricas
            int hasVinculaciones = connection.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';",
                transaction: transaction);

            if (hasVinculaciones > 0)
            {
                string querySql = @"
SELECT 
    v.idtercero as IdTercero,
    v.idprograma as IdPrograma,
    v.idasignatura as IdAsignatura,
    COALESCE(v.semestre, '') as Semestre,
    COALESCE(v.subunidad, '') as Subunidad,
    COUNT(DISTINCT v.id) as VecesAsignada,
    MAX(v.fecha_inicio) as FechaUltimaAsignacionRaw
FROM vinculaciones v
WHERE v.idtercero > 0 AND v.idasignatura > 0
GROUP BY v.idtercero, v.idprograma, v.idasignatura, COALESCE(v.semestre, ''), COALESCE(v.subunidad, '');";

                var historicos = connection.Query(querySql, transaction: transaction).ToList();

                string insertSql = @"
INSERT INTO terceros_asignaturas (
    id, idtercero, idprograma, idasignatura, semestre, activo, prioridad,
    veces_asignada, fecha_ultima_asignacion, origen, observaciones, subunidad,
    created_at, updated_at
) VALUES (
    @Id, @IdTercero, @IdPrograma, @IdAsignatura, @Semestre, 1, 0,
    @VecesAsignada, @FechaUltimaAsignacion, 'HISTORICO', '', @Subunidad,
    datetime('now'), datetime('now')
)
ON CONFLICT(id) DO UPDATE SET
    veces_asignada = excluded.veces_asignada,
    fecha_ultima_asignacion = excluded.fecha_ultima_asignacion,
    updated_at = datetime('now');";

                foreach (var h in historicos)
                {
                    int idTercero = (int)h.IdTercero;
                    int idPrograma = (int)h.IdPrograma;
                    int idAsignatura = (int)h.IdAsignatura;
                    string semestre = ((string)h.Semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string subunidad = (string)h.Subunidad ?? string.Empty;
                    int veces = (int)h.VecesAsignada;
                    string rawFecha = (string)h.FechaUltimaAsignacionRaw;
                    string fechaIso = NormalizarFechaIso(rawFecha);

                    long id = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre);

                    connection.Execute(insertSql, new
                    {
                        Id = id,
                        IdTercero = idTercero,
                        IdPrograma = idPrograma,
                        IdAsignatura = idAsignatura,
                        Semestre = semestre,
                        VecesAsignada = veces,
                        FechaUltimaAsignacion = fechaIso,
                        Subunidad = subunidad
                    }, transaction: transaction);
                }
            }
        }

        private static string NormalizarFechaIso(string rawFecha)
        {
            if (string.IsNullOrWhiteSpace(rawFecha)) return string.Empty;
            string[] formats = { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy/MM/dd" };
            if (DateTime.TryParseExact(rawFecha.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return rawFecha.Trim();
        }
    }
}
