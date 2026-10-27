using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración v14: Actualiza el schema de la tabla sesiones.
    /// - Renombra fechasesionInicial → fechasesion.
    /// - Elimina fechasesionFinal y duracion.
    /// - Añade horariosesion.
    /// - Recrea vw_SesionesDetalle con la nueva estructura.
    /// </summary>
    public class MigrationV14_UpdateSesionesSchema : IMigration
    {
        public int Version => 14;
        public string Name => "Actualizar schema sesiones: fechasesion + horariosesion, eliminar fechasesionFinal y duracion";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // 0. Eliminar vistas dependientes primero para evitar errores de validación de SQLite al tocar tablas base
            try { connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: transaction); } catch { }
            try { connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction); } catch { }

            // 1. Verificar columnas actuales de sesiones
            var sesionCols = connection.Query("PRAGMA table_info(sesiones);", transaction: transaction)
                                       .Select(c => (string)c.name).ToList();

            bool tieneInicial = sesionCols.Contains("fechasesionInicial", System.StringComparer.OrdinalIgnoreCase);
            bool tieneFechaSesion = sesionCols.Contains("fechasesion", System.StringComparer.OrdinalIgnoreCase);
            bool tieneHorario = sesionCols.Contains("horariosesion", System.StringComparer.OrdinalIgnoreCase);

            // SQLite no soporta DROP COLUMN ni RENAME COLUMN en versiones antiguas.
            // Usamos el patrón de recreación de tabla.
            if (tieneInicial || !tieneFechaSesion || !tieneHorario)
            {
                // 1a. Tabla temporal con nuevo schema
                connection.Execute(@"
                    CREATE TABLE IF NOT EXISTS sesiones_v14 (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER,
                        idtercero INTEGER,
                        fechasesion TEXT,
                        horariosesion TEXT,
                        horas REAL,
                        viatica INTEGER,
                        FOREIGN KEY(idvinculacion) REFERENCES vinculaciones(id),
                        FOREIGN KEY(idtercero) REFERENCES terceros(id)
                    );
                ", transaction: transaction);

                // 1b. Copiar datos existentes: fechasesionInicial → fechasesion
                if (tieneInicial)
                {
                    connection.Execute(@"
                        INSERT INTO sesiones_v14 (id, idvinculacion, idtercero, fechasesion, horariosesion, horas, viatica)
                        SELECT id, idvinculacion, idtercero, fechasesionInicial, NULL, horas, viatica
                        FROM sesiones;
                    ", transaction: transaction);
                }
                else if (tieneFechaSesion)
                {
                    connection.Execute(@"
                        INSERT INTO sesiones_v14 (id, idvinculacion, idtercero, fechasesion, horariosesion, horas, viatica)
                        SELECT id, idvinculacion, idtercero, fechasesion,
                               CASE WHEN horariosesion IS NULL THEN NULL ELSE horariosesion END,
                               horas, viatica
                        FROM sesiones;
                    ", transaction: transaction);
                }

                // 1c. Reemplazar tabla (las vistas dependientes ya se eliminaron al inicio)
                connection.Execute("DROP TABLE IF EXISTS sesiones;", transaction: transaction);
                connection.Execute("ALTER TABLE sesiones_v14 RENAME TO sesiones;", transaction: transaction);
            }

            // 2. Recrear vw_SesionesDetalle con la nueva estructura
            connection.Execute(@"
                CREATE VIEW vw_SesionesDetalle AS
                SELECT 
                    s.id AS IdSesion,
                    s.idvinculacion AS IdVinculacion,
                    v.idperiodo AS IdPeriodo,
                    v.idsede AS IdSede,
                    v.idprograma AS IdPrograma,
                    s.idtercero AS IdTercero,
                    p.modalidad AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    t.nombre AS DocenteNombre,
                    t.cedula AS CedulaTercero,
                    a.nombre AS AsignaturaNombre,
                    a.abreviatura AS AsignaturaAbreviatura,
                    s2.ciudad AS SedeCiudad,
                    CASE strftime('%w', NULLIF(s.fechasesion, ''))
                        WHEN '0' THEN 'DOMINGO'
                        WHEN '1' THEN 'LUNES'
                        WHEN '2' THEN 'MARTES'
                        WHEN '3' THEN 'MIÉRCOLES'
                        WHEN '4' THEN 'JUEVES'
                        WHEN '5' THEN 'VIERNES'
                        WHEN '6' THEN 'SÁBADO'
                    END AS DiaSemana,
                    s.horas AS Horas,
                    s.horariosesion AS HorarioSesion,
                    NULLIF(s.fechasesion, '') AS FechaSesion,
                    s.viatica AS Viatica,
                    s2.departamento AS SedeDepartamento,
                    p.nombreprograma AS ProgramaNombre,
                    a.creditos AS AsignaturaCreditos,
                    a.metodologia AS AsignaturaMetodologia,
                    a.cod_materia AS AsignaturaCodMateria,
                    v.grupo AS Grupo,
                    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
                    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s2 ON v.idsede = s2.id
                LEFT JOIN terceros t ON s.idtercero = t.id;
            ", transaction: transaction);

            // 3. Recrear vw_VinculacionesDetalle
            connection.Execute(@"
                CREATE VIEW vw_VinculacionesDetalle AS
                SELECT 
                    v.id AS IdVinculacion,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    pe.idperiodo AS IdPeriodo,
                    pe.periodo AS PeriodoNombre,
                    p.modalidad AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    t.nombre AS TerceroNombre,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    a.nombre AS AsignaturaNombre,
                    s.ciudad AS SedeCiudad,
                    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS SesionFechaInicial,
                    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS SesionFechaFinal,
                    v.total_horas AS SesionHoras,
                    (SELECT CASE WHEN EXISTS (
                        SELECT 1 FROM sesiones ses WHERE ses.idvinculacion = v.id AND ses.viatica = 1
                    ) THEN 1 ELSE 0 END) AS SesionViatica
                FROM vinculaciones v
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s ON v.idsede = s.id
                LEFT JOIN terceros t ON v.idtercero = t.id
                LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;
            ", transaction: transaction);
        }
    }
}
