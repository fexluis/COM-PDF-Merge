using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV12_ExpandSesionesView : IMigration
    {
        public int Version => 12;
        public string Name => "Expandir vw_VinculacionesDetalle y vw_SesionesDetalle con columnas administrativas para Excel";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // 1. Asegurar que las columnas existan en la tabla 'sedes' (robustez de migración incremental)
            var sedesCols = connection.Query("PRAGMA table_info(sedes);", transaction: transaction)
                                     .Select(c => (string)c.name).ToList();
            if (!sedesCols.Contains("departamento", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE sedes ADD COLUMN departamento TEXT;", transaction: transaction);
            }
            if (!sedesCols.Contains("categoria", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE sedes ADD COLUMN categoria TEXT;", transaction: transaction);
            }
            if (!sedesCols.Contains("transporte", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE sedes ADD COLUMN transporte TEXT;", transaction: transaction);
            }

            // 2. Asegurar que las columnas existan en la tabla 'asignaturas'
            var asignaturasCols = connection.Query("PRAGMA table_info(asignaturas);", transaction: transaction)
                                           .Select(c => (string)c.name).ToList();
            if (!asignaturasCols.Contains("creditos", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN creditos INTEGER;", transaction: transaction);
            }
            if (!asignaturasCols.Contains("cod_materia", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN cod_materia TEXT;", transaction: transaction);
            }

            // 3. Eliminar y recrear vistas
            connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: transaction);
            connection.Execute(@"
                CREATE VIEW vw_VinculacionesDetalle AS
                SELECT 
                    v.id AS IdVinculacion,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    v.idprograma AS IdPrograma,
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
                    ) THEN 1 ELSE 0 END) AS SesionViatica,
                    s.departamento AS SedeDepartamento,
                    p.nombreprograma AS ProgramaNombre,
                    a.creditos AS AsignaturaCreditos,
                    a.metodologia AS AsignaturaMetodologia,
                    a.cod_materia AS AsignaturaCodMateria,
                    v.grupo AS Grupo,
                    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
                    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin
                FROM vinculaciones v
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s ON v.idsede = s.id
                LEFT JOIN terceros t ON v.idtercero = t.id
                LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;
            ", transaction: transaction);

            connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction);
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
                    CASE strftime('%w', NULLIF(NULLIF(s.fechasesionInicial, ''), '--'))
                        WHEN '0' THEN 'DOMINGO'
                        WHEN '1' THEN 'LUNES'
                        WHEN '2' THEN 'MARTES'
                        WHEN '3' THEN 'MIÉRCOLES'
                        WHEN '4' THEN 'JUEVES'
                        WHEN '5' THEN 'VIERNES'
                        WHEN '6' THEN 'SÁBADO'
                    END AS DiaSemana,
                    s.horas AS Horas,
                    s.duracion AS Duracion,
                    NULLIF(NULLIF(s.fechasesionInicial, ''), '--') AS FechaInicial,
                    NULLIF(NULLIF(s.fechasesionFinal, ''), '--') AS FechaFinal,
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
        }
    }
}
