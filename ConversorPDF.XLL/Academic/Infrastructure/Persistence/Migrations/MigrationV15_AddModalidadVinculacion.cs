using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV15_AddModalidadVinculacion : IMigration
    {
        public int Version => 15;
        public string Name => "Agregar columna modalidad a vinculaciones";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Fase 63 FIX: Verificar que 'vinculaciones' existe antes del ALTER.
            // hasModalidad=0 ocurre tanto si la tabla no existe como si falta la columna.
            // Sin este guard, el ALTER crashea con "no such table" en vaults parciales.
            int vincExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';",
                transaction: transaction);

            if (vincExists > 0)
            {
                int hasModalidad = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones' AND sql LIKE '%modalidadVinculacion%';",
                    transaction: transaction);

                if (hasModalidad == 0)
                {
                    connection.Execute(
                        "ALTER TABLE vinculaciones ADD COLUMN modalidadVinculacion TEXT;",
                        transaction: transaction
                    );
                }
            }

            // Fix P0/Fase 7 y 50: Validar que TODAS las tablas base de los JOINs existan antes de recrear vistas.
            string[] requiredTables = { "vinculaciones", "sesiones", "programas", "asignaturas", "sedes", "terceros", "periodos", "terminales" };
            foreach (var table in requiredTables)
            {
                int exists = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name;",
                    new { name = table }, transaction: transaction);
                
                if (exists == 0) 
                {
                    return; // tabla faltante, no se pueden recrear las vistas aún
                }
            }

            // Eliminar vistas dependientes
            try { connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: transaction); } catch { }
            try { connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction); } catch { }

            // Fase 50 FIX: Recrear vw_SesionesDetalle sincronizada con el schema canónico
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
                    s2.subsede AS Subsede,
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
                    s2.transporte AS TransporteSede,
                    p.nombreprograma AS ProgramaNombre,
                    a.creditos AS AsignaturaCreditos,
                    a.metodologia AS AsignaturaMetodologia,
                    a.cod_materia AS AsignaturaCodMateria,
                    a.num_nivel AS Nivel,
                    v.grupo AS Grupo,
                    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
                    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin,
                    v.modalidadVinculacion AS Modalidad,
                    v.isTerminalAereo AS IsTerminalAereo,
                    v.idTerminales AS IdTerminales,
                    term.terminal AS TerminalAereaNombre,
                    v.subunidad AS Subunidad
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s2 ON v.idsede = s2.id
                LEFT JOIN terceros t ON s.idtercero = t.id
                LEFT JOIN terminales term ON v.idTerminales = term.id;
            ", transaction: transaction);

            // Recrear vw_VinculacionesDetalle
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
                    ) THEN 1 ELSE 0 END) AS SesionViatica,
                    v.modalidadVinculacion AS Modalidad,
                    v.subunidad AS Subunidad
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
