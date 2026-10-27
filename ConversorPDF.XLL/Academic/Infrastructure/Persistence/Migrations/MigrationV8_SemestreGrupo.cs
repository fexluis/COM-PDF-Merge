using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV8_SemestreGrupo : IMigration
    {
        public int Version => 8;
        public string Name => "Combinar Semestre y Grupo en Vistas";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // 1. Asegurar que la columna 'grupo' existe en 'vinculaciones' (Resiliencia para instalaciones existentes)
            var columns = connection.Query("PRAGMA table_info(vinculaciones)", transaction: transaction)
                                    .Select(c => (string)c.name)
                                    .ToList();
            
            if (!columns.Contains("grupo", System.StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN grupo TEXT;", transaction: transaction);
            }

            // 2. Redefinir vw_VinculacionesDetalle con IdTercero y Semestre combinado
            connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: transaction);
            connection.Execute(@"
                CREATE VIEW vw_VinculacionesDetalle AS
                SELECT 
                    v.id AS IdVinculacion,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    v.grupo AS Grupo,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    p.nombreprograma AS ProgramaNombre,
                    p.abreviatura AS ProgramaAbreviatura,
                    a.nombreasignatura AS AsignaturaNombre,
                    a.abreviatura AS AsignaturaAbreviatura,
                    s.ciudad AS SedeCiudad,
                    t.nombre AS DocenteNombre,
                    t.cedula AS CedulaTercero,
                    v.total_horas AS TotalHoras,
                    v.fecha_inicio AS FechaInicio,
                    v.fecha_fin AS FechaFin,
                    (SELECT CASE WHEN EXISTS (
                        SELECT 1 FROM sesiones ses WHERE ses.idvinculacion = v.id AND ses.viatica = 1
                    ) THEN 1 ELSE 0 END) AS SesionViatica,
                    pe.idperiodo AS IdPeriodo,
                    pe.periodo AS Periodo
                FROM vinculaciones v
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s ON v.idsede = s.id
                LEFT JOIN terceros t ON v.idtercero = t.id
                LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;", transaction: transaction);

            // 3. Redefinir vw_SesionesDetalle con Semestre combinado
            connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction);
            connection.Execute(@"
                CREATE VIEW vw_SesionesDetalle AS
                SELECT 
                    s.id AS IdSesion,
                    s.idvinculacion AS IdVinculacion,
                    v.grupo AS Grupo,
                    v.idperiodo AS IdPeriodo,
                    v.idsede AS IdSede,
                    s2.ciudad AS SedeCiudad,
                    s.idtercero AS IdTercero,
                    t.nombre AS DocenteNombre,
                    t.cedula AS CedulaTercero,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    p.tipo AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    a.nombreasignatura AS AsignaturaNombre,
                    a.abreviatura AS AsignaturaAbreviatura,
                    CASE strftime('%w', s.fechasesionInicial)
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
                    s.fechasesionInicial AS FechaInicial,
                    s.fechasesionFinal AS FechaFinal,
                    s.viatica AS Viatica
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id
                JOIN sedes s2 ON v.idsede = s2.id
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                LEFT JOIN terceros t ON s.idtercero = t.id;", transaction: transaction);
        }
    }
}
