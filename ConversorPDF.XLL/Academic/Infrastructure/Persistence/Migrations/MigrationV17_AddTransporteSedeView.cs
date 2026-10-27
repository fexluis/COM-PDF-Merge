using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV17_AddTransporteSedeView : IMigration
    {
        public int Version => 17;
        public string Name => "Agregar transporte a vw_SesionesDetalle";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // Eliminar vista dependiente
            try { connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction); } catch { }

            // Recrear vw_SesionesDetalle añadiendo s2.transporte AS TransporteSede
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
                    s2.transporte AS TransporteSede,
                    p.nombreprograma AS ProgramaNombre,
                    a.creditos AS AsignaturaCreditos,
                    a.metodologia AS AsignaturaMetodologia,
                    a.cod_materia AS AsignaturaCodMateria,
                    v.grupo AS Grupo,
                    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
                    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin,
                    v.modalidadVinculacion AS Modalidad
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
