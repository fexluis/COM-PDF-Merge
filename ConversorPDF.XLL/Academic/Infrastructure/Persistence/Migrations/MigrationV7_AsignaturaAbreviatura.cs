using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV7_AsignaturaAbreviatura : IMigration
    {
        public int Version => 7;
        public string Name => "Abreviatura de Asignatura en Vista de Sesiones";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            string viewSql = @"
                DROP VIEW IF EXISTS vw_SesionesDetalle;
                CREATE VIEW vw_SesionesDetalle AS
                SELECT 
                    s.id AS IdSesion,
                    v.id AS IdVinculacion,
                    v.idperiodo AS IdPeriodo,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    v.semestre AS Semestre,
                    p.tipo AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    t.nombre AS DocenteNombre,
                    t.cedula AS CedulaTercero,
                    a.nombreasignatura AS AsignaturaNombre,
                    a.abreviatura AS AbreviaturaAsignatura,
                    s2.ciudad AS SedeCiudad,
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
                JOIN programas p ON v.idprograma = p.id
                JOIN terceros t ON v.idtercero = t.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s2 ON v.idsede = s2.id;";

            connection.Execute(viewSql, transaction: transaction);
        }
    }
}
