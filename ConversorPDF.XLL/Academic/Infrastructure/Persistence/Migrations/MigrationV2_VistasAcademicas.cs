using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV2_VistasAcademicas : IMigration
    {
        public int Version => 2;
        public string Name => "Vistas Iniciales Académicas";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            string viewSql = @"
                DROP VIEW IF EXISTS vw_VinculacionesDetalle;
                CREATE VIEW vw_VinculacionesDetalle AS
                SELECT 
                    v.id AS IdVinculacion,
                    p.tipo AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    t.nombre AS TerceroNombre,
                    v.semestre AS AsignaturaSemestre,
                    a.nombreasignatura AS AsignaturaNombre,
                    s.ciudad AS SedeCiudad,
                    s.id AS IdSede,
                    t.id AS IdTercero,
                    v.fecha_inicio AS SesionFechaInicial,
                    v.fecha_fin AS SesionFechaFinal,
                    v.total_horas AS SesionHoras,
                    0 AS SesionViatica 
                FROM vinculaciones v
                JOIN programas p ON v.idprograma = p.id
                JOIN terceros t ON v.idtercero = t.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s ON v.idsede = s.id;";

            connection.Execute(viewSql, transaction: transaction);
        }
    }
}
