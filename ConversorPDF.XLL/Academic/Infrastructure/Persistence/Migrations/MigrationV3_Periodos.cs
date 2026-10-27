using System;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV3_Periodos : IMigration
    {
        public int Version => 3;
        public string Name => "Períodos Académicos e Integración de Período en Vinculaciones";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            string sql = @"
                CREATE TABLE IF NOT EXISTS periodos (
                    idperiodo INTEGER PRIMARY KEY,
                    periodo TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1
                );

                INSERT OR IGNORE INTO periodos (idperiodo, periodo, activo) VALUES 
                (1, '2026-01', 1),
                (2, '2026-02', 0),
                (3, '2027-01', 1),
                (4, '2027-02', 0),
                (5, '2028-01', 1),
                (6, '2028-02', 0);";

            connection.Execute(sql, transaction: transaction);

            try 
            {
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN idperiodo INTEGER DEFAULT 1;", transaction: transaction);
                connection.Execute("UPDATE vinculaciones SET idperiodo = 1 WHERE idperiodo IS NULL;", transaction: transaction);
            }
            catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0 || ex.Message.IndexOf("no such table", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Ignorar si la columna ya existe o la tabla no ha sido creada aún
            }

            try
            {
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
                        pe.periodo AS PeriodoNombre,
                        pe.idperiodo AS IdPeriodo,
                        v.fecha_inicio AS SesionFechaInicial,
                        v.fecha_fin AS SesionFechaFinal,
                        v.total_horas AS SesionHoras,
                        0 AS SesionViatica 
                    FROM vinculaciones v
                    JOIN programas p ON v.idprograma = p.id
                    JOIN terceros t ON v.idtercero = t.id
                    JOIN asignaturas a ON v.idasignatura = a.id
                    JOIN sedes s ON v.idsede = s.id
                    LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;";

                connection.Execute(viewSql, transaction: transaction);
            }
            catch (Exception ex) when (ex.Message.IndexOf("no such table", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Ignorar si las tablas vinculadas no existen aún
            }
        }
    }
}
