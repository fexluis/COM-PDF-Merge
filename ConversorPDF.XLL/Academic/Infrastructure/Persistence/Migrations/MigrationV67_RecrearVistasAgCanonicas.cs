using System;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// V67: Recrea vw_AgSesionesDetalle / vw_AgVinculacionesDetalle con el DDL canónico
    /// y validación simétrica de columnas requeridas (Fases 854/861/864/871/875).
    /// Responsabilidad única: DDL y validación estructural de vistas sin mutar datos de sesiones (delegado a V68).
    /// </summary>
    public sealed class MigrationV67_RecrearVistasAgCanonicas : IMigration, IMigrationTolerante
    {
        public int Version => 67;
        public string Name => "RecrearVistasAgCanonicas";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            int agSesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';", transaction: transaction);
            if (agSesExists == 0) return;

            SqliteHelper.EnsureAgViews(connection, transaction);

            // Fase 864: SQLite no valida columnas al crear vistas -> forzar resolución real
            connection.Execute("SELECT * FROM vw_AgSesionesDetalle LIMIT 0;", transaction: transaction);
            connection.Execute("SELECT * FROM vw_AgVinculacionesDetalle LIMIT 0;", transaction: transaction);

            // Validar columnas requeridas de vw_AgSesionesDetalle
            var requeridas = new[] { "DiaSemana", "IdAula", "IdGen", "EsVacante",
                                     "MotivoVacante", "CandidatosPool", "DiaHorasSemanales" };
            var cols = connection.Query<string>(
                "SELECT name FROM pragma_table_info('vw_AgSesionesDetalle');", transaction: transaction).ToList();
            var faltan = requeridas.Except(cols).ToList();

            if (faltan.Count > 0)
            {
                throw new InvalidOperationException(
                    "V67: faltan columnas en vw_AgSesionesDetalle -> " + string.Join(",", faltan));
            }

            // Fase 871: Validación simétrica de columnas requeridas de vw_AgVinculacionesDetalle
            var reqVinc = new[] { "IdVinculacion", "IdCorrida", "IdIndividuo", "IdPrograma", "IdAsignatura",
                                  "Grupo", "Estado", "DiaHorasSemanales", "TotalHoras",
                                  "EsVacante", "MotivoVacante", "CandidatosPool" };
            var colsVinc = connection.Query<string>(
                "SELECT name FROM pragma_table_info('vw_AgVinculacionesDetalle');", transaction: transaction).ToList();
            var faltanVinc = reqVinc.Except(colsVinc).ToList();

            if (faltanVinc.Count > 0)
            {
                throw new InvalidOperationException(
                    "V67: faltan columnas en vw_AgVinculacionesDetalle -> " + string.Join(",", faltanVinc));
            }
        }
    }
}
