using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// V69: Revalida y recrea las vistas canónicas del motor AG (vw_AgSesionesDetalle y vw_AgVinculacionesDetalle).
    /// Corrige el hueco generado si V67 falló previamente por la referencia histórica a v.idindividuo (Fases 892/893/895).
    /// </summary>
    public sealed class MigrationV69_RevalidarVistasAg : IMigration, IMigrationTolerante
    {
        public int Version => 69;
        public string Name => "RevalidarVistasAg";
        public bool Aditiva => true; // FASE 899: Tolerante, no bloquea el arranque en vaults legados con esquemas atípicos.

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            int agSesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';", transaction: transaction);
            if (agSesExists == 0) return;

            // Recrear vistas canónicas con columnas correctas y validación por vista
            SqliteHelper.EnsureAgViews(connection, transaction);

            // Validación estricta inmediata de resolución de columnas
            connection.Execute("SELECT * FROM vw_AgSesionesDetalle LIMIT 0;", transaction: transaction);

            bool hayPeriodos = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';",
                transaction: transaction) > 0;

            if (hayPeriodos)
            {
                connection.Execute("SELECT * FROM vw_AgVinculacionesDetalle LIMIT 0;", transaction: transaction);
            }
        }
    }
}
