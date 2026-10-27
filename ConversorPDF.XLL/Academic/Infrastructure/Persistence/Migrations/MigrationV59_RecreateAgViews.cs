using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V59: Recrea las vistas de staging vw_AgSesionesDetalle y vw_AgVinculacionesDetalle con el esquema completo
    /// de V54/V55/V58 (vacantes, pool, id_corrida, etc.) y asegura la existencia de ag_salvedades_conflictos (Fase 468).
    /// </summary>
    public class MigrationV59_RecreateAgViews : IMigration, IMigrationTolerante
    {
        public int Version => 59;
        public string Name => "RecreateAgViews";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V59] Recreando vistas de staging vw_AgSesionesDetalle y vw_AgVinculacionesDetalle...");

            int agSesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';", transaction: transaction);
            if (agSesExists == 0) return;

            // FASE 855: Delegar creación de vistas en SqliteHelper.EnsureAgViews para evitar DDL divergente
            SqliteHelper.EnsureAgViews(connection, transaction);

            string ddl = @"
-- Tabla referenciada por AgModoDatos.TablaSalvedades
CREATE TABLE IF NOT EXISTS ag_salvedades_conflictos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idsesion_a INTEGER NOT NULL DEFAULT 0,
    idsesion_b INTEGER NOT NULL DEFAULT 0,
    motivo TEXT DEFAULT '', subunidad TEXT DEFAULT '',
    UNIQUE(idsesion_a, idsesion_b)
);";

            connection.Execute(ddl, transaction: transaction);
            LoggerHelper.Info("[V59] Vistas de staging y tabla ag_salvedades_conflictos aseguradas exitosamente.");
        }
    }
}
