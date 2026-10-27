using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V60: Crea índices de optimización para consultas y purga de staging por corrida y período (Fase 498).
    /// </summary>
    public class MigrationV60_AddAgStagingIndices : IMigration, IMigrationTolerante
    {
        public int Version => 60;
        public string Name => "AddAgStagingIndices";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V60] Creando índices de optimización en ag_vinculaciones y ag_sesiones...");

            string ddl = @"
CREATE INDEX IF NOT EXISTS idx_ag_vinc_corrida ON ag_vinculaciones(id_corrida, idperiodo, idprograma);
CREATE INDEX IF NOT EXISTS idx_ag_ses_vinc ON ag_sesiones(idvinculacion);
CREATE INDEX IF NOT EXISTS idx_ag_gen_ind ON ag_gen(id_individuo);
";
            connection.Execute(ddl, transaction: transaction);
            LoggerHelper.Info("[V60] Índices de optimización de staging creados exitosamente.");
        }
    }
}
