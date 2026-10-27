using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V56: Ajusta la fecha de inicio del período 2026-02 al 2026-08-11 para coincidir con la primera sesión real del semestre (Fase 386 / Fase 406).
    /// </summary>
    public class MigrationV56_AjustarInicioPeriodo2026_02 : IMigration, IMigrationTolerante
    {
        public int Version => 56;
        public string Name => "AjustarInicioPeriodo2026_02";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V56] Ajustando periodo_inicio para 2026-02 a 2026-08-11 (H14)...");

            int tablaExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';",
                transaction: transaction);

            if (tablaExiste == 0) return;

            connection.Execute(@"
                UPDATE periodos
                SET periodo_inicio = '2026-08-11'
                WHERE (idperiodo = 2 OR periodo = '2026-02')
                  AND (periodo_inicio = '2026-08-15' OR periodo_inicio IS NULL OR periodo_inicio = '');",
                transaction: transaction);

            LoggerHelper.Info("[V56] Fecha de inicio ajustada exitosamente para 2026-02.");
        }
    }
}
