using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V58: Asegura la existencia y activación del tercero centinela VACANTE (cédula 1111111111) (Fase 465).
    /// FASE 819: Protegido contra triggers CRDT obsoletos en la tabla 'terceros' con auto-recuperación y reintento.
    /// </summary>
    public class MigrationV58_EnsureTerceroVacante : IMigration, IMigrationTolerante
    {
        private const string InsertVacanteSql =
            @"INSERT INTO terceros (nombre, cedula, residencia, correo, telefono, activo, subunidad, tipo, categoria)
              VALUES ('VACANTE', '1111111111', '', 'VACANTE@esap.edu.co', '0000', 1, '', 'CATEDRA', 'AUXILIAR');";

        public int Version => 58;
        public string Name => "EnsureTerceroVacante";
        public bool Aditiva => true;

        private static bool EsFalloDeTriggerCrdt(Exception ex)
        {
            string m = ex?.Message ?? string.Empty;
            return m.IndexOf("values", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("trigger", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("no such column", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("no column named", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            int existe = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='terceros';", transaction: transaction);
            if (existe == 0)
            {
                return;
            }

            int ya = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM terceros WHERE cedula = '1111111111' OR UPPER(TRIM(nombre)) = 'VACANTE';", transaction: transaction);
            if (ya > 0)
            {
                const string updateSql = "UPDATE terceros SET activo = 1 WHERE cedula = '1111111111' OR UPPER(TRIM(nombre)) = 'VACANTE';";
                try
                {
                    connection.Execute(updateSql, transaction: transaction);
                }
                catch (Exception ex) when (EsFalloDeTriggerCrdt(ex))
                {
                    LoggerHelper.Warning("[V58] Trigger CRDT obsoleto en 'terceros' durante UPDATE; soltando triggers y reintentando.", ex);
                    CrdtTriggerRepair.DropStaleTriggers(connection, transaction, "terceros");
                    connection.Execute(updateSql, transaction: transaction);
                }

                LoggerHelper.Info("[V58] Tercero centinela VACANTE (cédula 1111111111) existente reactivado.");
                return;
            }

            try
            {
                connection.Execute(InsertVacanteSql, transaction: transaction);
            }
            catch (Exception ex) when (EsFalloDeTriggerCrdt(ex))
            {
                LoggerHelper.Warning("[V58] Trigger CRDT obsoleto en 'terceros'; soltando triggers y reintentando inserción.", ex);
                CrdtTriggerRepair.DropStaleTriggers(connection, transaction, "terceros");
                connection.Execute(InsertVacanteSql, transaction: transaction);
            }

            LoggerHelper.Info("[V58] Tercero centinela VACANTE (cédula 1111111111) asegurado e insertado.");
        }
    }
}
