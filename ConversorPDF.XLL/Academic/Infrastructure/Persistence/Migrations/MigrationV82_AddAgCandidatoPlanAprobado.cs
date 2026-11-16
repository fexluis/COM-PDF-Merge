using System.Data;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>V82: sólo tablas nuevas; no transforma ni borra sesiones, aulas o auditoría previa.</summary>
    public sealed class MigrationV82_AddAgCandidatoPlanAprobado : IMigration, IMigrationTolerante
    {
        public int Version => 82;
        public string Name => "AddAgCandidatoPlanAprobado";
        public bool Aditiva => true;

        /// <summary>Crea las identidades y el registro indivisible de ratificación; nunca ejecutado durante este cambio.</summary>
        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            connection.Execute(@"CREATE TABLE IF NOT EXISTS ag_candidato_plan_aprobado (
                operacion TEXT PRIMARY KEY NOT NULL,
                intento TEXT NOT NULL,
                id_aula INTEGER NOT NULL,
                actor TEXT NOT NULL,
                motivo TEXT NOT NULL,
                alcance TEXT NOT NULL,
                huella_manifiesto TEXT NOT NULL,
                huella_origen TEXT NOT NULL,
                huella_previa TEXT NOT NULL,
                huella_efecto TEXT NOT NULL,
                desde TEXT NOT NULL,
                hacia TEXT NOT NULL,
                consentimiento TEXT NOT NULL,
                inicio TEXT NOT NULL,
                fin TEXT NOT NULL,
                creado_utc TEXT NOT NULL,
                UNIQUE (intento, id_aula)
            );
            CREATE TABLE IF NOT EXISTS ag_auditoria_ratificacion (
                operacion TEXT PRIMARY KEY NOT NULL REFERENCES ag_candidato_plan_aprobado(operacion),
                actor TEXT NOT NULL,
                huella_origen TEXT NOT NULL,
                motivo TEXT NOT NULL,
                alcance TEXT NOT NULL,
                creado_utc TEXT NOT NULL
            );", transaction: transaction);
        }
    }
}
