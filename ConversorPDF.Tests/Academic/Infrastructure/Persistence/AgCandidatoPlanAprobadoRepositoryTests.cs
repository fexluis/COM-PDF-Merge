using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ConversorPDF.Tests.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public sealed class AgCandidatoPlanAprobadoRepositoryTests
    {
        [Fact]
        public void GuardaAuditoriaYCandidatoJuntosYNoConsumeOtroPasoDelMismoIntento()
        {
            string ruta = Path.Combine(Path.GetTempPath(), "ag_ratificacion_" + Guid.NewGuid().ToString("N") + ".sqlite");
            try
            {
                using (var db = new SQLiteConnection("Data Source=" + ruta + ";Pooling=False"))
                {
                    db.Open();
                    new MigrationV82_AddAgCandidatoPlanAprobado().Up(db);
                    var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
                    int aula = m.Relaciones.First().Clave.IdAula;
                    var intento = Guid.NewGuid();
                    var c = new AgCandidatoPlanAprobado(Guid.NewGuid(), intento, aula, "actor", "motivo", m,
                        "origen", "anterior", "MAR-VIE", "LUN-VIE", new int[0]);
                    using (var tx = db.BeginTransaction())
                    {
                        Assert.True(new AgCandidatoPlanAprobadoRepository().Guardar(db, tx, c));
                        Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_candidato_plan_aprobado", transaction: tx));
                        Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_auditoria_ratificacion", transaction: tx));
                        tx.Rollback();
                    }
                    Assert.Equal(0, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_candidato_plan_aprobado"));
                    Assert.Equal(0, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_auditoria_ratificacion"));
                    using (var tx = db.BeginTransaction())
                    {
                        Assert.True(new AgCandidatoPlanAprobadoRepository().Guardar(db, tx, c));
                        tx.Commit();
                    }
                    using (var tx = db.BeginTransaction())
                    {
                        var otro = new AgCandidatoPlanAprobado(Guid.NewGuid(), intento, aula, "actor", "motivo", m,
                            "origen", "anterior", "LUN-VIE", "LUN-VIE", new int[0]);
                        Assert.Throws<SQLiteException>(() => new AgCandidatoPlanAprobadoRepository().Guardar(db, tx, otro));
                        tx.Rollback();
                    }
                    Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_candidato_plan_aprobado"));
                    Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_auditoria_ratificacion"));
                }
            }
            finally { SQLiteConnection.ClearAllPools(); if (File.Exists(ruta)) File.Delete(ruta); }
        }
    }
}
