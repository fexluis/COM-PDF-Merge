using System;
using System.Data.SQLite;
using System.Reflection;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using Dapper;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public sealed class AgConfirmadorPlanAprobadoTests
    {
        [Theory]
        [InlineData(1, "AgHorariosAcademicos", true)]
        [InlineData(0, "AgHorariosAcademicos", false)]
        [InlineData(1, "XAgHorariosAcademicos", false)]
        [InlineData(1, "AgHorariosAcademicosX", false)]
        [InlineData(1, "X, AgHorariosAcademicos ,Y", true)]
        [InlineData(1, "X,AgHorariosAcademicosX,Y", false)]
        public void SoloUsuarioActivoYTokenCsvCompleto(int activo, string csv, bool esperado)
        {
            using (var db = new SQLiteConnection("Data Source=:memory:;Pooling=False"))
            {
                db.Open();
                db.Execute(@"CREATE TABLE usuarios(nombreusuario TEXT, perfil TEXT, activo INTEGER);
CREATE TABLE permisos(perfil TEXT, opcionespermitidas TEXT);
INSERT INTO usuarios VALUES ('actor','perfil',@Activo);
INSERT INTO permisos VALUES ('perfil',@Csv);", new { Activo = activo, Csv = csv });
                using (var tx = db.BeginTransaction())
                {
                    var metodo = typeof(AgConfirmadorPlanAprobado).GetMethod("PermisoVigente", BindingFlags.NonPublic | BindingFlags.Static);
                    Assert.NotNull(metodo);
                    Assert.Equal(esperado, (bool)metodo.Invoke(null, new object[] { db, tx, "actor" }));
                    Assert.False((bool)metodo.Invoke(null, new object[] { db, tx, "otro" }));
                    Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM usuarios", transaction: tx));
                    tx.Rollback();
                }
            }
        }
    }
}
