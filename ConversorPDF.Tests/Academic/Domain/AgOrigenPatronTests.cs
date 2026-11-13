using System.Data;
using System.Data.SQLite;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgOrigenPatronTests
    {
        [Fact]
        public void AgAulaAsignatura_DefaultOrigenPatron_ShouldBeLegado()
        {
            var aa = new AgAulaAsignatura();
            aa.OrigenPatron.Should().Be("LEGADO");
        }

        [Fact]
        public void AgOrigenPatron_ValoresEnum_DebenSerConsistentes()
        {
            ((int)AgOrigenPatron.Automatico).Should().Be(0);
            ((int)AgOrigenPatron.Fijo).Should().Be(1);
            ((int)AgOrigenPatron.Legado).Should().Be(2);
        }

        [Fact]
        public void MigrationV80_WhenTableExists_ShouldAddColumnAndNormalize()
        {
            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;"))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE ag_aula_asignatura (
                        id_aula INTEGER,
                        id_asignatura INTEGER,
                        dias_patron TEXT
                    );
                    INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, dias_patron)
                    VALUES (1, 101, '5,6');");

                var mig = new MigrationV80_AddOrigenPatronAgAulaAsignatura();
                mig.Version.Should().Be(80);
                mig.Up(conn);

                // Verificar que la columna fue añadida y el valor por defecto es LEGADO
                string origen = conn.QueryFirstOrDefault<string>(
                    "SELECT origen_patron FROM ag_aula_asignatura WHERE id_aula = 1;");

                origen.Should().Be("LEGADO");

                // Idempotencia: ejecutar nuevamente no debe fallar
                mig.Up(conn);
            }
        }
    }
}

