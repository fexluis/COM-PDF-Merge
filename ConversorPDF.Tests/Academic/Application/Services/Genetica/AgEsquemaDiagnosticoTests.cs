using System;
using System.Collections.Generic;
using System.Data.SQLite;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgEsquemaDiagnosticoTests
    {
        [Fact]
        public void Verificar_WhenAllRequiredColumnsExist_ShouldReportOk()
        {
            // Arrange
            var lineas = new List<string>();
            var traza = new AgDiagnosticoTraza("PruebaEsquemaOk", msg => lineas.Add(msg), "ESQ-01") { Habilitado = true };

            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                conn.Open();
                conn.Execute("CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, idprograma INT, semestre TEXT, subunidad TEXT);");

                // Act
                var faltantes = AgEsquemaDiagnostico.Verificar(conn, traza, "asignaturas", "id", "idprograma", "semestre");

                // Assert
                faltantes.Should().BeEmpty();
                lineas.Should().Contain(l => l.Contains("OK") && l.Contains("columnas"));
            }
        }

        [Fact]
        public void Verificar_WhenRequiredColumnsAreMissing_ShouldReportMissingColumns()
        {
            // Arrange
            var lineas = new List<string>();
            var traza = new AgDiagnosticoTraza("PruebaEsquemaFaltante", msg => lineas.Add(msg), "ESQ-02") { Habilitado = true };

            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                conn.Open();
                conn.Execute("CREATE TABLE ag_aula (id INTEGER PRIMARY KEY, codigo TEXT, nombre TEXT);");

                // Act
                var faltantes = AgEsquemaDiagnostico.Verificar(conn, traza, "ag_aula", "id", "aula_ref", "subunidad");

                // Assert
                faltantes.Should().BeEquivalentTo(new[] { "aula_ref", "subunidad" });
                lineas.Should().Contain(l => l.Contains("FALTAN COLUMNAS") && l.Contains("aula_ref"));
            }
        }

        [Fact]
        public void Verificar_WhenTableDoesNotExist_ShouldReportMissingAllColumns()
        {
            // Arrange
            var lineas = new List<string>();
            var traza = new AgDiagnosticoTraza("PruebaTablaInexistente", msg => lineas.Add(msg), "ESQ-03") { Habilitado = true };

            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                conn.Open();

                // Act
                var faltantes = AgEsquemaDiagnostico.Verificar(conn, traza, "tabla_fantasma", "id", "nombre");

                // Assert
                faltantes.Should().BeEquivalentTo(new[] { "id", "nombre" });
                lineas.Should().Contain(l => l.Contains("TABLA INEXISTENTE"));
            }
        }
    }
}
