using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AgProgramaCatalogoSqlTests
    {
        private static List<Programa> Consultar()
        {
            using (var cn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                cn.Open();
                cn.Execute(@"
                    CREATE TABLE programas (
                        id INTEGER PRIMARY KEY, nombreprograma TEXT, abreviatura TEXT,
                        pensum INTEGER, actual INTEGER, activo INTEGER, modalidad TEXT);
                    INSERT INTO programas (id, nombreprograma, abreviatura, pensum, actual, activo, modalidad) VALUES
                        (1,  'ADMINISTRACION PUBLICA TERRITORIAL', 'APT-53',      1, 1, 1, 'PREGRADO'),
                        (2,  'ADMINISTRACION PUBLICA TERRITORIAL', 'APT-52',      1, 1, 1, 'PREGRADO'),
                        (9,  'ESPECIALIZACION EN ESTADO Y EMPLEO', 'BEE-505',     1, 1, 0, 'POSGRADO'),
                        (21, 'MAESTRIA EN DERECHOS HUMANOS',       'MDH-409',     1, 1, 0, 'POSGRADO'),
                        (99, 'PROGRAMA SIN ACTIVO',                'SIN-ACT',     1, 1, NULL, 'POSGRADO');");

                return cn.Query<Programa>(AgProgramaCatalogoSql.SeleccionProgramas).ToList();
            }
        }

        [Fact]
        public void SeleccionProgramas_WhenHayInactivos_ShouldTraerTodoElCatalogo()
        {
            // Act
            var filas = Consultar();

            // Assert (Fase 1466)
            filas.Should().HaveCount(5);
            filas.Should().Contain(p => p.Activo == 0, "los programas inactivos también alimentan el combo AG");
            filas.Should().Contain(p => p.Modalidad == "POSGRADO");
        }

        [Fact]
        public void SeleccionProgramas_WhenSeFiltraPosgrado_ShouldDevolverPosgradosInactivos()
        {
            // Act
            var posgrados = AgTipoProgramaCatalog.Filtrar(Consultar(), AgTipoProgramaCatalog.Posgrado);

            // Assert
            posgrados.Should().HaveCount(3);
            posgrados.Should().Contain(p => p.Id == 9);
            posgrados.Should().Contain(p => p.Id == 21);
        }

        [Fact]
        public void SeleccionProgramas_ShouldOrdenarActivosPrimeroYLuegoPorEtiqueta()
        {
            // Act
            var ids = Consultar().Select(p => p.Id).ToList();

            // Assert
            ids.Should().ContainInOrder(2, 1, 9, 21, 99);
        }

        [Fact]
        public void SeleccionProgramas_WhenPensumOActualSonNull_ShouldMapearACeroSinLanzarExcepcion()
        {
            // Arrange (Fase 1467: blindaje ante pensum/actual NULL en otros vaults)
            using (var cn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                cn.Open();
                cn.Execute(@"
                    CREATE TABLE programas (
                        id INTEGER PRIMARY KEY, nombreprograma TEXT, abreviatura TEXT,
                        pensum INTEGER, actual INTEGER, activo INTEGER, modalidad TEXT);
                    INSERT INTO programas (id, nombreprograma, abreviatura, pensum, actual, activo, modalidad) VALUES
                        (100, 'PROGRAMA CON NULOS', 'NUL-001', NULL, NULL, 0, 'POSGRADO');");

                // Act
                var programas = cn.Query<Programa>(AgProgramaCatalogoSql.SeleccionProgramas).ToList();

                // Assert
                programas.Should().HaveCount(1);
                var p = programas.First();
                p.Id.Should().Be(100);
                p.Pensum.Should().Be(0);
                p.Actual.Should().Be(0);
                p.Activo.Should().Be(0);
            }
        }
    }
}
