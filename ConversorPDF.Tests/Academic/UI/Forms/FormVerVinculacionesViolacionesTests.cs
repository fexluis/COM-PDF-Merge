using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.UI.Forms;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConversorPDF.Tests.Academic.UI.Forms
{
    [TestClass]
    public class FormVerVinculacionesViolacionesTests
    {
        private sealed class STATestMethodAttribute : TestMethodAttribute
        {
            public override TestResult[] Execute(ITestMethod testMethod)
            {
                if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                {
                    return base.Execute(testMethod);
                }

                TestResult[] results = null;
                var thread = new Thread(() =>
                {
                    results = base.Execute(testMethod);
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                return results;
            }
        }

        private static SQLiteConnection CreateInMemoryDb()
        {
            var conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            conn.Open();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
CREATE TABLE ag_corrida (
    id INTEGER PRIMARY KEY,
    id_individuo_optimo INTEGER,
    subunidad TEXT
);
CREATE TABLE ag_individuo (
    id INTEGER PRIMARY KEY,
    id_corrida INTEGER
);
CREATE TABLE ag_violacion (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    id_individuo INTEGER NOT NULL,
    id_gen INTEGER NULL,
    regla INTEGER NOT NULL,
    es_dura INTEGER NOT NULL DEFAULT 1,
    mensaje TEXT NOT NULL DEFAULT '',
    detalle_json TEXT NOT NULL DEFAULT '{}',
    subunidad TEXT NOT NULL DEFAULT ''
);
CREATE TABLE ag_sesiones (
    id INTEGER PRIMARY KEY,
    idgen INTEGER
);";
                cmd.ExecuteNonQuery();
            }
            return conn;
        }

        [STATestMethod]
        public async Task CargarViolacionesAsync_WhenViolacionHasIdGenMatchingSession_ShouldLinkToViolacionesPorSesion()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, 'CALI');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, 101, 1, 1, 'Docente superpuesto en franja', '{}', 'CALI');";
                    cmd.ExecuteNonQuery();
                }

                var items = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 55,
                        IdGen = 101,
                        IdCorrida = 1,
                        DocenteNombre = "Carlos Lopez",
                        SedeCiudad = "CALI",
                        Horas = 4,
                        FechaSesion = new DateTime(2026, 11, 20),
                        Subunidad = "CALI"
                    }
                };

                using (var form = new FormVerVinculaciones(items, conn, null, AgModoDatos.Ag))
                {
                    // Act
                    await form.CargarViolacionesAsync();

                    // Assert
                    form.ViolacionesPorSesion.Should().ContainKey(55);
                    form.ViolacionesPorSesion[55].Should().HaveCount(1);
                    form.ViolacionesPorSesion[55][0].Regla.Should().Be("1");
                    form.ViolacionesPorSesion[55][0].EsDura.Should().BeTrue();
                    form.ViolacionesPorSesion[55][0].Mensaje.Should().Be("Docente superpuesto en franja");
                    form.ViolacionesGlobales.Should().BeEmpty();
                }
            }
        }

        [STATestMethod]
        public async Task CargarViolacionesAsync_WhenViolacionHasNoIdGen_ShouldAddToViolacionesGlobales()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, 'CALI');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, NULL, 4, 1, 'Supera límite máximo global de franjas', '{}', 'CALI');";
                    cmd.ExecuteNonQuery();
                }

                var items = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 55,
                        IdGen = 101,
                        IdCorrida = 1,
                        DocenteNombre = "Carlos Lopez",
                        SedeCiudad = "CALI",
                        Horas = 4,
                        FechaSesion = new DateTime(2026, 11, 20),
                        Subunidad = "CALI"
                    }
                };

                using (var form = new FormVerVinculaciones(items, conn, null, AgModoDatos.Ag))
                {
                    // Act
                    await form.CargarViolacionesAsync();

                    // Assert
                    form.ViolacionesPorSesion.Should().BeEmpty();
                    form.ViolacionesGlobales.Should().HaveCount(1);
                    form.ViolacionesGlobales[0].IdGen.Should().BeNull();
                    form.ViolacionesGlobales[0].Regla.Should().Be("4");
                    form.ViolacionesGlobales[0].EsDura.Should().BeTrue();
                    form.ViolacionesGlobales[0].Mensaje.Should().Be("Supera límite máximo global de franjas");
                }
            }
        }

        [STATestMethod]
        public async Task CargarViolacionesAsync_WhenViolacionHasIdGenWithoutMatchingSessionInItems_ShouldAddToViolacionesGlobales()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, '');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, 999, 2, 0, 'Aula sin correspondencia en items', '{}', '');";
                    cmd.ExecuteNonQuery();
                }

                var items = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 12,
                        IdGen = 50,
                        IdCorrida = 1,
                        DocenteNombre = "Maria Perez",
                        SedeCiudad = "POPAYAN",
                        Horas = 2,
                        FechaSesion = new DateTime(2026, 11, 21)
                    }
                };

                using (var form = new FormVerVinculaciones(items, conn, null, AgModoDatos.Ag))
                {
                    // Act
                    await form.CargarViolacionesAsync();

                    // Assert
                    form.ViolacionesPorSesion.Should().BeEmpty();
                    form.ViolacionesGlobales.Should().HaveCount(1);
                    form.ViolacionesGlobales[0].IdGen.Should().Be(999);
                    form.ViolacionesGlobales[0].EsDura.Should().BeFalse();
                }
            }
        }

        [STATestMethod]
        public async Task ConstruirArbol_WhenViolacionesGlobalesExist_ShouldDisplayDedicatedRootNode()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, '');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, NULL, 5, 1, 'Infracción global de crédito', '{{""extra"":1}}', '');";
                    cmd.ExecuteNonQuery();
                }

                var items = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 1,
                        IdGen = 10,
                        IdCorrida = 1,
                        DocenteNombre = "Ana Gomez",
                        SedeCiudad = "CALI",
                        Semestre = "I",
                        AsignaturaNombre = "Matemáticas",
                        Horas = 4,
                        FechaSesion = new DateTime(2026, 11, 22),
                        HorarioSesion = "SABADO|08:00-12:00"
                    }
                };

                using (var form = new FormVerVinculaciones(items, conn, null, AgModoDatos.Ag))
                {
                    // Act
                    await form.CargarViolacionesAsync();
                    form.ConstruirArbol();

                    // Assert
                    var rootNodes = form.TreeVinculaciones.Nodes.Cast<TreeNode>().ToList();
                    var nodoGlobal = rootNodes.FirstOrDefault(n => n.Text == "⚠️ Violaciones Generales de la Corrida");
                    nodoGlobal.Should().NotBeNull();
                    nodoGlobal.Nodes.Count.Should().Be(1);
                    nodoGlobal.Nodes[0].Text.Should().Contain("5: Infracción global de crédito");
                }
            }
        }

        [STATestMethod]
        public async Task ConstruirArbolPorFechas_WhenViolacionesGlobalesExist_ShouldDisplayDedicatedRootNode()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, '');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, NULL, 8, 1, 'Infracción no enlazada a gen', '{}', '');";
                    cmd.ExecuteNonQuery();
                }

                var items = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 1,
                        IdGen = 10,
                        IdCorrida = 1,
                        DocenteNombre = "Ana Gomez",
                        SedeCiudad = "CALI",
                        Semestre = "I",
                        AsignaturaNombre = "Matemáticas",
                        Horas = 4,
                        FechaSesion = new DateTime(2026, 11, 22),
                        HorarioSesion = "SABADO|08:00-12:00"
                    }
                };

                using (var form = new FormVerVinculaciones(items, conn, null, AgModoDatos.Ag))
                {
                    // Act
                    await form.CargarViolacionesAsync();
                    form.ConstruirArbolPorFechas();

                    // Assert
                    var rootNodes = form.TreeVinculaciones.Nodes.Cast<TreeNode>().ToList();
                    var nodoGlobal = rootNodes.FirstOrDefault(n => n.Text == "⚠️ Violaciones Generales de la Corrida");
                    nodoGlobal.Should().NotBeNull();
                    nodoGlobal.Nodes.Count.Should().Be(1);
                    nodoGlobal.Nodes[0].Text.Should().Contain("8: Infracción no enlazada a gen");
                }
            }
        }

        [STATestMethod]
        public async Task CargarViolacionesAsync_WhenSubunidadFilterMatchesOrMismatch_ShouldFilterCorrectly()
        {
            // Arrange
            using (var conn = CreateInMemoryDb())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO ag_corrida (id, id_individuo_optimo, subunidad) VALUES (1, 10, 'BUENAVENTURA');
INSERT INTO ag_individuo (id, id_corrida) VALUES (10, 1);
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (10, NULL, 3, 1, 'Conflicto Buenaventura', '{}', 'BUENAVENTURA');";
                    cmd.ExecuteNonQuery();
                }

                // Subunidad diferente
                var itemsDistintos = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 1,
                        IdGen = 10,
                        IdCorrida = 1,
                        Subunidad = "CALI"
                    }
                };

                using (var form = new FormVerVinculaciones(itemsDistintos, conn, null, AgModoDatos.Ag))
                {
                    await form.CargarViolacionesAsync();
                    form.ViolacionesGlobales.Should().BeEmpty();
                }

                // Subunidad coincidente
                var itemsCoincidentes = new List<SesionListItem>
                {
                    new SesionListItem
                    {
                        IdSesion = 1,
                        IdGen = 10,
                        IdCorrida = 1,
                        Subunidad = "BUENAVENTURA"
                    }
                };

                using (var form = new FormVerVinculaciones(itemsCoincidentes, conn, null, AgModoDatos.Ag))
                {
                    await form.CargarViolacionesAsync();
                    form.ViolacionesGlobales.Should().HaveCount(1);
                    form.ViolacionesGlobales[0].Mensaje.Should().Be("Conflicto Buenaventura");
                }
            }
        }
    }
}
