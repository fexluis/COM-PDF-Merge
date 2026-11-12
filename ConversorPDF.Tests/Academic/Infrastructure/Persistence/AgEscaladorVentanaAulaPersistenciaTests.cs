using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgEscaladorVentanaAulaPersistenciaTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private readonly Func<IDbConnection> _connectionFactory;

        public AgEscaladorVentanaAulaPersistenciaTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgEscaladorPersistencia_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";
            _connectionFactory = () => new SQLiteConnection(_connectionString);

            InicializarEsquema();
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                {
                    File.Delete(_dbPath);
                }
            }
            catch
            {
                // Limpieza defensiva en temp
            }
        }

        private void InicializarEsquema()
        {
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);

CREATE TABLE ag_arbol_nodo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idarbol INTEGER NOT NULL DEFAULT 1,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);");
            }
        }

        [Theory]
        [InlineData("LUN-JUE")]
        [InlineData("MAR-VIE")]
        public async Task AplicarAsync_WhenEscaladoALunVieConDemanda5W_ShouldActualizarAulaYNodoActivoPreservandoOtrosArboles(string ventanaInicial)
        {
            // Arrange
            string aulaRef = "A101";
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (@Ref, 'MANANA', 4, @V, '');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (1, @Ref, 'MANANA', 4, @V, '');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (2, @Ref, 'MANANA', 4, @V, '');",
                    new { Ref = aulaRef, V = ventanaInicial });
            }

            int w = 4;
            int demandaSesiones = 5 * w; // 20 sesiones para ventana de 16 -> requiere LUN-VIE
            var contexto = CrearContextoNominal(w, ventanaInicial, demandaSesiones, aulaRef);
            var sut = new AgEscaladorVentanaAula(_connectionFactory);

            // Act 1: Evaluar
            var decisiones = sut.Evaluar(contexto);

            decisiones.Should().ContainSingle();
            decisiones[0].Desde.Should().Be(ventanaInicial);
            decisiones[0].Hacia.Should().Be("LUN-VIE");
            decisiones[0].Aplica.Should().BeTrue();

            // Act 2: Aplicar en idArbol = 1
            int filasAfectadas = await sut.AplicarAsync(decisiones, 1, CancellationToken.None);
            filasAfectadas.Should().Be(1);

            // Assert por consulta a base de datos
            using (var conn = _connectionFactory())
            {
                conn.Open();
                string aulaVentana = conn.QuerySingle<string>(
                    "SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref", new { Ref = aulaRef });
                string nodoArbol1 = conn.QuerySingle<string>(
                    "SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 1", new { Ref = aulaRef });
                string nodoArbol2 = conn.QuerySingle<string>(
                    "SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 2", new { Ref = aulaRef });

                aulaVentana.Should().Be("LUN-VIE");
                nodoArbol1.Should().Be("LUN-VIE");
                nodoArbol2.Should().Be(ventanaInicial, "el árbol inactivo no debe ser modificado");
            }

            // Assert por contexto recargado: la nueva evaluación debe ser vacía (idempotente)
            var contextoRecargado = CrearContextoNominal(w, "LUN-VIE", demandaSesiones, aulaRef);
            var decisionesRecargadas = sut.Evaluar(contextoRecargado);
            decisionesRecargadas.Should().BeEmpty();
        }

        [Fact]
        public async Task AplicarAsync_WhenCicloCompletoVieDomConDemanda6W_ShouldEscalarEnDosPasosYTerminarVacio()
        {
            // Arrange
            string aulaRef = "A201";
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (1, @Ref, 'MANANA', 4, 'VIE-DOM', '');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (2, @Ref, 'MANANA', 4, 'VIE-DOM', '');",
                    new { Ref = aulaRef });
            }

            int w = 4;
            int demandaSesiones = 6 * w; // 24 sesiones: VIE-DOM tiene 16, JUE-DOM tiene 20, MIE-DOM tiene 24
            var sut = new AgEscaladorVentanaAula(_connectionFactory);

            // --- Ciclo 1: VIE-DOM -> JUE-DOM ---
            var ctx1 = CrearContextoNominal(w, "VIE-DOM", demandaSesiones, aulaRef);
            var d1 = sut.Evaluar(ctx1);
            d1.Should().ContainSingle();
            d1[0].Desde.Should().Be("VIE-DOM");
            d1[0].Hacia.Should().Be("JUE-DOM");
            d1[0].CapacidadAntes.Should().Be(16);
            d1[0].CapacidadDespues.Should().Be(20);

            await sut.AplicarAsync(d1, 1, CancellationToken.None);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref", new { Ref = aulaRef })
                    .Should().Be("JUE-DOM");
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 1", new { Ref = aulaRef })
                    .Should().Be("JUE-DOM");
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 2", new { Ref = aulaRef })
                    .Should().Be("VIE-DOM");
            }

            // --- Ciclo 2: JUE-DOM -> MIE-DOM ---
            var ctx2 = CrearContextoNominal(w, "JUE-DOM", demandaSesiones, aulaRef);
            var d2 = sut.Evaluar(ctx2);
            d2.Should().ContainSingle();
            d2[0].Desde.Should().Be("JUE-DOM");
            d2[0].Hacia.Should().Be("MIE-DOM");
            d2[0].CapacidadAntes.Should().Be(20);
            d2[0].CapacidadDespues.Should().Be(24);

            await sut.AplicarAsync(d2, 1, CancellationToken.None);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref", new { Ref = aulaRef })
                    .Should().Be("MIE-DOM");
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 1", new { Ref = aulaRef })
                    .Should().Be("MIE-DOM");
                conn.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE aula_ref = @Ref AND idarbol = 2", new { Ref = aulaRef })
                    .Should().Be("VIE-DOM");
            }

            // --- Ciclo 3: MIE-DOM ya satisface 24 sesiones o está en tope -> sin propuestas ---
            var ctx3 = CrearContextoNominal(w, "MIE-DOM", demandaSesiones, aulaRef);
            var d3 = sut.Evaluar(ctx3);
            d3.Should().BeEmpty();
        }

        private static AgContexto CrearContextoNominal(int w, string diasVentana, int demandaSesiones, string aulaRef)
        {
            var fInicio = new DateTime(2026, 8, 3);
            var fFin = fInicio.AddDays(7 * w - 1);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-2", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Test",
                Modalidad = "PRESENCIAL"
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = aulaRef,
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30,
                IdPrograma = 1,
                Semestre = "I"
            };
            contexto.Aulas.Add(aula);

            for (int d = 1; d <= 7; d++)
            {
                if (d == 6)
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 61,
                        DiaSemanaNumero = 6,
                        Franja = "MANANA",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = 62,
                        DiaSemanaNumero = 6,
                        Franja = "TARDE",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                }
                else
                {
                    contexto.Bloques.Add(new AgBloqueHorario
                    {
                        Id = d,
                        DiaSemanaNumero = d,
                        Franja = "MANANA",
                        DuracionHoras = 4,
                        Activo = 1
                    });
                }
            }

            double horasTotales = demandaSesiones * 4.0;
            contexto.Asignaturas.Add(new Asignatura
            {
                Id = 101,
                Nombre = "Asig Demanda",
                IdPrograma = 1,
                Semestre = "I",
                Horas = (int)horasTotales
            });

            contexto.AulaAsignaturaPorAulaAsig.Add((1, 101), new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 101,
                HorasObjetivo = horasTotales,
                OrdenCurricular = 1
            });

            return contexto;
        }
    }
}
