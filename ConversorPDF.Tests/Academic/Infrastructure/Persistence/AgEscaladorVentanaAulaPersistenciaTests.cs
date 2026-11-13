using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using CSharpFunctionalExtensions;
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
CREATE TABLE ag_arbol (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL DEFAULT '',
    descripcion TEXT NOT NULL DEFAULT '',
    idperiodo INTEGER NOT NULL DEFAULT 0,
    fecha_inicio TEXT,
    fecha_fin TEXT,
    ventana_origen TEXT NOT NULL DEFAULT '',
    manual_ventana INTEGER NOT NULL DEFAULT 0,
    poblacion INTEGER NOT NULL DEFAULT 0,
    generaciones INTEGER NOT NULL DEFAULT 0,
    tasa_cruce REAL NOT NULL DEFAULT 0,
    tasa_mutacion REAL NOT NULL DEFAULT 0,
    elitismo INTEGER NOT NULL DEFAULT 0,
    estancamiento_max INTEGER NOT NULL DEFAULT 0,
    semilla INTEGER NOT NULL DEFAULT 0,
    usuario TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT '',
    creado_en TEXT,
    modificado_en TEXT,
    activo INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT '',
    id_programa INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE ag_arbol_nodo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idarbol INTEGER NOT NULL DEFAULT 1,
    idperiodo INTEGER NOT NULL DEFAULT 0,
    idsede INTEGER NOT NULL DEFAULT 0,
    idprograma INTEGER NOT NULL DEFAULT 0,
    semestre TEXT NOT NULL DEFAULT '',
    grupo TEXT NOT NULL DEFAULT '',
    modalidad TEXT NOT NULL DEFAULT '',
    modalidad_efectiva TEXT NOT NULL DEFAULT '',
    aula_ref TEXT NOT NULL DEFAULT '',
    capacidad INTEGER NOT NULL DEFAULT 0,
    tipo_aula TEXT NOT NULL DEFAULT '',
    caracteristicas_json TEXT NOT NULL DEFAULT '',
    codigo TEXT NOT NULL DEFAULT '',
    nombre TEXT NOT NULL DEFAULT '',
    es_virtual INTEGER NOT NULL DEFAULT 0,
    habilitado INTEGER NOT NULL DEFAULT 1,
    subunidad TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT ''
);

CREATE TABLE ag_aula_asignatura (
    id_aula INTEGER,
    id_asignatura INTEGER,
    exclusivo INTEGER NOT NULL DEFAULT 0,
    prioridad INTEGER NOT NULL DEFAULT 0,
    horas_objetivo REAL NOT NULL DEFAULT 0.0,
    orden_curricular INTEGER NOT NULL DEFAULT 0,
    requerida INTEGER NOT NULL DEFAULT 1,
    dias_patron TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT '',
    origen_patron TEXT NOT NULL DEFAULT 'LEGADO'
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

            // Assert por contexto recargado leyendo persistencia real: la nueva evaluación debe ser vacía (idempotente)
            string persistida;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                persistida = conn.QuerySingle<string>(
                    "SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref",
                    new { Ref = aulaRef });
            }

            var contextoRecargado = CrearContextoNominal(w, persistida, demandaSesiones, aulaRef);
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

            // --- Ciclo 2: JUE-DOM -> MIE-DOM (leyendo persistencia real) ---
            string persistidaCiclo1;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                persistidaCiclo1 = conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref", new { Ref = aulaRef });
            }
            var ctx2 = CrearContextoNominal(w, persistidaCiclo1, demandaSesiones, aulaRef);
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

            // --- Ciclo 3: MIE-DOM ya satisface 24 sesiones o está en tope (leyendo persistencia real) ---
            string persistidaCiclo2;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                persistidaCiclo2 = conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE aula_ref = @Ref", new { Ref = aulaRef });
            }
            var ctx3 = CrearContextoNominal(w, persistidaCiclo2, demandaSesiones, aulaRef);
            var d3 = sut.Evaluar(ctx3);
            d3.Should().BeEmpty();
        }

        [Fact]
        public async Task FlujoIntegracion_Escalar_RecargarDto_CambiarJornada_Guardar_Recargar_ShouldConservarVentanaYPerfilAjenoIntacto()
        {
            // Arrange
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_arbol (id, nombre, descripcion) VALUES (1, 'Perfil 1', 'Test 1');
INSERT INTO ag_arbol (id, nombre, descripcion) VALUES (2, 'Perfil 2', 'Test 2');

INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@Ref, 'MANANA', 4, 'LUN-JUE', '', 1);

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad, idperiodo, idsede, idprograma, semestre, grupo, modalidad, modalidad_efectiva)
VALUES (1, @Ref, 'MANANA', 4, 'LUN-JUE', '', 1, 1, 1, '1', '1', 'PRESENCIAL', 'PRESENCIAL');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad, idperiodo, idsede, idprograma, semestre, grupo, modalidad, modalidad_efectiva)
VALUES (2, @Ref, 'MANANA', 4, 'LUN-JUE', '', 1, 1, 1, '1', '1', 'PRESENCIAL', 'PRESENCIAL');",
                    new { Ref = aulaRef });
            }

            var repo = new AgArbolRepository(_connectionFactory);
            var servicio = new AgArbolService(repo);
            var sut = new AgEscaladorVentanaAula(_connectionFactory);
            var sync = new AgAulaDimensionesSync(_connectionFactory);

            // 1. Escalar de LUN-JUE a LUN-VIE en idArbol = 1
            var ctx = CrearContextoNominal(4, "LUN-JUE", 20, aulaRef);
            var decisiones = sut.Evaluar(ctx);
            decisiones.Should().ContainSingle();
            decisiones[0].Hacia.Should().Be("LUN-VIE");

            int aplicadas = await sut.AplicarAsync(decisiones, 1, CancellationToken.None);
            aplicadas.Should().Be(1);

            // 2. Recargar DTO con servicio real
            var carga1 = await servicio.CargarPerfilAsync(1, CancellationToken.None);
            carga1.IsSuccess.Should().BeTrue();
            var dtos1 = carga1.Value.Item2;
            var dtoAula = dtos1.Single(d => string.Equals(d.AulaRef, aulaRef, StringComparison.OrdinalIgnoreCase));
            dtoAula.DiasVentana.Should().Be("LUN-VIE");

            // 3. Mapear a AgNodoArbolKey y cambiar Jornada a 'TARDE'
            var key = new AgNodoArbolKey(6, 1, "PRESENCIAL", "PRESENCIAL");
            var mapRes = AgNodoArbolKeyMapper.AplicarDimensiones(key, dtoAula);
            mapRes.IsSuccess.Should().BeTrue();
            key.DiasVentana.Should().Be("LUN-VIE");

            // Mutar jornada en la clave y persistir con AgAulaDimensionesSync
            key.Jornada = "TARDE";
            int sincronizadas = await sync.SincronizarAsync(aulaRef, key.Jornada, key.HorasDia, key.DiasVentana, string.Empty, CancellationToken.None);
            sincronizadas.Should().BeGreaterThan(0);

            // Comprobar persistencia en ag_aula
            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aulaPersistida = conn.QuerySingle<AgAula>(
                    "SELECT jornada, dias_ventana as DiasVentana, horas_dia as HorasDia FROM ag_aula WHERE aula_ref = @Ref",
                    new { Ref = aulaRef });
                aulaPersistida.Jornada.Should().Be("TARDE");
                aulaPersistida.DiasVentana.Should().Be("LUN-VIE");
            }

            // Fase 1725: Persistir también el nodo en ag_arbol_nodo del perfil 1 para reflejar el cambio en el árbol
            var nodos = await repo.ObtenerNodosAsync(1, CancellationToken.None);
            nodos.IsSuccess.Should().BeTrue();
            var nodo = nodos.Value.Single(n => string.Equals(n.AulaRef, aulaRef, StringComparison.OrdinalIgnoreCase));
            nodo.Jornada = key.Jornada;
            nodo.HorasDia = key.HorasDia;
            nodo.DiasVentana = key.DiasVentana;

            var guardado = await repo.GuardarAsync(carga1.Value.Item1, nodos.Value, CancellationToken.None);
            guardado.IsSuccess.Should().BeTrue();

            // 4. Recargar y verificar que DiasVentana conserva LUN-VIE y perfil 2 sigue intacto
            var carga2 = await servicio.CargarPerfilAsync(1, CancellationToken.None);
            carga2.IsSuccess.Should().BeTrue();
            var dtoRecargado = carga2.Value.Item2.Single(d => string.Equals(d.AulaRef, aulaRef, StringComparison.OrdinalIgnoreCase));
            dtoRecargado.DiasVentana.Should().Be("LUN-VIE");
            dtoRecargado.Jornada.Should().Be("TARDE");

            var cargaPerfil2 = await servicio.CargarPerfilAsync(2, CancellationToken.None);
            cargaPerfil2.IsSuccess.Should().BeTrue();
            var dtoPerfil2 = cargaPerfil2.Value.Item2.Single(d => string.Equals(d.AulaRef, aulaRef, StringComparison.OrdinalIgnoreCase));
            dtoPerfil2.DiasVentana.Should().Be("LUN-JUE", "el árbol 2 debe permanecer intacto");
        }

        [Fact]
        public async Task AplicarAsync_WhenDecisionesInexistentesEnBD_ShouldRetornarCeroFilasAfectadas()
        {
            var sut = new AgEscaladorVentanaAula(_connectionFactory);
            var decisionesInexistentes = new List<AgEscaladorVentanaAula.Decision>
            {
                new AgEscaladorVentanaAula.Decision
                {
                    AulaRef = "NO_EXISTE_999",
                    Desde = "LUN-JUE",
                    Hacia = "LUN-VIE",
                    CapacidadAntes = 16,
                    CapacidadDespues = 20,
                    Motivo = "Test"
                }
            };

            int aplicadas = await sut.AplicarAsync(decisionesInexistentes, 1, CancellationToken.None);
            aplicadas.Should().Be(0, "si el aula no existe en BD, las filas afectadas deben ser 0");
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenValoresDifieren_ShouldActualizarYRetornarFilasAfectadas()
        {
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, subunidad, jornada, dias_ventana, horas_dia, id_programa)
VALUES (@Ref, @Sub, 'MANANA', 'LUN-JUE', 4, 1);",
                    new { Ref = aulaRef, Sub = "" });
            }

            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync(aulaRef, "TARDE", 8, "LUN-VIE", string.Empty, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().Be(1);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aula = conn.QuerySingle<AgAula>(
                    "SELECT jornada, dias_ventana as DiasVentana, horas_dia as HorasDia FROM ag_aula WHERE aula_ref = @Ref",
                    new { Ref = aulaRef });
                aula.Jornada.Should().Be("TARDE");
                aula.DiasVentana.Should().Be("LUN-VIE");
                aula.HorasDia.Should().Be(8);
            }
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenValoresCoinciden_ShouldRetornarExitoCeroFilas()
        {
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, subunidad, jornada, dias_ventana, horas_dia, id_programa)
VALUES (@Ref, @Sub, 'MANANA', 'LUN-JUE', 4, 1);",
                    new { Ref = aulaRef, Sub = "" });
            }

            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync(aulaRef, "MANANA", 4, "LUN-JUE", string.Empty, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenAulaNoExiste_ShouldRetornarFailure()
        {
            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync("AULA_INEXISTENTE_999", "MANANA", 4, "LUN-VIE", string.Empty, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no existe en la base de datos");
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenErrorSql_ShouldRetornarFailure()
        {
            Func<IDbConnection> connFalla = () =>
            {
                var conn = new SQLiteConnection("Data Source=:memory:;");
                conn.Open();
                // Tabla ag_aula no creada intencionalmente para inducir fallo SQL
                return conn;
            };

            var sync = new AgAulaDimensionesSync(connFalla);
            var res = await sync.SincronizarConfirmadoAsync("AULA_CUALQUIERA", "MANANA", 4, "LUN-VIE", string.Empty, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("Error al sincronizar dimensiones");
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenSubunidadEspecificaCoincide_ShouldActualizarCorrectamente()
        {
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, subunidad, jornada, dias_ventana, horas_dia, id_programa)
VALUES (@Ref, @Sub, 'MANANA', 'LUN-JUE', 4, 1);",
                    new { Ref = aulaRef, Sub = "CENTRAL" });
            }

            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync(aulaRef, "TARDE", 8, "LUN-VIE", "CENTRAL", CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().Be(1);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aula = conn.QuerySingle<AgAula>(
                    "SELECT jornada, dias_ventana as DiasVentana, horas_dia as HorasDia FROM ag_aula WHERE aula_ref = @Ref AND subunidad = 'CENTRAL'",
                    new { Ref = aulaRef });
                aula.Jornada.Should().Be("TARDE");
                aula.DiasVentana.Should().Be("LUN-VIE");
                aula.HorasDia.Should().Be(8);
            }
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenSubunidadDifiere_ShouldRetornarFailureYSinModificar()
        {
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, subunidad, jornada, dias_ventana, horas_dia, id_programa)
VALUES (@Ref, @Sub, 'MANANA', 'LUN-JUE', 4, 1);",
                    new { Ref = aulaRef, Sub = "CENTRAL" });
            }

            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync(aulaRef, "TARDE", 8, "LUN-VIE", "NORTE", CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no existe en la base de datos");

            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aula = conn.QuerySingle<AgAula>(
                    "SELECT jornada, dias_ventana as DiasVentana, horas_dia as HorasDia FROM ag_aula WHERE aula_ref = @Ref AND subunidad = 'CENTRAL'",
                    new { Ref = aulaRef });
                aula.Jornada.Should().Be("MANANA", "no debe mutar si la subunidad difiere");
            }
        }

        [Fact]
        public async Task SincronizarConfirmadoAsync_WhenFilaLegacyConSubunidadVacia_ShouldActualizarConSubunidadActiva()
        {
            string aulaRef = AgAulaFactory.ConstruirAulaRef(1, 1, 1, "1", "1", "PRESENCIAL");
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, subunidad, jornada, dias_ventana, horas_dia, id_programa)
VALUES (@Ref, @Sub, 'MANANA', 'LUN-JUE', 4, 1);",
                    new { Ref = aulaRef, Sub = "" });
            }

            var sync = new AgAulaDimensionesSync(_connectionFactory);
            var res = await sync.SincronizarConfirmadoAsync(aulaRef, "TARDE", 8, "LUN-VIE", "CENTRAL", CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().Be(1);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aula = conn.QuerySingle<AgAula>(
                    "SELECT jornada, dias_ventana as DiasVentana, horas_dia as HorasDia FROM ag_aula WHERE aula_ref = @Ref",
                    new { Ref = aulaRef });
                aula.Jornada.Should().Be("TARDE");
                aula.DiasVentana.Should().Be("LUN-VIE");
                aula.HorasDia.Should().Be(8);
            }
        }

        [Fact]
        public async Task AplicarAsync_WhenEscalaVentana_ShouldResetearPatronesAutomaticosYConservarFijos()
        {
            string aulaRef = "A102";
            long idAula = 0;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '');", new { Ref = aulaRef });

                idAula = conn.QueryFirstOrDefault<long>("SELECT id FROM ag_aula WHERE aula_ref = @Ref;", new { Ref = aulaRef });

                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, dias_patron, origen_patron)
VALUES (@IdAula, 101, '5,6', 'AUTOMATICO');

INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, dias_patron, origen_patron)
VALUES (@IdAula, 102, '7', 'FIJO');", new { IdAula = idAula });
            }

            var sut = new AgEscaladorVentanaAula(_connectionFactory);
            var decisiones = new List<AgEscaladorVentanaAula.Decision>
            {
                new AgEscaladorVentanaAula.Decision
                {
                    AulaRef = aulaRef,
                    Subunidad = "",
                    Desde = "VIE-DOM",
                    Hacia = "JUE-DOM",
                    CapacidadAntes = 68,
                    CapacidadDespues = 85,
                    IdAsignaturasBottleneck = new List<int> { 101 }
                }
            };

            int afectadas = await sut.AplicarAsync(decisiones, 0, CancellationToken.None);
            afectadas.Should().Be(1);

            using (var conn = _connectionFactory())
            {
                conn.Open();
                string ventana = conn.QueryFirstOrDefault<string>("SELECT dias_ventana FROM ag_aula WHERE id = @Id;", new { Id = idAula });
                ventana.Should().Be("JUE-DOM");

                string patronAuto = conn.QueryFirstOrDefault<string>(
                    "SELECT dias_patron FROM ag_aula_asignatura WHERE id_aula = @Id AND id_asignatura = 101;",
                    new { Id = idAula });
                patronAuto.Should().Be(string.Empty, "los patrones AUTOMATICO se deben limpiar para permitir redistribución en la nueva ventana");

                string patronFijo = conn.QueryFirstOrDefault<string>(
                    "SELECT dias_patron FROM ag_aula_asignatura WHERE id_aula = @Id AND id_asignatura = 102;",
                    new { Id = idAula });
                patronFijo.Should().Be("7", "los patrones FIJO se deben conservar intactos");
            }
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
