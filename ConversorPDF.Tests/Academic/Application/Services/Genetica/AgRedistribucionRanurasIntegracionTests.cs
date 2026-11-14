using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Pruebas de integración SQLite para la persistencia transaccional de redistribución de ranuras,
    /// huella SHA256, idempotencia, rollback por discrepancia y protección de patrones FIJO y LEGADO.
    /// </summary>
    public class AgRedistribucionRanurasIntegracionTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private readonly Func<IDbConnection> _connectionFactory;

        public AgRedistribucionRanurasIntegracionTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgRedistribucionInt_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
);

CREATE TABLE asignaturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL DEFAULT '',
    idprograma INTEGER NOT NULL DEFAULT 0,
    semestre TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT '',
    creditos INTEGER NOT NULL DEFAULT 0,
    horas INTEGER NOT NULL DEFAULT 0,
    int_horaria INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE programas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL DEFAULT '',
    modalidad TEXT NOT NULL DEFAULT ''
);

CREATE TABLE ag_bloque_horario (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dia_semana_numero INTEGER NOT NULL DEFAULT 1,
    franja TEXT NOT NULL DEFAULT '',
    hora_inicio TEXT NOT NULL DEFAULT '',
    hora_fin TEXT NOT NULL DEFAULT '',
    duracion_horas REAL NOT NULL DEFAULT 4.0,
    activo INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE festivos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_inicio TEXT,
    fecha_fin TEXT,
    descripcion TEXT
);

CREATE TABLE presenciales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ciudad TEXT NOT NULL DEFAULT ''
);

CREATE TABLE periodos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL DEFAULT ''
);");
            }
        }

        // =========================================================================
        // Caso 4: Patrones automáticos -> Aprovechar días adicionales
        // =========================================================================
        [Fact]
        public async Task AplicarAsync_WhenPatronAutomaticoYVentanaEscalada_ShouldLiberarDiasPatronParaRedistribucion()
        {
            // Arrange: IDs alineados entre base de datos (1, 2) y contexto (1, 2)
            string aulaRef = "A104";
            long idAula = 0;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '', 1);", new { Ref = aulaRef });
                idAula = conn.QueryFirstOrDefault<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = aulaRef });

                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 1, 64, 1, '5,6', 'AUTOMATICO');
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 2, 16, 2, '7', 'FIJO');", new { IdAula = idAula });

                conn.Execute("INSERT INTO ag_arbol (id, nombre) VALUES (1, 'Perfil Test');");
                conn.Execute(@"
INSERT INTO ag_arbol_nodo (idarbol, aula_ref, subunidad, dias_ventana)
VALUES (1, @Ref, '', 'VIE-DOM');", new { Ref = aulaRef });
            }

            var ctx = CrearContextoSintetico(4, "VIE-DOM", numAsignaturas: 2, sesionesPorAsignatura: 10, aulaRef: aulaRef);
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].HorasObjetivo = 64;
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "5,6";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "AUTOMATICO";
            ctx.Asignaturas.First(a => a.Id == 1).Horas = 64;
            ctx.Asignaturas.First(a => a.Id == 1).IntHoraria = 64;

            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].HorasObjetivo = 16;
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].DiasPatron = "7";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron = "FIJO";
            ctx.Asignaturas.First(a => a.Id == 2).Horas = 16;
            ctx.Asignaturas.First(a => a.Id == 2).IntHoraria = 16;

            var sut = new AgEscaladorRanuraAula(_connectionFactory);

            // Act
            var decisiones = sut.Evaluar(ctx, 1);
            decisiones.Should().ContainSingle();

            int confirmadas = await sut.AplicarAsync(decisiones, 1, CancellationToken.None);

            // Assert
            confirmadas.Should().Be(1);
            using (var conn = _connectionFactory())
            {
                conn.Open();
                string ventana = conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE id=@Id;", new { Id = idAula });
                ventana.Should().Be("JUE-DOM");

                string patronAuto = conn.QuerySingle<string>(
                    "SELECT dias_patron FROM ag_aula_asignatura WHERE id_aula=@Id AND id_asignatura=1;", new { Id = idAula });
                patronAuto.Should().Be(string.Empty, "el patrón AUTOMATICO debe limpiarse para aprovechar los días adicionales");

                string patronFijo = conn.QuerySingle<string>(
                    "SELECT dias_patron FROM ag_aula_asignatura WHERE id_aula=@Id AND id_asignatura=2;", new { Id = idAula });
                patronFijo.Should().Be("7", "el patrón FIJO debe conservarse inalterado");
            }
        }

        // =========================================================================
        // Caso 6: Huella modificada -> Rechazar guardado completo
        // =========================================================================
        [Fact]
        public async Task AplicarDetalladoAsync_WhenHuellaModificada_ShouldRechazarGuardadoCompleto()
        {
            // Arrange: IDs alineados (1) entre base y contexto
            string aulaRef = "A106";
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '', 1);", new { Ref = aulaRef });
                long idAula = conn.QuerySingle<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = aulaRef });
                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 1, 80, 1, '', 'AUTOMATICO');", new { IdAula = idAula });
            }

            var ctx = CrearContextoSintetico(4, "VIE-DOM", numAsignaturas: 1, sesionesPorAsignatura: 20, aulaRef: aulaRef);
            var sut = new AgEscaladorRanuraAula(_connectionFactory);
            var decisiones = sut.Evaluar(ctx, 0);
            decisiones.Should().ContainSingle();

            // Mutar la base de datos externamente para simular cambio concurrente que invalida la huella
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute("UPDATE ag_aula SET dias_ventana='LUN-VIE' WHERE aula_ref=@Ref;", new { Ref = aulaRef });
            }

            // Act
            var res = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);

            // Assert: Rechazo completo y seguro
            res.Exito.Should().BeFalse();
            res.Error.Should().Contain("modificada");
            res.Confirmadas.Should().Be(0);
            res.Modificadas.Should().Be(0);
        }

        // =========================================================================
        // Caso 7: Segunda aplicación -> Sin nuevas escrituras (Idempotencia)
        // =========================================================================
        [Fact]
        public async Task AplicarDetalladoAsync_WhenSegundaAplicacion_ShouldSerIdempotenteSinNuevasEscrituras()
        {
            // Arrange: IDs alineados (1) entre base y contexto
            string aulaRef = "A107";
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '', 1);", new { Ref = aulaRef });
                long idAula = conn.QuerySingle<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = aulaRef });
                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 1, 80, 1, '', 'AUTOMATICO');", new { IdAula = idAula });
            }

            var ctx = CrearContextoSintetico(4, "VIE-DOM", numAsignaturas: 1, sesionesPorAsignatura: 20, aulaRef: aulaRef);
            var sut = new AgEscaladorRanuraAula(_connectionFactory);
            var decisiones = sut.Evaluar(ctx, 0);

            // Act 1: Primera aplicación
            var res1 = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);
            res1.Exito.Should().BeTrue();
            res1.Confirmadas.Should().Be(1);
            res1.Modificadas.Should().Be(1);

            // Act 2: Segunda aplicación con el mismo lote de decisiones
            var res2 = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);

            // Assert: Idempotencia estricta: confirmada sin nuevas escrituras
            res2.Exito.Should().BeTrue();
            res2.Confirmadas.Should().Be(1);
            res2.Modificadas.Should().Be(0, "la reaplicación idempotente no debe realizar escrituras redundantes");
        }

        // =========================================================================
        // Autorización Parcial: Actualizar exclusivamente relaciones autorizadas
        // =========================================================================
        [Fact]
        public async Task AplicarDetalladoAsync_WhenAutorizacionParcial_ShouldActualizarExclusivamenteRelacionesAutorizadas()
        {
            // Arrange: Dos materias LEGADO en el aula, solo se autoriza la materia 1
            string aulaRef = "A108";
            long idAula = 0;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@Ref, 'MANANA', 4, 'VIE-DOM', '', 1);", new { Ref = aulaRef });
                idAula = conn.QuerySingle<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = aulaRef });

                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 1, 48, 1, '7', 'LEGADO');
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdAula, 2, 48, 2, '7', 'LEGADO');", new { IdAula = idAula });
            }

            var ctx = CrearContextoSintetico(4, "VIE-DOM", numAsignaturas: 2, sesionesPorAsignatura: 12, aulaRef: aulaRef);
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron = "7";
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron = "LEGADO";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].DiasPatron = "7";
            ctx.AulaAsignaturaPorAulaAsig[(1, 2)].OrigenPatron = "LEGADO";

            var sut = new AgEscaladorRanuraAula(_connectionFactory);
            var decisiones = sut.Evaluar(ctx, 0);

            // Autorizar exclusivamente la materia 1
            var patron1 = decisiones[0].ReclasificacionesPropuestas.First(p => p.IdAsignatura == 1);
            patron1.Autorizado = true;

            var autorizadas = new List<AgReclasificacionPatron> { patron1 };

            // Simular y evaluar con la sobrecarga oficial que sella la huella automáticamente
            var decisionesConfirmadas = sut.Evaluar(ctx, 0, autorizadas);

            // Act
            var res = await sut.AplicarDetalladoAsync(decisionesConfirmadas, 0, CancellationToken.None);

            // Assert
            res.Exito.Should().BeTrue();
            using (var conn = _connectionFactory())
            {
                conn.Open();
                var aa1 = conn.QuerySingle<(string dias, string origen)>(
                    "SELECT dias_patron, origen_patron FROM ag_aula_asignatura WHERE id_aula=@Id AND id_asignatura=1;", new { Id = idAula });
                aa1.origen.Should().Be("AUTOMATICO", "la relación autorizada debe transicionar a AUTOMATICO");
                aa1.dias.Should().Be(string.Empty);

                var aa2 = conn.QuerySingle<(string dias, string origen)>(
                    "SELECT dias_patron, origen_patron FROM ag_aula_asignatura WHERE id_aula=@Id AND id_asignatura=2;", new { Id = idAula });
                aa2.origen.Should().Be("LEGADO", "la relación no autorizada debe permanecer intacta como LEGADO");
                aa2.dias.Should().Be("7");
            }
        }

        // =========================================================================
        // Un aula seleccionada y otra descartada (Paso 4): no revierte el lote
        // =========================================================================
        [Fact]
        public async Task AplicarDetalladoAsync_WhenUnAulaSeleccionadaYOtraDescartada_ShouldAplicarSoloLaSeleccionadaSinRevertirLote()
        {
            // Arrange: Aula A (seleccionada para escalado) y Aula B (descartada o terminal sin cambios)
            string refA = "AULA_A";
            string refB = "AULA_B";
            long idA = 0;
            long idB = 0;
            using (var conn = _connectionFactory())
            {
                conn.Open();
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@RefA, 'MANANA', 4, 'VIE-DOM', '', 1);
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad, id_programa)
VALUES (@RefB, 'MANANA', 4, 'MIE-DOM', '', 1);", new { RefA = refA, RefB = refB });
                idA = conn.QuerySingle<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = refA });
                idB = conn.QuerySingle<long>("SELECT id FROM ag_aula WHERE aula_ref=@Ref;", new { Ref = refB });

                conn.Execute(@"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdA, 1, 80, 1, '', 'AUTOMATICO');
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, horas_objetivo, orden_curricular, dias_patron, origen_patron)
VALUES (@IdB, 2, 48, 1, '7', 'LEGADO');", new { IdA = idA, IdB = idB });
            }

            var ctx = CrearContextoSintetico(4, "VIE-DOM", numAsignaturas: 1, sesionesPorAsignatura: 20, aulaRef: refA);
            ctx.Aulas[0].Id = (int)idA;

            var aulaB = new AgAula
            {
                Id = (int)idB,
                AulaRef = refB,
                Subunidad = "",
                DiasVentana = "MIE-DOM",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                IdPrograma = 1
            };
            ctx.Aulas.Add(aulaB);

            var aaB = new AgAulaAsignatura
            {
                IdAula = (int)idB,
                IdAsignatura = 2,
                HorasObjetivo = 48,
                OrdenCurricular = 1,
                DiasPatron = "7",
                OrigenPatron = "LEGADO"
            };
            ctx.AulaAsignaturas.Add(aaB);
            ctx.AulaAsignaturaPorAulaAsig.Add(((int)idB, 2), aaB);

            ctx.Asignaturas.Add(new Asignatura
            {
                Id = 2,
                Nombre = "Materia B",
                IdPrograma = 1,
                Horas = 48,
                IntHoraria = 48
            });
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(_connectionFactory);
            var decisiones = sut.Evaluar(ctx, 0);

            // Aula A tiene Aplica=true (escala a JUE-DOM); Aula B en MIE-DOM terminal sin autorizaciones tiene Aplica=false
            decisiones.Should().HaveCount(2);
            var decA = decisiones.First(d => d.AulaRef == refA);
            var decB = decisiones.First(d => d.AulaRef == refB);
            decA.Aplica.Should().BeTrue();
            decB.Aplica.Should().BeFalse();

            // Act: Aplicar el lote conteniendo tanto el aula aplicable como el aula descartada
            var res = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);

            // Assert: Aplica exitosamente el aula A sin revertir el lote por la presencia de aula B
            res.Exito.Should().BeTrue();
            res.Confirmadas.Should().Be(1);
            using (var conn = _connectionFactory())
            {
                conn.Open();
                string ventanaA = conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE id=@Id;", new { Id = idA });
                ventanaA.Should().Be("JUE-DOM", "el aula seleccionada debe actualizar su ventana");

                string ventanaB = conn.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE id=@Id;", new { Id = idB });
                ventanaB.Should().Be("MIE-DOM", "el aula descartada debe permanecer intacta en su ventana terminal");
            }
        }

        // =========================================================================
        // Conservación de FIJO: Rechazar cualquier intento de reclasificar FIJO
        // =========================================================================
        [Fact]
        public void ReclasificacionPatron_WhenPatronEsFijo_ShouldRechazarTransicion()
        {
            // Arrange
            var reclasificacion = new AgReclasificacionPatron
            {
                IdAula = 1,
                AulaRef = "A101",
                IdAsignatura = 1,
                NombreAsignatura = "Materia Fija",
                OrigenAnterior = "FIJO",
                OrigenPropuesto = "AUTOMATICO",
                DiasPatronAnterior = "7",
                DiasPatronPropuesto = "",
                Autorizado = true
            };

            // Act
            bool esValida = reclasificacion.EsTransicionValida(out string error);

            // Assert
            esValida.Should().BeFalse();
            error.Should().Contain("FIJO");
        }

        private static AgContexto CrearContextoSintetico(
            int w,
            string diasVentana,
            int numAsignaturas,
            int sesionesPorAsignatura,
            int horasDia = 4,
            DateTime? fechaInicio = null,
            string aulaRef = "A101")
        {
            var fInicio = fechaInicio ?? new DateTime(2027, 2, 1);
            var fFin = fInicio.AddDays(7 * w - 1);

            var contexto = new AgContexto();
            contexto.IdCorrida = 1;
            contexto.IdSede = 1;
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", fInicio, fFin, AgOrigenVentana.Presencial).Value;

            contexto.Programas.Add(new Programa
            {
                Id = 1,
                Nombre = "Programa Aceptación",
                Modalidad = "PRESENCIAL"
            });

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = aulaRef,
                Subunidad = "",
                DiasVentana = diasVentana,
                Jornada = "MANANA",
                HorasDia = horasDia,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 35,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I"
            };
            contexto.Aulas.Add(aula);

            for (int d = 1; d <= 7; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario
                {
                    Id = d,
                    DiaSemanaNumero = d,
                    Franja = "MANANA",
                    HoraInicio = "08:00",
                    HoraFin = "12:00",
                    DuracionHoras = horasDia,
                    Activo = 1
                });
            }

            for (int i = 1; i <= numAsignaturas; i++)
            {
                int horas = sesionesPorAsignatura * horasDia;
                var asig = new Asignatura
                {
                    Id = i,
                    Nombre = $"Materia {i}",
                    IdPrograma = 1,
                    Semestre = "I",
                    Horas = horas,
                    IntHoraria = horas
                };
                contexto.Asignaturas.Add(asig);

                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = horas,
                    OrdenCurricular = i,
                    DiasPatron = string.Empty,
                    OrigenPatron = "AUTOMATICO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig.Add((1, i), aa);
            }

            contexto.Docentes.Add(new Tercero { Id = 10, Nombre = "Docente 10", Residencia = "CALI" });
            contexto.Sedes.Add(new Sede { Id = 1, Ciudad = "CALI" });
            contexto.Presenciales.Add((fInicio, fFin));

            contexto.ConstruirIndices();
            return contexto;
        }
    }
}

