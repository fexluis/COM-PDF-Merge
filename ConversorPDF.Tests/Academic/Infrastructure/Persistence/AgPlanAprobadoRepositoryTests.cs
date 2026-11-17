using System;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.Tests.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgPlanAprobadoRepositoryTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "AgPlanLectura_" + Guid.NewGuid().ToString("N"));
        private readonly string _cs;
        private readonly AgPlanAprobadoRepository _repo;

        public AgPlanAprobadoRepositoryTests()
        {
            Directory.CreateDirectory(_dir);
            _cs = "Data Source=" + Path.Combine(_dir, "fixture.sqlite") + ";Pooling=False;";
            _repo = new AgPlanAprobadoRepository(() => new SQLiteConnection(_cs));
            using (var db = new SQLiteConnection(_cs))
            {
                db.Open();
                db.Execute(@"
CREATE TABLE territoriales (id INTEGER PRIMARY KEY, subunidad TEXT);
INSERT INTO territoriales VALUES (16,'VALLE'),(6,'CAUCA');
CREATE TABLE vinculaciones (id INTEGER PRIMARY KEY, idsede INTEGER, idprograma INTEGER, idperiodo INTEGER,
 idasignatura INTEGER, idtercero INTEGER, semestre TEXT, grupo TEXT, modalidadVinculacion TEXT,
 total_horas INTEGER, subunidad TEXT, fecha_inicio TEXT, fecha_fin TEXT);
CREATE TABLE sesiones (id INTEGER PRIMARY KEY, idvinculacion INTEGER, idtercero INTEGER,
 fechasesion TEXT, horariosesion TEXT, horas REAL, subunidad TEXT, viatica INTEGER);
CREATE TABLE ag_aula (id INTEGER PRIMARY KEY, idsede INTEGER, idprograma INTEGER, idperiodo INTEGER,
 semestre TEXT, grupo TEXT, modalidadVinculacion TEXT, subunidad TEXT);
INSERT INTO ag_aula VALUES (13,1,1,3,'I','A','Presencial','VALLE'),(14,1,1,3,'I','B','Presencial','VALLE');
INSERT INTO vinculaciones VALUES (1,1,1,3,7,91,'1','A','Presencial',68,'VALLE','2027-02-05','2027-06-01'),
 (2,1,1,3,7,92,'I','B','Presencial',68,'VALLE','','');
INSERT INTO sesiones VALUES (10,1,91,'2027-06-01','18:00 A 22:00',4,'VALLE',0);");
            }
        }

        private void Sql(string sql)
        {
            using (var db = new SQLiteConnection(_cs)) { db.Open(); db.Execute(sql); }
        }

        [Fact]
        public async Task LeeOrigenSinRecorteIncluyeVinculacionSinSesionesYNoAprueba()
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Equal("REQUIERE_RATIFICACION", r.Value.EstadoOrigen);
            Assert.Equal(2, r.Value.Relaciones.Count);
            var a = r.Value.Relaciones.Single(v => v.IdVinculacion == 1);
            Assert.Equal(13, Assert.Single(a.ClavesCandidatas).IdAula);
            var s = Assert.Single(a.Sesiones);
            Assert.Equal(10, s.IdSesion); Assert.Equal(91, s.IdTercero);
            Assert.Equal(new DateTime(2027, 6, 1), s.Fecha);
            Assert.Equal(TimeSpan.FromHours(18), s.HoraInicio);
            Assert.Equal(TimeSpan.FromHours(22), s.HoraFin); Assert.Equal(4, s.HorasLectivas);
            Assert.Null(s.RevisionAprobacion); Assert.Null(s.EstadoAprobacion);
            Assert.Empty(r.Value.Relaciones.Single(v => v.IdVinculacion == 2).Sesiones);
            Assert.Empty(r.Value.RelacionesAusentes);
            var segunda = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.Equal(r.Value.HuellaOrigen, segunda.Value.HuellaOrigen);
            using (var db = new SQLiteConnection(_cs))
            {
                db.Open(); Assert.Equal(1, db.ExecuteScalar<int>("SELECT COUNT(*) FROM sesiones"));
                Assert.Equal(2, db.ExecuteScalar<int>("SELECT COUNT(*) FROM vinculaciones"));
            }
        }

        [Fact]
        public async Task ReportaExtrasAusenciasYAmbiguedadSinOcultarFilas()
        {
            Sql(@"INSERT INTO vinculaciones VALUES (3,1,1,3,99,91,'I','A','Presencial',4,'VALLE','','');
INSERT INTO ag_aula VALUES (15,1,1,3,'I','A','Presencial','VALLE');
UPDATE vinculaciones SET subunidad='CAUCA' WHERE id=2;");
            var r = await _repo.ObtenerAsync(AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto()), CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Equal(2, r.Value.Relaciones.Count);
            Assert.DoesNotContain(r.Value.Relaciones, v => v.IdVinculacion == 2);
            Assert.Contains(r.Value.Diagnosticos, d => d == "RELACION_EXTRA:3");
            Assert.Contains(r.Value.Diagnosticos, d => d == "IDENTIDAD_AMBIGUA:1");
            Assert.Equal(14, Assert.Single(r.Value.RelacionesAusentes).IdAula);
            Assert.Equal("PLAN_NO_VERIFICABLE", r.Value.EstadoOrigen);
        }

        [Fact]
        public async Task LeeTodasLasFilasYHuellaEsInvariantePeroSensibleAlOrigenCompleto()
        {
            for (int id = 11; id <= 1011; id++)
                Sql("INSERT INTO sesiones VALUES (" + id + ",1,91,'2027-02-05','18:00 A 22:00',4,'VALLE',0)");
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("es-CO");
                var a = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                var b = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
                Assert.Equal(1002, a.Relaciones.Sum(v => v.Sesiones.Count));
                Assert.Equal(a.HuellaOrigen, b.HuellaOrigen);
                Sql("UPDATE vinculaciones SET total_horas=69 WHERE id=1");
                var c = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
                Assert.NotEqual(a.HuellaOrigen, c.HuellaOrigen);
                Sql("UPDATE sesiones SET viatica=1 WHERE id=10");
                var d = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
                Assert.NotEqual(c.HuellaOrigen, d.HuellaOrigen);
            }
            finally { CultureInfo.CurrentCulture = original; }
        }

        [Fact]
        public async Task DatosInvalidosNoDesaparecenYCancelacionNoEsAusencia()
        {
            Sql("UPDATE sesiones SET fechasesion='?', horariosesion='LEGADO', subunidad='' WHERE id=10");
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Empty(r.Value.Relaciones.First().Sesiones);
            Assert.Contains("SESION_SUBUNIDAD_INCONSISTENTE:10", r.Value.Diagnosticos);
            Assert.Equal("PLAN_NO_VERIFICABLE", r.Value.EstadoOrigen);
            var ct = new CancellationToken(true);
            var cancelada = await _repo.ObtenerAsync(m, ct);
            Assert.True(cancelada.IsFailure); Assert.Equal("LECTURA_PLAN_CANCELADA", cancelada.Error);
        }

        [Fact]
        public async Task ParticionAjenaConocidaNoCambiaSnapshotDiagnosticoNiHuella()
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var antes = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Sql(@"INSERT INTO ag_aula VALUES (50,1,1,3,'I','A','Presencial','CAUCA');
INSERT INTO vinculaciones VALUES (50,1,1,3,99,900,'I','A','Presencial',99,'CAUCA','AJENO','AJENO');
INSERT INTO sesiones VALUES (50,50,900,'AJENO','AJENO',99,'CAUCA',1),
 (51,999,900,'AJENO','AJENO',99,'CAUCA',1);");
            var despues = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Assert.Equal(antes.HuellaOrigen, despues.HuellaOrigen);
            Assert.Equal(antes.Diagnosticos.ToArray(), despues.Diagnosticos.ToArray());
            Assert.Equal(antes.EstadoOrigen, despues.EstadoOrigen);
            Assert.Equal(new long[] { 1, 2 }, despues.Relaciones.Select(v => v.IdVinculacion).ToArray());
            Assert.Single(despues.Relaciones.SelectMany(v => v.Sesiones));
        }

        [Theory]
        [InlineData("VALLE", "SESION_HUERFANA_SUBUNIDAD_SELECCIONADA:80")]
        [InlineData("", "SESION_HUERFANA_ORIGEN_DESCONOCIDO:80")]
        [InlineData(null, "SESION_HUERFANA_ORIGEN_DESCONOCIDO:80")]
        [InlineData("SIN_CATALOGO", "SESION_HUERFANA_ORIGEN_DESCONOCIDO:80")]
        public async Task HuerfanaBloqueaSinAtribuirAulaYCambiaHuella(string subunidad, string diagnostico)
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var antes = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            using (var db = new SQLiteConnection(_cs))
            {
                db.Open();
                db.Execute("INSERT INTO sesiones VALUES (80,999,900,'2027-02-05','18:00 A 22:00',4,@subunidad,0)", new { subunidad });
            }
            var despues = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Assert.Contains(diagnostico, despues.Diagnosticos);
            Assert.Equal("PLAN_NO_VERIFICABLE", despues.EstadoOrigen);
            Assert.NotEqual(antes.HuellaOrigen, despues.HuellaOrigen);
            Assert.DoesNotContain(despues.Relaciones.SelectMany(v => v.Sesiones), s => s.IdSesion == 80);
            Assert.Equal(2, despues.Relaciones.Count);
            Assert.Empty(despues.RelacionesAusentes);
            var repetida = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Assert.Equal(despues.HuellaOrigen, repetida.HuellaOrigen);
            if (subunidad == "VALLE")
            {
                Sql("UPDATE sesiones SET horas=8 WHERE id=80");
                var cambiada = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
                Assert.NotEqual(despues.HuellaOrigen, cambiada.HuellaOrigen);
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData("SIN_CATALOGO")]
        public async Task SubunidadDesconocidaSoloExponeIncertidumbreNoPayload(string subunidad)
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            using (var db = new SQLiteConnection(_cs))
            {
                db.Open();
                db.Execute("UPDATE vinculaciones SET subunidad=@subunidad, fecha_inicio='NO_COPIAR' WHERE id=2", new { subunidad });
                db.Execute("UPDATE ag_aula SET subunidad=@subunidad WHERE id=14", new { subunidad });
                db.Execute("UPDATE sesiones SET subunidad='CAUCA' WHERE id=10");
            }
            var r = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Assert.Single(r.Relaciones);
            Assert.Empty(r.Relaciones.SelectMany(v => v.Sesiones));
            Assert.Contains("VINCULACION_SUBUNIDAD_DESCONOCIDA:2", r.Diagnosticos);
            Assert.Contains("AULA_SUBUNIDAD_DESCONOCIDA:14", r.Diagnosticos);
            Assert.Contains("SESION_SUBUNIDAD_INCONSISTENTE:10", r.Diagnosticos);
            Assert.Equal(14, Assert.Single(r.RelacionesAusentes).IdAula);
            Assert.Equal("PLAN_NO_VERIFICABLE", r.EstadoOrigen);
            Assert.DoesNotContain("NO_COPIAR", r.Relaciones.Single().OrigenCanonico);
            Sql("UPDATE vinculaciones SET fecha_inicio='OTRO_PAYLOAD_AJENO' WHERE id=2");
            var repetida = (await _repo.ObtenerAsync(m, CancellationToken.None)).Value;
            Assert.Equal(r.HuellaOrigen, repetida.HuellaOrigen);
        }

        [Fact]
        public async Task TerritorialInconsistenteNoAcreditaOrigen()
        {
            Sql("UPDATE territoriales SET subunidad='CAUCA' WHERE id=16");
            var r = await _repo.ObtenerAsync(AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto()), CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Equal("PLAN_NO_VERIFICABLE", r.Value.EstadoOrigen);
            Assert.Contains(r.Value.Diagnosticos, d => d.StartsWith("TERRITORIAL_NO_COINCIDE:", StringComparison.Ordinal));
            Assert.Single(r.Value.Relaciones.SelectMany(v => v.Sesiones));
        }

        [Fact]
        public async Task FalloTecnicoSeDistingueDeLecturaVacia()
        {
            Sql("ALTER TABLE sesiones RENAME TO sesiones_no_disponibles");
            var r = await _repo.ObtenerAsync(AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto()), CancellationToken.None);
            Assert.True(r.IsFailure); Assert.StartsWith("LECTURA_PLAN_ERROR:", r.Error);
        }

        // Decisión humana 2026-11-15 «B: extras como ocupación fija»: una relación productiva declarada
        // como extra no requerida deja de bloquear la preparación; se conserva inmutable.
        [Fact]
        public async Task MateriaSoloProductivaSeDeclaraExtraSinEscribirNiSumarDemanda()
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            Sql(@"INSERT INTO vinculaciones VALUES (4,1,1,3,8,93,'I','A','Presencial',32,'VALLE','','');
INSERT INTO sesiones VALUES (20,4,93,'2027-02-06','18:00 A 22:00',4,'VALLE',0);");
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess, r.IsFailure ? r.Error : "");
            Assert.Equal("REQUIERE_RATIFICACION", r.Value.EstadoOrigen);
            Assert.Empty(r.Value.Diagnosticos);
            var extra = Assert.Single(r.Value.RelacionesExtra);
            Assert.Equal(4, extra.IdVinculacion);
            Assert.Equal(8, Assert.Single(r.Value.Manifiesto.RelacionesExtra).IdAsignatura);
            Assert.NotEqual(m.Huella, r.Value.Manifiesto.Huella);
            Assert.Equal(m.Relaciones.Sum(x => x.HorasEfectivas), r.Value.Manifiesto.Relaciones.Sum(x => x.HorasEfectivas));
            Assert.Empty(r.Value.RelacionesAusentes);
            using (var db = new SQLiteConnection(_cs))
            {
                db.Open();
                Assert.Equal(3, db.ExecuteScalar<int>("SELECT COUNT(*) FROM vinculaciones"));
                Assert.Equal(2, db.ExecuteScalar<int>("SELECT COUNT(*) FROM sesiones"));
            }
        }

        [Fact]
        public async Task ExtraDeclaradaSeConservaSinBloquearElPlan()
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.ContextoConExtra());
            Sql(@"INSERT INTO vinculaciones VALUES (4,1,1,3,8,93,'I','A','Presencial',32,'VALLE','','');
INSERT INTO sesiones VALUES (20,4,93,'2027-02-06','18:00 A 22:00',4,'VALLE',0);");
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Equal("REQUIERE_RATIFICACION", r.Value.EstadoOrigen);
            Assert.Empty(r.Value.Diagnosticos);
            Assert.Empty(r.Value.RelacionesAusentes);
            var extra = Assert.Single(r.Value.RelacionesExtra);
            Assert.Equal(4, extra.IdVinculacion);
            Assert.Equal(8, Assert.Single(extra.ClavesCandidatas).IdAsignatura);
            Assert.Equal(20, Assert.Single(extra.Sesiones).IdSesion);
            Assert.DoesNotContain(r.Value.Relaciones, v => v.IdVinculacion == 4);
        }

        [Fact]
        public async Task ExtraDeclaradaNoOcultaAmbiguedadNiRelacionRequeridaAusente()
        {
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.ContextoConExtra());
            Sql(@"INSERT INTO vinculaciones VALUES (4,1,1,3,8,93,'I','A','Presencial',32,'VALLE','','');
INSERT INTO ag_aula VALUES (15,1,1,3,'I','A','Presencial','VALLE');
UPDATE vinculaciones SET subunidad='CAUCA' WHERE id=2;");
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Single(r.Value.RelacionesExtra);
            Assert.Contains(r.Value.Diagnosticos, d => d == "IDENTIDAD_AMBIGUA:1");
            Assert.Equal(14, Assert.Single(r.Value.RelacionesAusentes).IdAula);
            Assert.Equal("PLAN_NO_VERIFICABLE", r.Value.EstadoOrigen);
        }

        // Escenario productivo 2026-11-16 (aulas 12/13/14): el período seleccionado no tiene plan cargado.
        // El lector no debe inventar relaciones ni ocultar la ausencia.
        [Fact]
        public async Task SinFilasProductivasEnElPeriodo_MarcaTodasRequeridasComoAusentesYNoVerificable()
        {
            Sql("DELETE FROM sesiones; DELETE FROM vinculaciones;");
            var m = AgManifiestoCoberturaTests.Crear(AgManifiestoCoberturaTests.Contexto());
            var r = await _repo.ObtenerAsync(m, CancellationToken.None);
            Assert.True(r.IsSuccess);
            Assert.Empty(r.Value.Relaciones);
            Assert.Empty(r.Value.RelacionesExtra);
            Assert.Empty(r.Value.Diagnosticos);
            Assert.Equal(2, r.Value.RelacionesAusentes.Count);
            Assert.All(r.Value.RelacionesAusentes,
                k => Assert.Contains(k.Valor, m.Relaciones.Select(x => x.Clave.Valor)));
            Assert.Equal("PLAN_NO_VERIFICABLE", r.Value.EstadoOrigen);
        }

        public void Dispose() { Directory.Delete(_dir, true); }
    }
}
