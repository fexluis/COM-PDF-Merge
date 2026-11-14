using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
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
    public class AgHuellaEscaladoTests : IDisposable
    {
        private readonly string _archivo = Path.Combine(Path.GetTempPath(), "HuellaEscalado_" + Guid.NewGuid().ToString("N") + ".sqlite");
        private IDbConnection Conexion() { return new SQLiteConnection("Data Source=" + _archivo + ";Version=3;"); }

        public AgHuellaEscaladoTests()
        {
            SQLiteConnection.CreateFile(_archivo);
            using (var db = Conexion())
            {
                db.Open();
                db.Execute(@"
CREATE TABLE ag_aula(id INTEGER PRIMARY KEY,aula_ref TEXT,subunidad TEXT,jornada TEXT,horas_dia INTEGER,dias_ventana TEXT,id_programa INTEGER);
CREATE TABLE ag_aula_asignatura(id_aula INTEGER,id_asignatura INTEGER,horas_objetivo REAL,orden_curricular INTEGER,origen_patron TEXT,dias_patron TEXT);
CREATE TABLE ag_arbol_nodo(id INTEGER PRIMARY KEY,idarbol INTEGER,aula_ref TEXT,subunidad TEXT,jornada TEXT,horas_dia INTEGER,dias_ventana TEXT);
CREATE TABLE ag_arbol(id INTEGER PRIMARY KEY,fecha_fin TEXT);
CREATE TABLE asignaturas(id INTEGER PRIMARY KEY,horas INTEGER);
CREATE TABLE programas(id INTEGER PRIMARY KEY,modalidad TEXT);
CREATE TABLE ag_bloque_horario(id INTEGER PRIMARY KEY,activo INTEGER);
CREATE TABLE festivos(id INTEGER PRIMARY KEY,fecha TEXT);
CREATE TABLE auditoria(n INTEGER);
INSERT INTO auditoria VALUES(0);
CREATE TRIGGER contar_aula AFTER UPDATE ON ag_aula BEGIN UPDATE auditoria SET n=n+1; END;
INSERT INTO ag_arbol VALUES(42,'2027-02-28'),(43,'2027-02-28');
INSERT INTO asignaturas VALUES(101,40),(102,40);
INSERT INTO programas VALUES(1,'PRESENCIAL');
INSERT INTO ag_bloque_horario VALUES(10,1);");
            }
        }

        public void Dispose()
        {
            if (File.Exists(_archivo)) { File.Delete(_archivo); }
        }

        [Fact]
        public async Task Aplicar_WhenSinCambios_ShouldGuardarYReaplicarSinEscrituras()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            decisiones.Should().ContainSingle();
            // Act
            var primera = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            var segunda = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            primera.Exito.Should().BeTrue(primera.Error);
            primera.Modificadas.Should().Be(1);
            segunda.Exito.Should().BeTrue(segunda.Error);
            segunda.Confirmadas.Should().Be(1);
            segunda.Modificadas.Should().Be(0);
            using (var db = Conexion())
            {
                db.Open();
                db.QuerySingle<string>("SELECT dias_ventana FROM ag_aula").Should().Be("LUN-VIE");
                db.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE idarbol=42").Should().Be("LUN-VIE");
                db.QuerySingle<string>("SELECT dias_ventana FROM ag_arbol_nodo WHERE idarbol=43").Should().Be("LUN-JUE");
                db.QuerySingle<int>("SELECT COUNT(*) FROM ag_aula_asignatura WHERE dias_patron<>''").Should().Be(0);
                db.QuerySingle<int>("SELECT n FROM auditoria").Should().Be(1);
            }
        }

        [Theory]
        [InlineData("horas")]
        [InlineData("orden")]
        [InlineData("origen")]
        [InlineData("patron")]
        [InlineData("bloque")]
        [InlineData("festivo")]
        [InlineData("periodo")]
        [InlineData("asignatura")]
        [InlineData("programa")]
        [InlineData("nodo")]
        [InlineData("aula")]
        public async Task Aplicar_WhenContenidoPersistidoCambio_ShouldRechazarSinGuardar(string cambio)
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            decisiones.Should().ContainSingle();
            var sql = new Dictionary<string, string>
            {
                { "horas", "UPDATE ag_aula_asignatura SET horas_objetivo=CASE id_asignatura WHEN 101 THEN 44 ELSE 36 END" },
                { "orden", "UPDATE ag_aula_asignatura SET orden_curricular=99 WHERE id_asignatura=101" },
                { "origen", "UPDATE ag_aula_asignatura SET origen_patron='FIJO' WHERE id_asignatura=101" },
                { "patron", "UPDATE ag_aula_asignatura SET dias_patron='3,4' WHERE id_asignatura=101" },
                { "bloque", "UPDATE ag_bloque_horario SET activo=0" },
                { "festivo", "INSERT INTO festivos VALUES(1,'2027-02-20')" },
                { "periodo", "UPDATE ag_arbol SET fecha_fin='2027-03-01' WHERE id=42" },
                { "asignatura", "UPDATE asignaturas SET horas=44 WHERE id=101" },
                { "programa", "UPDATE programas SET modalidad='POSGRADO'" },
                { "nodo", "UPDATE ag_arbol_nodo SET horas_dia=8 WHERE idarbol=42" },
                { "aula", "UPDATE ag_aula SET jornada='TARDE'" }
            };
            using (var db = Conexion()) { db.Open(); db.Execute(sql[cambio]); }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Error.Should().Contain("Configuración modificada");
            result.Confirmadas.Should().Be(0);
            result.Modificadas.Should().Be(0);
            Ventana().Should().Be("LUN-JUE");
        }

        [Theory]
        [InlineData("ventana")]
        [InlineData("festivo")]
        [InlineData("bloque")]
        [InlineData("horas")]
        public async Task Aplicar_WhenContextoVigenteCambio_ShouldRechazar(string cambio)
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            var vigente = Contexto(false);
            if (cambio == "ventana") { vigente.Ventana.FechaFin = vigente.Ventana.FechaFin.AddDays(1); }
            if (cambio == "festivo") { vigente.Festivos.Add((new DateTime(2027,2,20),new DateTime(2027,2,20))); }
            if (cambio == "bloque") { vigente.Bloques[0].Activo = 0; }
            if (cambio == "horas") { vigente.AulaAsignaturas[0].HorasObjetivo += 4; vigente.AulaAsignaturas[1].HorasObjetivo -= 4; }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, vigente, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
            Ventana().Should().Be("LUN-JUE");
        }

        [Fact]
        public async Task Aplicar_WhenPerfilCambio_ShouldRechazar()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 43, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Error.Should().Contain("perfil");
            Ventana().Should().Be("LUN-JUE");
        }

        [Fact]
        public async Task Aplicar_WhenDecisionAlteradaEnSegundoElemento_ShouldRevertirLoteCompleto()
        {
            // Arrange
            var ctx = Preparar(true);
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            decisiones.Should().HaveCount(2);
            decisiones[1].Hacia = "MIE-DOM";
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
            result.Modificadas.Should().Be(0);
            using (var db = Conexion())
            {
                db.Open();
                db.QuerySingle<int>("SELECT COUNT(*) FROM ag_aula WHERE dias_ventana<>'LUN-JUE'").Should().Be(0);
                db.QuerySingle<int>("SELECT COUNT(*) FROM ag_aula_asignatura WHERE dias_patron=''").Should().Be(0);
                db.QuerySingle<int>("SELECT n FROM auditoria").Should().Be(0);
            }
        }

        [Fact]
        public async Task Aplicar_WhenFallaNodoDespuesDeActualizarAula_ShouldRevertir()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            using (var db = Conexion())
            {
                db.Open();
                db.Execute("CREATE TRIGGER impedir_nodo BEFORE UPDATE ON ag_arbol_nodo BEGIN SELECT RAISE(ABORT,'fallo nodo'); END;");
            }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
            Ventana().Should().Be("LUN-JUE");
        }

        [Fact]
        public async Task Aplicar_WhenPostestadoIncompleto_ShouldNoConfirmarReaplicacion()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            (await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None)).Exito.Should().BeTrue();
            using (var db = Conexion()) { db.Open(); db.Execute("UPDATE ag_arbol_nodo SET dias_ventana='LUN-JUE' WHERE idarbol=42"); }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
        }

        [Fact]
        public async Task Aplicar_WhenCancelado_ShouldNoGuardar()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            using (var cancelacion = new CancellationTokenSource())
            {
                cancelacion.Cancel();
                // Act
                var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, cancelacion.Token);
                // Assert
                result.Exito.Should().BeFalse();
                result.Confirmadas.Should().Be(0);
                Ventana().Should().Be("LUN-JUE");
            }
        }

        [Fact]
        public async Task Aplicar_WhenSinInstantanea_ShouldRechazar()
        {
            // Arrange
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = new[] { new AgEscaladorVentanaAula.Decision
                { AulaRef="A",Desde="LUN-JUE",Hacia="LUN-VIE",CapacidadAntes=16,CapacidadDespues=20 } };
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
        }

        [Fact]
        public async Task Aplicar_WhenReaplicacionConContextoRecargado_ShouldConfirmarSinEscribir()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            (await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None)).Exito.Should().BeTrue();
            var vigente = Contexto(false);
            vigente.Aulas[0].DiasVentana = "LUN-VIE";
            foreach (var aa in vigente.AulaAsignaturas) { aa.DiasPatron = string.Empty; }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, vigente, CancellationToken.None);
            // Assert
            result.Exito.Should().BeTrue(result.Error);
            result.Confirmadas.Should().Be(1);
            result.Modificadas.Should().Be(0);
        }

        [Fact]
        public async Task Aplicar_WhenPatronAutomaticoReaparece_ShouldRechazarReaplicacion()
        {
            // Arrange
            var ctx = Preparar();
            var sut = new AgEscaladorVentanaAula(Conexion);
            var decisiones = sut.Evaluar(ctx, 42);
            (await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None)).Exito.Should().BeTrue();
            using (var db = Conexion()) { db.Open(); db.Execute("UPDATE ag_aula_asignatura SET dias_patron='5' WHERE id_asignatura=101"); }
            // Act
            var result = await sut.AplicarDetalladoAsync(decisiones, 42, ctx, CancellationToken.None);
            // Assert
            result.Exito.Should().BeFalse();
            result.Confirmadas.Should().Be(0);
        }

        private string Ventana()
        {
            using (var db = Conexion()) { db.Open(); return db.QuerySingle<string>("SELECT dias_ventana FROM ag_aula WHERE id=1"); }
        }

        private AgContexto Preparar(bool doble = false)
        {
            var ctx = Contexto(doble);
            using (var db = Conexion())
            {
                db.Open();
                foreach (var aula in ctx.Aulas)
                {
                    db.Execute(@"INSERT INTO ag_aula VALUES(@Id,@AulaRef,@Subunidad,@Jornada,@HorasDia,@DiasVentana,@IdPrograma)", aula);
                    foreach (int perfil in new[] { 42,43 })
                    {
                        db.Execute(@"INSERT INTO ag_arbol_nodo(idarbol,aula_ref,subunidad,jornada,horas_dia,dias_ventana)
VALUES(@Perfil,@AulaRef,@Subunidad,@Jornada,@HorasDia,@DiasVentana)",
                            new { Perfil=perfil,aula.AulaRef,aula.Subunidad,aula.Jornada,aula.HorasDia,aula.DiasVentana });
                    }
                }
                foreach (var aa in ctx.AulaAsignaturas)
                {
                    db.Execute(@"INSERT INTO ag_aula_asignatura VALUES(@IdAula,@IdAsignatura,@HorasObjetivo,@OrdenCurricular,@OrigenPatron,@DiasPatron)", aa);
                }
            }
            return ctx;
        }

        private static AgContexto Contexto(bool doble)
        {
            var ctx = new AgContexto { Ventana = new VentanaSemestre
                { FechaInicio=new DateTime(2027,2,1),FechaFin=new DateTime(2027,2,28) } };
            ctx.Programas.Add(new Programa { Id=1,Modalidad="PRESENCIAL" });
            for (int d=1;d<=7;d++)
            {
                ctx.Bloques.Add(new AgBloqueHorario { Id=d*10,DiaSemanaNumero=d,Franja="MANANA",DuracionHoras=4,Activo=1 });
                if (d==6) { ctx.Bloques.Add(new AgBloqueHorario { Id=61,DiaSemanaNumero=6,Franja="TARDE",DuracionHoras=4,Activo=1 }); }
            }
            ctx.Asignaturas.Add(new Asignatura { Id=101,IdPrograma=1,Semestre="I",Horas=40 });
            ctx.Asignaturas.Add(new Asignatura { Id=102,IdPrograma=1,Semestre="I",Horas=40 });
            for (int id=1;id<=(doble?2:1);id++)
            {
                ctx.Aulas.Add(new AgAula { Id=id,AulaRef="A"+id,Subunidad="CENTRAL",IdPrograma=1,Semestre="I",
                    DiasVentana="LUN-JUE",Jornada="MANANA",HorasDia=4,Estado=AgEstadoAula.Disponible });
                foreach (int asig in new[] { 101,102 })
                {
                    ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula=id,IdAsignatura=asig,HorasObjetivo=40,
                        OrdenCurricular=asig,OrigenPatron="AUTOMATICO",DiasPatron="1,2" });
                }
            }
            ctx.ConstruirIndices();
            return ctx;
        }
    }
}
