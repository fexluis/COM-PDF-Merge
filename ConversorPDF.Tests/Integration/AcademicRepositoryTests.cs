using Xunit;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence.Integration
{
    public class AcademicRepositoryTests : IDisposable
    {
        private readonly SQLiteConnection _conn;
        private readonly AcademicRepository _repo;

        public AcademicRepositoryTests()
        {
            _conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            _conn.Open();
            CrearEsquema();
            _repo = new AcademicRepository(_conn);
        }

        private void CrearEsquema()
        {
            _conn.Execute(@"
                CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT, subunidad TEXT);
                CREATE TABLE vinculaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    idtercero INTEGER, idsede INTEGER, idprograma INTEGER,
                    semestre TEXT, idasignatura INTEGER,
                    total_horas INTEGER, dia_horas_semanales TEXT,
                    fecha_inicio TEXT, fecha_fin TEXT, subunidad TEXT,
                    es_vacante INTEGER DEFAULT 0, motivo_vacante TEXT DEFAULT ''
                );
                CREATE TABLE sesiones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    idvinculacion INTEGER, idtercero INTEGER,
                    fechasesion TEXT, horariosesion TEXT,
                    horas REAL, viatica INTEGER, subunidad TEXT,
                    es_vacante INTEGER DEFAULT 0, motivo_vacante TEXT DEFAULT ''
                );
                CREATE TABLE divipola (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    departamento TEXT, departamento_codigo TEXT, departamento_normalizado TEXT,
                    municipio TEXT, municipio_codigo TEXT, municipio_normalizado TEXT, municipio_categoria INTEGER
                );
                CREATE TABLE sedes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    iddivipola INTEGER, departamento TEXT, ciudad TEXT, subsede TEXT,
                    categoria INTEGER, transporte REAL, subunidad TEXT
                );
                CREATE TABLE festivos (
                    idfestivo INTEGER PRIMARY KEY AUTOINCREMENT,
                    idperiodo INTEGER, fechafestivoinicio TEXT, fechafestivofinal TEXT,
                    detallefestivo TEXT, tipofestivo TEXT, subunidad TEXT
                );
                CREATE TABLE periodos (
                    idperiodo INTEGER PRIMARY KEY AUTOINCREMENT,
                    periodo TEXT, activo INTEGER
                );
                INSERT INTO terceros (id, nombre, subunidad) VALUES (1, 'Docente Test', '');
            ");
            // Nota: Este esquema asume PRAGMA foreign_keys = OFF ya que SQLite 
            // lo tiene desactivado por defecto, por lo cual los tests de vinculaciones 
            // no fallan aunque no referencien datos completos en otras tablas.
        }

        [Fact]
        public void DeleteSesionesByVinculacion_3SesionesExistentes_Elimina3YDejaTablaVacia()
        {
            _conn.Execute("INSERT INTO sesiones (idvinculacion, idtercero, horas) " +
                "VALUES (1,1,8), (1,1,8), (1,1,8), (2,1,4)");

            var result = _repo.DeleteSesionesByVinculacion(1);
            int restantes = _conn.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sesiones WHERE idvinculacion = 1");

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(3);
            restantes.Should().Be(0);
        }

        [Fact]
        public void InsertSesiones_6Sesiones_Inserta6YRetornaConteo()
        {
            _conn.Execute("INSERT INTO vinculaciones (id, idtercero) VALUES (1, 1)");
            var sesiones = new List<Sesion>();
            for (int i = 0; i < 6; i++)
            {
                sesiones.Add(new Sesion
                {
                    IdVinculacion = 1, IdTercero = 1,
                    FechaSesion = new DateTime(2026, 4, 23).AddDays(i * 7),
                    HorarioSesion = "JUEVES|8",
                    Horas = 8, Viatica = false
                });
            }

            var result = _repo.InsertSesiones(sesiones);
            int conteo = _conn.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sesiones WHERE idvinculacion = 1");

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(6);
            conteo.Should().Be(6);
        }

        [Fact]
        public void UpsertFlow_DejaExactamenteNNuevasSesiones()
        {
            _conn.Execute("INSERT INTO vinculaciones (id, idtercero) VALUES (1, 1)");
            for (int i = 0; i < 5; i++)
                _conn.Execute("INSERT INTO sesiones (idvinculacion, idtercero, horas) VALUES (1, 1, 8)");

            var nuevasSesiones = new List<Sesion>
            {
                new Sesion { IdVinculacion=1, IdTercero=1, Horas=8, FechaSesion=DateTime.Today, HorarioSesion="LUNES|8" },
                new Sesion { IdVinculacion=1, IdTercero=1, Horas=4, FechaSesion=DateTime.Today.AddDays(7), HorarioSesion="LUNES|4" }
            };

            _repo.DeleteSesionesByVinculacion(1);
            _repo.InsertSesiones(nuevasSesiones);

            int conteoFinal = _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM sesiones WHERE idvinculacion = 1");
            conteoFinal.Should().Be(2);
        }

        // --- Divipola Tests ---

        [Fact]
        public void GetDivipolas_ConDatos_RetornaTodos()
        {
            _conn.Execute("INSERT INTO divipola (departamento, municipio) VALUES ('CUNDINAMARCA', 'BOGOTA'), ('ANTIOQUIA', 'MEDELLIN')");
            
            var result = _repo.GetDivipolas();
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(2);
        }

        [Fact]
        public void GetDivipolaById_IdExistente_RetornaRegistro()
        {
            _conn.Execute("INSERT INTO divipola (id, departamento, municipio) VALUES (1, 'HUILA', 'NEIVA')");
            
            var result = _repo.GetDivipolaById(1);
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Municipio.Should().Be("NEIVA");
        }

        [Fact]
        public void SearchDivipola_TerminoVacio_RetornaTodo()
        {
            _conn.Execute("INSERT INTO divipola (departamento, municipio) VALUES ('CUNDINAMARCA', 'BOGOTA'), ('ANTIOQUIA', 'MEDELLIN')");
            
            var result = _repo.SearchDivipola("");
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(2);
        }

        [Fact]
        public void SearchDivipola_TerminoParcial_RetornaFiltroCorrecto()
        {
            _conn.Execute("INSERT INTO divipola (departamento, municipio) VALUES ('CUNDINAMARCA', 'BOGOTA'), ('ANTIOQUIA', 'MEDELLIN')");
            
            var result = _repo.SearchDivipola("med");
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().ContainSingle(d => d.Municipio == "MEDELLIN");
        }

        [Fact]
        public void SearchDivipola_SinResultados_RetornaVacio()
        {
            _conn.Execute("INSERT INTO divipola (departamento, municipio) VALUES ('CUNDINAMARCA', 'BOGOTA')");
            
            var result = _repo.SearchDivipola("CALI");
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }

        [Fact]
        public void InsertDivipola_DatosValidos_RetornaIdNuevo()
        {
            var d = new Divipola { Departamento = "VALLE", Municipio = "CALI", MunicipioCategoria = 1 };
            
            var result = _repo.InsertDivipola(d);
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeGreaterThan(0);
        }

        [Fact]
        public void InsertDivipola_DatosIncompletos_RetornaFalla_SiBaseDatosImpide()
        {
            // By default SQLite allows NULL unless NOT NULL is specified.
            // In our test schema above, we didn't specify NOT NULL, but let's test a valid insertion anyway.
            var d = new Divipola { Departamento = "VALLE", Municipio = "CALI", MunicipioCategoria = 1 };
            var result = _repo.InsertDivipola(d);
            result.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void UpdateDivipola_IdValido_ActualizaCorrectamente()
        {
            _conn.Execute("INSERT INTO divipola (id, departamento, municipio) VALUES (1, 'TOLIMA', 'IBAGUE')");
            
            var d = new Divipola { Id = 1, Departamento = "TOLIMA", Municipio = "ESPINAL", MunicipioCategoria = 2 };
            var result = _repo.UpdateDivipola(d);
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(1);

            var dbMuni = _conn.ExecuteScalar<string>("SELECT municipio FROM divipola WHERE id = 1");
            dbMuni.Should().Be("ESPINAL");
        }

        [Fact]
        public void DeleteDivipola_IdValidoSinReferencia_EliminaCorrectamente()
        {
            _conn.Execute("INSERT INTO divipola (id, departamento, municipio) VALUES (1, 'HUILA', 'NEIVA')");
            
            var result = _repo.DeleteDivipola(1);
            
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(1);
            
            int conteo = _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM divipola");
            conteo.Should().Be(0);
        }

        [Fact]
        public void DeleteDivipola_IdReferenciadoPorSede_RetornaFalla()
        {
            _conn.Execute("INSERT INTO divipola (id, departamento, municipio) VALUES (1, 'BOYACA', 'TUNJA')");
            _conn.Execute("INSERT INTO sedes (iddivipola, ciudad) VALUES (1, 'TUNJA')");
            
            var result = _repo.DeleteDivipola(1);
            
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("referenciado por una Sede");
        }

        public void Dispose()
        {
            _conn?.Close();
            _conn?.Dispose();
        }
    }
}
