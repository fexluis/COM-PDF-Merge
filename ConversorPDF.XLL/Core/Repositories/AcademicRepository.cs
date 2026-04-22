using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using ConversorPDF.XLL.Core.Data;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Models.Academic;
using Dapper;

namespace ConversorPDF.XLL.Core.Repositories
{
    /// <summary>
    /// Implementación del repositorio académico para SQLite usando Dapper y App-Level Encryption (Vault).
    ///
    /// PATRÓN CORRECTO DE CONTEXTO:
    ///   using (var context = SqliteHelper.GetContext()) {
    ///       var db = context.Open();   // Sin 'using' — el contexto gestiona la conexión
    ///       db.Query(...);
    ///   }
    ///   Los métodos de escritura requieren llamar a ReEncryptAndPersist() explícitamente 
    ///   fuera del hilo UI para asegurar la persistencia en el Vault maestro.
    /// </summary>
    public class AcademicRepository : IAcademicRepository
    {
        private readonly System.Data.SQLite.SQLiteConnection _connection;
        private readonly System.Data.SQLite.SQLiteTransaction _transaction;

        public AcademicRepository(System.Data.SQLite.SQLiteConnection connection, System.Data.SQLite.SQLiteTransaction transaction = null)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
            _transaction = transaction;
        }

        public SqliteResult<IEnumerable<Programa>> GetProgramas()
        {
            try
            {
                var result = _connection.Query<Programa>("SELECT * FROM programas", transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetProgramas: {result.Count} programas encontrados.");
                return SqliteResult<IEnumerable<Programa>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Programa>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Programa> GetProgramaById(int id)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Programa>("SELECT * FROM programas WHERE id = @id", new { id }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Programa>.Success(result)
                    : SqliteResult<Programa>.Failure("Programa no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Programa>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertPrograma(Programa programa)
        {
            try
            {
                string sql = "INSERT INTO programas (nombreprograma, abreviatura, tipo) VALUES (@NombrePrograma, @Abreviatura, @Tipo); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, programa, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<IEnumerable<Asignatura>> GetAsignaturasByPrograma(int idPrograma, string semestre = null)
        {
            try
            {
                string sql = "SELECT * FROM asignaturas WHERE idprograma = @idPrograma";
                if (!string.IsNullOrEmpty(semestre)) sql += " AND semestre = @semestre";

                var result = _connection.Query<Asignatura>(sql, new { idPrograma, semestre }, transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetAsignaturasByPrograma(id={idPrograma}, sem={semestre ?? "ALL"}): {result.Count} asignaturas encontradas.");
                return SqliteResult<IEnumerable<Asignatura>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Asignatura>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Asignatura> GetAsignaturaByAbreviatura(string abreviatura)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Asignatura>("SELECT * FROM asignaturas WHERE abreviatura = @abreviatura", new { abreviatura }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Asignatura>.Success(result)
                    : SqliteResult<Asignatura>.Failure("Asignatura no encontrada");
            }
            catch (Exception ex)
            {
                return SqliteResult<Asignatura>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertAsignatura(Asignatura asignatura)
        {
            try
            {
                string sql = @"INSERT INTO asignaturas (idprograma, semestre, código, abreviatura, área, nombreasignatura, creditos, tipoasignatura, tipocredito, prerequisito)
                               VALUES (@IdPrograma, @Semestre, @Código, @Abreviatura, @Área, @NombreAsignatura, @Creditos, @TipoAsignatura, @TipoCredito, @Prerequisito);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, asignatura, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<Microcurriculo> GetMicrocurriculoByAsignatura(int idAsignatura)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Microcurriculo>("SELECT * FROM microcurriculo WHERE idasignatura = @idAsignatura", new { idAsignatura }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Microcurriculo>.Success(result)
                    : SqliteResult<Microcurriculo>.Failure("Microcurrículo no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Microcurriculo>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertMicrocurriculo(Microcurriculo microcurriculo)
        {
            try
            {
                string sql = "INSERT INTO microcurriculo (idasignatura, area, objetivogeneral, competencias) VALUES (@IdAsignatura, @Area, @ObjetivoGeneral, @Competencias); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, microcurriculo, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<IEnumerable<Contenido>> GetContenidosByMicrocurriculo(int idMicrocurriculo)
        {
            try
            {
                var result = _connection.Query<Contenido>("SELECT id, idmicrocurriculo, unidad, contenido as ContenidoTexto FROM contenidos WHERE idmicrocurriculo = @idMicrocurriculo", new { idMicrocurriculo }, transaction: _transaction);
                return SqliteResult<IEnumerable<Contenido>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Contenido>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertContenido(Contenido contenido)
        {
            try
            {
                string sql = "INSERT INTO contenidos (idmicrocurriculo, unidad, contenido) VALUES (@IdMicrocurriculo, @Unidad, @ContenidoTexto); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, contenido, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Terceros ---

        public SqliteResult<IEnumerable<Tercero>> GetTerceros()
        {
            try
            {
                var result = _connection.Query<Tercero>("SELECT id, departamento, nombre, cedula, tipo, salario, rol, entidad_bancaria as EntidadBancaria, cuenta_bancaria as CuentaBancaria, telefono, correo, fecha_nacimiento as FechaDeNacimiento, contrato FROM terceros", transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetTerceros: {result.Count} terceros encontrados.");
                return SqliteResult<IEnumerable<Tercero>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Tercero>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Tercero> GetTerceroByCedula(string cedula)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Tercero>("SELECT id, departamento, nombre, cedula, tipo, salario, rol, entidad_bancaria as EntidadBancaria, cuenta_bancaria as CuentaBancaria, telefono, correo, fecha_nacimiento as FechaDeNacimiento, contrato FROM terceros WHERE cedula = @cedula", new { cedula }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Tercero>.Success(result)
                    : SqliteResult<Tercero>.Failure("Tercero no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Tercero>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertTercero(Tercero tercero)
        {
            try
            {
                string sql = @"INSERT INTO terceros (departamento, nombre, cedula, tipo, salario, rol, entidad_bancaria, cuenta_bancaria, telefono, correo, fecha_nacimiento, contrato)
                               VALUES (@Departamento, @Nombre, @Cedula, @Tipo, @Salario, @Rol, @EntidadBancaria, @CuentaBancaria, @Telefono, @Correo, @FechaDeNacimiento, @Contrato);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, tercero, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Sedes ---

        public SqliteResult<IEnumerable<Sede>> GetSedes()
        {
            try
            {
                var result = _connection.Query<Sede>("SELECT * FROM sedes", transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetSedes: {result.Count} sedes encontradas.");
                return SqliteResult<IEnumerable<Sede>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Sede>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Sede> GetSedeById(int id)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Sede>("SELECT * FROM sedes WHERE id = @id", new { id }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Sede>.Success(result)
                    : SqliteResult<Sede>.Failure("Sede no encontrada");
            }
            catch (Exception ex)
            {
                return SqliteResult<Sede>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertSede(Sede sede)
        {
            try
            {
                string sql = "INSERT INTO sedes (departamento, ciudad, categoria, transporte) VALUES (@Departamento, @Ciudad, @Categoria, @Transporte); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, sede, transaction: _transaction);
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Vinculaciones ---

        public SqliteResult<IEnumerable<Vinculacion>> GetVinculaciones()
        {
            try
            {
                var result = _connection.Query<Vinculacion>("SELECT id, idsede as IdSede, idtercero as IdTercero, idprograma as IdPrograma, semestre, idasignatura as IdAsignatura, total_horas as TotalHoras, dia_horas_semanales as DiaHorasSemanales, fecha_inicio as FechaInicio, fecha_fin as FechaFin FROM vinculaciones", transaction: _transaction);
                return SqliteResult<IEnumerable<Vinculacion>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Vinculacion>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertVinculacion(Vinculacion vinculacion)
        {
            try
            {
                string sql = @"INSERT INTO vinculaciones (idsede, idtercero, idprograma, semestre, idasignatura, total_horas, dia_horas_semanales, fecha_inicio, fecha_fin)
                               VALUES (@IdSede, @IdTercero, @IdPrograma, @Semestre, @IdAsignatura, @TotalHoras, @DiaHorasSemanales, @FechaInicio, @FechaFin);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, vinculacion, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] InsertVinculacion: Vinculación insertada con Id={id}");
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en InsertVinculacion", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<Vinculacion> GetVinculacionById(int id)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Vinculacion>(
                    @"SELECT id, idsede as IdSede, idtercero as IdTercero, idprograma as IdPrograma, semestre, idasignatura as IdAsignatura,
                             total_horas as TotalHoras, dia_horas_semanales as DiaHorasSemanales,
                             fecha_inicio as FechaInicio, fecha_fin as FechaFin
                      FROM vinculaciones WHERE id = @id",
                    new { id }, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] GetVinculacionById({id}): {(result != null ? "encontrada" : "no encontrada")}");
                return result != null
                    ? SqliteResult<Vinculacion>.Success(result)
                    : SqliteResult<Vinculacion>.Failure("No se encontró la vinculación.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en GetVinculacionById({id})", ex);
                return SqliteResult<Vinculacion>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdateVinculacion(Vinculacion vinculacion)
        {
            try
            {
                string sql = @"UPDATE vinculaciones SET
                                   idsede = @IdSede,
                                   idtercero = @IdTercero,
                                   idprograma = @IdPrograma,
                                   semestre = @Semestre,
                                   idasignatura = @IdAsignatura,
                                   total_horas = @TotalHoras,
                                   dia_horas_semanales = @DiaHorasSemanales,
                                   fecha_inicio = @FechaInicio,
                                   fecha_fin = @FechaFin
                               WHERE id = @Id";
                int rows = _connection.Execute(sql, vinculacion, transaction: _transaction);
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en UpdateVinculacion Id={vinculacion.Id}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }


        public SqliteResult<IEnumerable<Sesion>> GetSesionesByVinculacion(int idVinculacion)
        {
            try
            {
                var result = _connection.Query<Sesion>("SELECT id, idvinculacion as IdVinculacion, idtercero as IdTercero, fechasesionInicial, fechasesionFinal, duracion, horas, viatica FROM sesiones WHERE idvinculacion = @idVinculacion", new { idVinculacion }, transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetSesionesByVinculacion(id={idVinculacion}): {result.Count} sesiones encontradas.");
                return SqliteResult<IEnumerable<Sesion>>.Success(result);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en GetSesionesByVinculacion", ex);
                return SqliteResult<IEnumerable<Sesion>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertSesion(Sesion sesion)
        {
            try
            {
                string sql = @"INSERT INTO sesiones (idvinculacion, idtercero, fechasesionInicial, fechasesionFinal, duracion, horas, viatica)
                               VALUES (@IdVinculacion, @IdTercero, @fechasesionInicial, @fechasesionFinal, @Duracion, @Horas, @Viatica);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, sesion, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] InsertSesion: Sesión guardada con ID {id}.");
                return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en InsertSesion", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertSesiones(IEnumerable<Sesion> sesiones)
        {
            if (sesiones == null || !sesiones.Any()) return SqliteResult<int>.Success(0);

            try
            {
                const string sql = @"INSERT INTO sesiones (idvinculacion, idtercero, fechasesionInicial, fechasesionFinal, duracion, horas, viatica)
                                     VALUES (@IdVinculacion, @IdTercero, @fechasesionInicial, @fechasesionFinal, @Duracion, @Horas, @Viatica);";

                // Dapper ejecuta batch insert cuando recibe IEnumerable como parámetro
                int count = _connection.Execute(sql, sesiones, transaction: _transaction);
                
                LoggerHelper.Info($"[AcademicRepository] InsertSesiones: {count} sesiones insertadas vía batch (Dapper).");
                return SqliteResult<int>.Success(count);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en InsertSesiones (Batch)", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
        public SqliteResult<int> DeleteSesionesByVinculacion(int idVinculacion)
        {
            try
            {
                int rows = _connection.Execute("DELETE FROM sesiones WHERE idvinculacion = @idVinculacion", new { idVinculacion }, transaction: _transaction);
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en DeleteSesionesByVinculacion para VincID: {idVinculacion}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> ExecuteRawSql(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return SqliteResult<int>.Success(0);

            try
            {
                int rows = _connection.Execute(sql, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] ExecuteRawSql: {rows} filas afectadas.");
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en ExecuteRawSql", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> SeedCatalogosTransaccional()
        {
            try
            {
                string resourceName = "ConversorPDF.XLL.Core.Data.AcademicSeedData.sql";
                var assembly = Assembly.GetExecutingAssembly();
                
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        LoggerHelper.Warning($"[AcademicRepository] No se encontró el recurso embebido: {resourceName}");
                        return SqliteResult<int>.Success(0);
                    }

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string sql = reader.ReadToEnd();
                        
                        // Asegurar idempotencia reemplazando INSERT por INSERT OR IGNORE 
                        // si el recurso embebido no lo tiene aún por defecto.
                        if (!sql.Contains("OR IGNORE"))
                        {
                            sql = sql.Replace("INSERT INTO", "INSERT OR IGNORE INTO");
                        }

                        int rows = _connection.Execute(sql, transaction: _transaction);
                        LoggerHelper.Info($"[AcademicRepository] SeedCatalogosTransaccional: {rows} filas afectadas vía Recurso Embebido (Idempotente).");
                        return SqliteResult<int>.Success(rows);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en SeedCatalogosTransaccional", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
    }
}

