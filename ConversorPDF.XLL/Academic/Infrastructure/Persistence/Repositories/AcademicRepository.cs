using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Globalization;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Sync;
using Dapper;

using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
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
        private readonly IDbConnection _connection;
        private readonly IDbTransaction _transaction;
        private readonly string _subunidad;
        private readonly bool _suppressMarkDirty;
        private readonly AgModoDatos _modo;

        public AcademicRepository(IDbConnection connection, IDbTransaction transaction = null, string subunidad = null, bool suppressMarkDirty = false, AgModoDatos modo = null)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
            _transaction = transaction;
            _subunidad = subunidad ?? SyncContext.ActiveSubunidad ?? "";
            _suppressMarkDirty = suppressMarkDirty;
            _modo = modo ?? AgModoDatos.Determinista;
            if (string.IsNullOrEmpty(_subunidad))
            {
                // Fase 3: Omitir advertencia en entorno de pruebas (tests)
                bool isTestMode = AppDomain.CurrentDomain.FriendlyName.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!isTestMode)
                {
                    LoggerHelper.Warning("[AcademicRepository] Subunidad no resuelta. Los cambios no se sincronizarán correctamente.");
                }
            }
        }

        private void EnsureCrSqliteEnabled()
        {
            var rawConn = _connection;
            if (rawConn is ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context.NonDisposableDbConnection ndc)
            {
                rawConn = ndc.Unwrap();
            }

            if (rawConn is System.Data.SQLite.SQLiteConnection sqliteConn && sqliteConn.State == ConnectionState.Open)
            {
                if (!CrSqliteLoader.EnableOnConnection(sqliteConn))
                {
                    LoggerHelper.WarningLimitado("crSqliteNoHabilitado", "[AcademicRepository] CR-SQLite no pudo habilitarse en la conexión activa. Los cambios no se sincronizarán.", maxVeces: 1);
                }
            }
        }

        private void NotifyMarkDirty()
        {
            if (!_suppressMarkDirty)
            {
                ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
            }
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = "INSERT INTO programas (nombreprograma, abreviatura, pensum, actual, activo, modalidad) VALUES (@NombrePrograma, @Abreviatura, @Pensum, @Actual, @Activo, @Modalidad); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, programa, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"INSERT INTO asignaturas (idprograma, semestre, num_nivel, cod_unidad, cod_materia, uni_teorica, metodologia, abreviatura, nucleo, nombre, creditos, int_horaria, tipoasignatura, tipocredito, prerequisito)
                               VALUES (@IdPrograma, @Semestre, @NumNivel, @CodUnidad, @CodMateria, @UniTeorica, @Metodologia, @Abreviatura, @Nucleo, @Nombre, @Creditos, @IntHoraria, @TipoAsignatura, @TipoCredito, @Prerequisito);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, asignatura, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = "INSERT INTO microcurriculo (idasignatura, nucleo, objetivogeneral, competencias) VALUES (@IdAsignatura, @Nucleo, @ObjetivoGeneral, @Competencias); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, microcurriculo, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = "INSERT INTO contenidos (idmicrocurriculo, unidad, contenido) VALUES (@IdMicrocurriculo, @Unidad, @ContenidoTexto); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, contenido, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Terceros ---

        public SqliteResult<IEnumerable<Tercero>> GetTerceros(bool soloActivos = true)
        {
            try
            {
                string sql = @"SELECT id, departamento, nombre, cedula, tipo, salario, categoria as Categoria, pensionado as Pensionado, 
                                      entidad_bancaria as EntidadBancaria, cuenta_bancaria as CuentaBancaria, 
                                      tipo_cuenta_bancaria as TipoCuentaBancaria, telefono, correo, 
                                      correo_personal as CorreoPersonal, fecha_nacimiento as FechaDeNacimiento, 
                                      contrato, residencia as Residencia, id_residencia_divipola as IdResidenciaDivipola, activo as Activo, puntaje_ocasionales as PuntajeOcasionales,
                                      subunidad as Subunidad
                               FROM terceros";
                
                if (soloActivos) sql += " WHERE activo = 1";
                sql += " ORDER BY nombre";

                var result = _connection.Query<Tercero>(sql, transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetTerceros(soloActivos={soloActivos}): {result.Count} terceros encontrados.");
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
                string sql = @"SELECT id, departamento, nombre, cedula, tipo, salario, categoria as Categoria, pensionado as Pensionado, 
                                      entidad_bancaria as EntidadBancaria, cuenta_bancaria as CuentaBancaria, 
                                      tipo_cuenta_bancaria as TipoCuentaBancaria, telefono, correo, 
                                      correo_personal as CorreoPersonal, fecha_nacimiento as FechaDeNacimiento, 
                                      contrato, residencia as Residencia, id_residencia_divipola as IdResidenciaDivipola, activo as Activo, puntaje_ocasionales as PuntajeOcasionales,
                                      subunidad as Subunidad
                               FROM terceros WHERE cedula = @cedula";
                var result = _connection.QueryFirstOrDefault<Tercero>(sql, new { cedula }, transaction: _transaction);
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"INSERT INTO terceros (departamento, nombre, cedula, tipo, salario, categoria, pensionado, entidad_bancaria, cuenta_bancaria, tipo_cuenta_bancaria, telefono, correo, correo_personal, fecha_nacimiento, contrato, residencia, id_residencia_divipola, activo, subunidad, puntaje_ocasionales)
                               VALUES (@Departamento, @Nombre, @Cedula, @Tipo, @Salario, @Categoria, @Pensionado, @EntidadBancaria, @CuentaBancaria, @TipoCuentaBancaria, @Telefono, @Correo, @CorreoPersonal, @FechaDeNacimiento, @Contrato, @Residencia, @IdResidenciaDivipola, @Activo, @Subunidad, @PuntajeOcasionales);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, new
                {
                    tercero.Departamento,
                    tercero.Nombre,
                    tercero.Cedula,
                    tercero.Tipo,
                    tercero.Salario,
                    tercero.Categoria,
                    tercero.Pensionado,
                    tercero.EntidadBancaria,
                    tercero.CuentaBancaria,
                    tercero.TipoCuentaBancaria,
                    tercero.Telefono,
                    tercero.Correo,
                    tercero.CorreoPersonal,
                    tercero.FechaDeNacimiento,
                    tercero.Contrato,
                    tercero.Residencia,
                    tercero.IdResidenciaDivipola,
                    tercero.Activo,
                    tercero.PuntajeOcasionales,
                    Subunidad = tercero.Subunidad ?? _subunidad
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdateTercero(Tercero tercero)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"UPDATE terceros SET 
                    departamento = @Departamento, nombre = @Nombre, cedula = @Cedula,
                    tipo = @Tipo, salario = @Salario, categoria = @Categoria, pensionado = @Pensionado,
                    entidad_bancaria = @EntidadBancaria, cuenta_bancaria = @CuentaBancaria,
                    tipo_cuenta_bancaria = @TipoCuentaBancaria,
                    telefono = @Telefono, correo = @Correo, correo_personal = @CorreoPersonal,
                    fecha_nacimiento = @FechaDeNacimiento, contrato = @Contrato,
                    residencia = @Residencia, id_residencia_divipola = @IdResidenciaDivipola, activo = @Activo, puntaje_ocasionales = @PuntajeOcasionales, subunidad = @Subunidad
                    WHERE id = @Id";
                
                // FASE 18: Crear un objeto anónimo con fallback para Subunidad igual que en InsertTercero
                int filas = _connection.Execute(sql, new
                {
                    tercero.Departamento,
                    tercero.Nombre,
                    tercero.Cedula,
                    tercero.Tipo,
                    tercero.Salario,
                    tercero.Categoria,
                    tercero.Pensionado,
                    tercero.EntidadBancaria,
                    tercero.CuentaBancaria,
                    tercero.TipoCuentaBancaria,
                    tercero.Telefono,
                    tercero.Correo,
                    tercero.CorreoPersonal,
                    tercero.FechaDeNacimiento,
                    tercero.Contrato,
                    tercero.Residencia,
                    tercero.IdResidenciaDivipola,
                    tercero.Activo,
                    tercero.PuntajeOcasionales,
                    tercero.Id,
                    Subunidad = tercero.Subunidad ?? _subunidad
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex) { return SqliteResult<int>.Failure(ex.Message); }
        }

        public SqliteResult<int> DeleteTercero(int id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                var tercero = _connection.QueryFirstOrDefault<Tercero>(
                    "SELECT id, departamento, subunidad as Subunidad, activo as Activo FROM terceros WHERE id = @id",
                    new { id }, transaction: _transaction);
                
                if (tercero == null)
                    return SqliteResult<int>.Failure("Tercero no encontrado");

                var subunidades = (tercero.Subunidad ?? "")
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                // FASE 16: Si no tiene subunidades registradas (legado), asumir que pertenece al local.
                // Esto evita que Delete en una sede afecte un tercero sin sede.
                bool perteneceActual = subunidades.Count == 0 || 
                    subunidades.Any(s => string.Equals(s, _subunidad, StringComparison.OrdinalIgnoreCase));

                if (!perteneceActual)
                    return SqliteResult<int>.Failure($"El tercero no pertenece a la subunidad '{_subunidad}'.");

                subunidades.RemoveAll(s => string.Equals(s, _subunidad, StringComparison.OrdinalIgnoreCase));

                if (subunidades.Count == 0)
                {
                    // FASE 24: Actualizar departamento para que coincida con subunidad
                    int filas = _connection.Execute(
                        "UPDATE terceros SET subunidad = NULL, departamento = NULL, activo = 0 WHERE id = @id",
                        new { id }, transaction: _transaction);
                    NotifyMarkDirty();
                    return SqliteResult<int>.Success(filas);
                }
                else
                {
                    string nuevaSubunidad = string.Join(",", subunidades);
                    // FASE 24: Actualizar departamento para que coincida con subunidad
                    int filas = _connection.Execute(
                        "UPDATE terceros SET subunidad = @Subunidad, departamento = @Subunidad WHERE id = @id",
                        new { id, Subunidad = nuevaSubunidad }, transaction: _transaction);
                    NotifyMarkDirty();
                    return SqliteResult<int>.Success(filas);
                }
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
                var result = _connection.Query<Sede>("SELECT id as Id, iddivipola as IdDivipola, departamento as Departamento, ciudad as Ciudad, departamentoSiif as DepartamentoSiif, ciudadSiif as CiudadSiif, subsede as Subsede, categoria as Categoria, transporte as Transporte, activo as Activo, subunidad as Subunidad FROM sedes", transaction: _transaction).ToList();
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
                var result = _connection.QueryFirstOrDefault<Sede>("SELECT id as Id, iddivipola as IdDivipola, departamento as Departamento, ciudad as Ciudad, departamentoSiif as DepartamentoSiif, ciudadSiif as CiudadSiif, subsede as Subsede, categoria as Categoria, transporte as Transporte, activo as Activo, subunidad as Subunidad FROM sedes WHERE id = @id", new { id }, transaction: _transaction);
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
            EnsureCrSqliteEnabled();
            try
            {
                // Cambio Fase 26: Agregados departamentoSiif y ciudadSiif
                string sql = "INSERT INTO sedes (iddivipola, departamento, ciudad, departamentoSiif, ciudadSiif, subsede, categoria, transporte, activo, subunidad) VALUES (@IdDivipola, @Departamento, @Ciudad, @DepartamentoSiif, @CiudadSiif, @Subsede, @Categoria, @Transporte, @Activo, @Subunidad); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, new
                {
                    sede.IdDivipola,
                    sede.Departamento,
                    sede.Ciudad,
                    sede.DepartamentoSiif,
                    sede.CiudadSiif,
                    sede.Subsede,
                    sede.Categoria,
                    sede.Transporte,
                    sede.Activo,
                    Subunidad = !string.IsNullOrEmpty(sede.Subunidad) ? sede.Subunidad : _subunidad
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
        public SqliteResult<int> UpdateSede(Sede sede)
        {
            EnsureCrSqliteEnabled();
            try
            {
                // Cambio Fase 26: Agregados departamentoSiif y ciudadSiif
                string sql = @"UPDATE sedes SET 
                    iddivipola = @IdDivipola, departamento = @Departamento, ciudad = @Ciudad, departamentoSiif = @DepartamentoSiif, ciudadSiif = @CiudadSiif, subsede = @Subsede,
                    categoria = @Categoria, transporte = @Transporte, activo = @Activo, subunidad = @Subunidad WHERE id = @Id";
                int filas = _connection.Execute(sql, new
                {
                    sede.IdDivipola,
                    sede.Departamento,
                    sede.Ciudad,
                    sede.DepartamentoSiif,
                    sede.CiudadSiif,
                    sede.Subsede,
                    sede.Categoria,
                    sede.Transporte,
                    sede.Activo,
                    Subunidad = !string.IsNullOrEmpty(sede.Subunidad) ? sede.Subunidad : _subunidad,
                    sede.Id
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeleteSede(int id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int filas = _connection.Execute(
                    "DELETE FROM sedes WHERE id = @id AND (subunidad = @Subunidad OR subunidad = '' OR subunidad IS NULL)",
                    new { id, Subunidad = _subunidad }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Divipola ---

        public SqliteResult<IEnumerable<Divipola>> GetDivipolas()
        {
            try
            {
                var result = _connection.Query<Divipola>("SELECT id as Id, departamento as Departamento, departamento_codigo as DepartamentoCodigo, departamento_normalizado as DepartamentoNormalizado, municipio as Municipio, municipio_codigo as MunicipioCodigo, municipio_normalizado as MunicipioNormalizado, municipio_categoria as MunicipioCategoria FROM divipola ORDER BY departamento, municipio", transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Divipola>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Divipola>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Divipola> GetDivipolaById(int id)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Divipola>("SELECT id as Id, departamento as Departamento, departamento_codigo as DepartamentoCodigo, departamento_normalizado as DepartamentoNormalizado, municipio as Municipio, municipio_codigo as MunicipioCodigo, municipio_normalizado as MunicipioNormalizado, municipio_categoria as MunicipioCategoria FROM divipola WHERE id = @id", new { id }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Divipola>.Success(result)
                    : SqliteResult<Divipola>.Failure("Registro Divipola no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Divipola>.Failure(ex.Message);
            }
        }

        public SqliteResult<IEnumerable<Divipola>> SearchDivipola(string searchTerm)
        {
            if (searchTerm == null)
            {
                return SqliteResult<IEnumerable<Divipola>>.Success(System.Linq.Enumerable.Empty<Divipola>());
            }

            try
            {
                string escapedTerm = searchTerm
                    .Replace("\\", "\\\\")
                    .Replace("%", "\\%")
                    .Replace("_", "\\_");
                string query = $"%{escapedTerm}%";
                
                string sql = "SELECT id as Id, departamento as Departamento, departamento_codigo as DepartamentoCodigo, departamento_normalizado as DepartamentoNormalizado, municipio as Municipio, municipio_codigo as MunicipioCodigo, municipio_normalizado as MunicipioNormalizado, municipio_categoria as MunicipioCategoria FROM divipola WHERE municipio LIKE @query ESCAPE '\\' OR departamento LIKE @query ESCAPE '\\' ORDER BY departamento, municipio";
                var result = _connection.Query<Divipola>(sql, new { query }, transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Divipola>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Divipola>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertDivipola(Divipola divipola)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"INSERT INTO divipola (departamento, departamento_codigo, departamento_normalizado, municipio, municipio_codigo, municipio_normalizado, municipio_categoria) 
                               VALUES (@Departamento, @DepartamentoCodigo, @DepartamentoNormalizado, @Municipio, @MunicipioCodigo, @MunicipioNormalizado, @MunicipioCategoria); SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, divipola, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdateDivipola(Divipola divipola)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"UPDATE divipola SET 
                    departamento = @Departamento, departamento_codigo = @DepartamentoCodigo, departamento_normalizado = @DepartamentoNormalizado, 
                    municipio = @Municipio, municipio_codigo = @MunicipioCodigo, municipio_normalizado = @MunicipioNormalizado, municipio_categoria = @MunicipioCategoria 
                    WHERE id = @Id";
                int filas = _connection.Execute(sql, divipola, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeleteDivipola(int id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                // Simple validation to check if Sede is using it (can be extended)
                int count = _connection.QuerySingle<int>("SELECT COUNT(1) FROM sedes WHERE iddivipola = @id", new { id }, transaction: _transaction);
                if (count > 0)
                {
                    return SqliteResult<int>.Failure("No se puede eliminar el registro porque está siendo referenciado por una Sede.");
                }

                int filas = _connection.Execute("DELETE FROM divipola WHERE id = @id", new { id }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        // --- Periodos ---

        private bool PeriodosTieneVentana()
        {
            try
            {
                var cols = _connection.Query<string>("SELECT name FROM pragma_table_info('periodos');", transaction: _transaction)
                                      .Select(x => x.ToLowerInvariant())
                                      .ToList();
                return cols.Contains("periodo_inicio") && cols.Contains("periodo_fin");
            }
            catch
            {
                return false;
            }
        }

        public SqliteResult<IEnumerable<Periodo>> GetPeriodos()
        {
            try
            {
                string colIni = PeriodosTieneVentana() ? "periodo_inicio" : "NULL";
                string colFin = PeriodosTieneVentana() ? "periodo_fin" : "NULL";
                var result = _connection.Query<Periodo>($"SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, {colIni} as PeriodoInicio, {colFin} as PeriodoFin, activo as Activo FROM periodos", transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Periodo>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Periodo>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Periodo> GetPeriodoById(int id)
        {
            try
            {
                string colIni = PeriodosTieneVentana() ? "periodo_inicio" : "NULL";
                string colFin = PeriodosTieneVentana() ? "periodo_fin" : "NULL";
                var result = _connection.QueryFirstOrDefault<Periodo>($"SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, {colIni} as PeriodoInicio, {colFin} as PeriodoFin, activo as Activo FROM periodos WHERE idperiodo = @id", new { id }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Periodo>.Success(result)
                    : SqliteResult<Periodo>.Failure("Período no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Periodo>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertPeriodo(Periodo periodo)
        {
            if (periodo == null) return SqliteResult<int>.Failure("El período no puede ser nulo.");
            EnsureCrSqliteEnabled();
            try
            {
                bool tieneVentana = PeriodosTieneVentana();
                if (!tieneVentana && (!string.IsNullOrWhiteSpace(periodo.PeriodoInicio)
                                      || !string.IsNullOrWhiteSpace(periodo.PeriodoFin)))
                {
                    return SqliteResult<int>.Failure(
                        "El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones.");
                }

                string sql = tieneVentana
                    ? "INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo) VALUES (@IdPeriodo, @PeriodoNombre, @PeriodoInicio, @PeriodoFin, @Activo); SELECT last_insert_rowid();"
                    : "INSERT INTO periodos (idperiodo, periodo, activo) VALUES (@IdPeriodo, @PeriodoNombre, @Activo); SELECT last_insert_rowid();";

                int id = _connection.QuerySingle<int>(sql, new
                {
                    periodo.IdPeriodo,
                    periodo.PeriodoNombre,
                    periodo.PeriodoInicio,
                    periodo.PeriodoFin,
                    periodo.Activo
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdatePeriodo(Periodo periodo)
        {
            if (periodo == null) return SqliteResult<int>.Failure("El período no puede ser nulo.");
            EnsureCrSqliteEnabled();
            try
            {
                bool tieneVentana = PeriodosTieneVentana();
                if (!tieneVentana && (!string.IsNullOrWhiteSpace(periodo.PeriodoInicio)
                                      || !string.IsNullOrWhiteSpace(periodo.PeriodoFin)))
                {
                    return SqliteResult<int>.Failure(
                        "El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones.");
                }

                string sql = tieneVentana
                    ? @"UPDATE periodos SET 
                            periodo = @PeriodoNombre, 
                            periodo_inicio = @PeriodoInicio, 
                            periodo_fin = @PeriodoFin, 
                            activo = @Activo 
                        WHERE idperiodo = @IdPeriodo;"
                    : @"UPDATE periodos SET 
                            periodo = @PeriodoNombre, 
                            activo = @Activo 
                        WHERE idperiodo = @IdPeriodo;";

                int filas = _connection.Execute(sql, new
                {
                    periodo.IdPeriodo,
                    periodo.PeriodoNombre,
                    periodo.PeriodoInicio,
                    periodo.PeriodoFin,
                    periodo.Activo
                }, transaction: _transaction);

                NotifyMarkDirty();
                return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeletePeriodo(int id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                // Validación RN3: Integridad referencial en vinculaciones
                int countVinc = _connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(1) FROM vinculaciones WHERE idperiodo = @id;", new { id }, transaction: _transaction);
                if (countVinc > 0)
                {
                    return SqliteResult<int>.Failure("No se puede eliminar el período porque tiene vinculaciones académicas asociadas.");
                }

                // Validación RN3: Integridad referencial en festivos (Fase 204)
                int countFest = _connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(1) FROM festivos WHERE idperiodo = @id;", new { id }, transaction: _transaction);
                if (countFest > 0)
                {
                    return SqliteResult<int>.Failure(
                        "No se puede eliminar el período porque tiene días festivos configurados.");
                }

                // Validación RN3: Integridad referencial en config (Fase 205)
                bool configTieneIdPeriodo = _connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(1) FROM pragma_table_info('config') WHERE name = 'id_periodo';",
                    transaction: _transaction) > 0;

                if (configTieneIdPeriodo)
                {
                    int countConf = _connection.QueryFirstOrDefault<int>(
                        "SELECT COUNT(1) FROM config WHERE id_periodo = @id;", new { id }, transaction: _transaction);
                    if (countConf > 0)
                    {
                        return SqliteResult<int>.Failure(
                            "No se puede eliminar el período porque está configurado como período activo en el sistema.");
                    }
                }
                else
                {
                    LoggerHelper.Warning("[AcademicRepository] DeletePeriodo: 'config.id_periodo' ausente; guarda omitida.");
                }

                int filas = _connection.Execute(
                    "DELETE FROM periodos WHERE idperiodo = @id;", new { id }, transaction: _transaction);

                NotifyMarkDirty();
                return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<IEnumerable<Periodo>> GetPeriodosConVentana()
        {
            return GetPeriodos();
        }

        // --- Subunidades ---
        
        public SqliteResult<IEnumerable<string>> GetSubunidades()
        {
            try
            {
                var result = _connection.Query<string>("SELECT DISTINCT subunidad FROM territoriales WHERE subunidad IS NOT NULL AND subunidad <> '' ORDER BY subunidad", transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<string>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<string>>.Failure(ex.Message);
            }
        }

        // --- Terminales ---

        public SqliteResult<IEnumerable<Terminal>> GetTerminales()
        {
            try
            {
                var result = _connection.Query<Terminal>("SELECT id as Id, terminal as TerminalName, maxvalor as MaxValor FROM terminales WHERE id > 1 ORDER BY id", transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Terminal>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Terminal>>.Failure(ex.Message);
            }
        }

        // --- Vinculaciones ---

        public SqliteResult<IEnumerable<Vinculacion>> GetVinculaciones()
        {
            try
            {
                string sql = $"SELECT id, idsede as IdSede, idperiodo as IdPeriodo, idtercero as IdTercero, idprograma as IdPrograma, grupo, semestre, idasignatura as IdAsignatura, total_horas as TotalHoras, dia_horas_semanales as DiaHorasSemanales, fecha_inicio as FechaInicio, fecha_fin as FechaFin, modalidadVinculacion as Modalidad, isPernoctaUltimoDia as IsPernoctaUltimoDia, isTerminalAereo as IsTerminalAereo, idTerminales as IdTerminales, gastoTerminalAerea as GastoTerminalAerea FROM {_modo.TablaVinculaciones} WHERE (subunidad = @Subunidad OR subunidad = '' OR subunidad IS NULL)";
                var result = _connection.Query<Vinculacion>(
                    sql,
                    new { Subunidad = _subunidad }, transaction: _transaction);
                return SqliteResult<IEnumerable<Vinculacion>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Vinculacion>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertVinculacion(Vinculacion vinculacion)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = $@"INSERT INTO {_modo.TablaVinculaciones} (idsede, idperiodo, idtercero, idprograma, grupo, semestre, idasignatura, total_horas, dia_horas_semanales, fecha_inicio, fecha_fin, modalidadVinculacion, subunidad, isPernoctaUltimoDia, isTerminalAereo, idTerminales, gastoTerminalAerea, es_vacante, motivo_vacante)
                               VALUES (@IdSede, @IdPeriodo, @IdTercero, @IdPrograma, @Grupo, @Semestre, @IdAsignatura, @TotalHoras, @DiaHorasSemanales, @FechaInicio, @FechaFin, @Modalidad, @Subunidad, @IsPernoctaUltimoDia, @IsTerminalAereo, @IdTerminales, @GastoTerminalAerea, @EsVacante, @MotivoVacante);
                               SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, new
                {
                    vinculacion.IdSede,
                    vinculacion.IdPeriodo,
                    vinculacion.IdTercero,
                    vinculacion.IdPrograma,
                    vinculacion.Grupo,
                    vinculacion.Semestre,
                    vinculacion.IdAsignatura,
                    vinculacion.TotalHoras,
                    vinculacion.DiaHorasSemanales,
                    vinculacion.FechaInicio,
                    vinculacion.FechaFin,
                    vinculacion.Modalidad,
                    Subunidad = _subunidad,
                    vinculacion.IsPernoctaUltimoDia,
                    vinculacion.IsTerminalAereo,
                    vinculacion.IdTerminales,
                    vinculacion.GastoTerminalAerea,
                    vinculacion.EsVacante,
                    MotivoVacante = vinculacion.MotivoVacante ?? string.Empty
                }, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] InsertVinculacion: Vinculación insertada con Id={id}, subunidad={_subunidad}");
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
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
                var sql = $@"
                    SELECT id, idsede as IdSede, idperiodo as IdPeriodo, idtercero as IdTercero, idprograma as IdPrograma, grupo, semestre, idasignatura as IdAsignatura,
                             total_horas as TotalHoras, dia_horas_semanales as DiaHorasSemanales,
                             fecha_inicio as FechaInicio, fecha_fin as FechaFin, modalidadVinculacion as Modalidad,
                             isPernoctaUltimoDia as IsPernoctaUltimoDia,
                             isTerminalAereo as IsTerminalAereo, idTerminales as IdTerminales, gastoTerminalAerea as GastoTerminalAerea,
                             es_vacante as EsVacante, motivo_vacante as MotivoVacante
                      FROM {_modo.TablaVinculaciones} WHERE id = @id";
                
                var row = _connection.QueryFirstOrDefault<dynamic>(sql, new { id }, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] GetVinculacionById({id}): {(row != null ? "encontrada" : "no encontrada")}");
                if (row == null)
                    return SqliteResult<Vinculacion>.Failure("No se encontró la vinculación.");

                var dict = (IDictionary<string, object>)row;
                var isPernoctaUltimoDia = dict.ContainsKey("IsPernoctaUltimoDia") ? dict["IsPernoctaUltimoDia"] : null;
                var isTerminalAereo = dict.ContainsKey("IsTerminalAereo") ? dict["IsTerminalAereo"] : null;
                var idTerminales = dict.ContainsKey("IdTerminales") ? dict["IdTerminales"] : null;
                var gastoTerminalAerea = dict.ContainsKey("GastoTerminalAerea") ? dict["GastoTerminalAerea"] : null;

                var vinc = new Vinculacion
                {
                    Id = Convert.ToInt32(row.id),
                    IdTercero = Convert.ToInt32(row.IdTercero),
                    IdPeriodo = Convert.ToInt32(row.IdPeriodo),
                    IdSede = Convert.ToInt32(row.IdSede),
                    IdPrograma = Convert.ToInt32(row.IdPrograma),
                    Semestre = Convert.ToString(row.semestre),
                    IdAsignatura = Convert.ToInt32(row.IdAsignatura),
                    Grupo = Convert.ToString(row.grupo),
                    Modalidad = Convert.ToString(row.Modalidad),
                    TotalHoras = Convert.ToInt32(row.TotalHoras),
                    DiaHorasSemanales = Convert.ToString(row.DiaHorasSemanales),
                    FechaInicio = ParseDateSafe(row.FechaInicio),
                    FechaFin = ParseDateSafe(row.FechaFin),
                    IsPernoctaUltimoDia = isPernoctaUltimoDia != null && isPernoctaUltimoDia != DBNull.Value ? Convert.ToInt32(isPernoctaUltimoDia) : 0,
                    IsTerminalAereo = isTerminalAereo != null && isTerminalAereo != DBNull.Value ? Convert.ToInt32(isTerminalAereo) : 0,
                    IdTerminales = idTerminales != null && idTerminales != DBNull.Value ? Convert.ToInt32(idTerminales) : 0,
                    GastoTerminalAerea = gastoTerminalAerea != null && gastoTerminalAerea != DBNull.Value
                        ? (double.TryParse(Convert.ToString(gastoTerminalAerea), NumberStyles.Any, CultureInfo.InvariantCulture, out double g) ? g : 0)
                        : 0,
                    EsVacante = dict.ContainsKey("EsVacante") && dict["EsVacante"] != null && dict["EsVacante"] != DBNull.Value ? Convert.ToInt32(dict["EsVacante"]) : 0,
                    MotivoVacante = dict.ContainsKey("MotivoVacante") && dict["MotivoVacante"] != null && dict["MotivoVacante"] != DBNull.Value ? Convert.ToString(dict["MotivoVacante"]) : string.Empty
                };
                return SqliteResult<Vinculacion>.Success(vinc);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en GetVinculacionById({id})", ex);
                return SqliteResult<Vinculacion>.Failure(ex.Message);
            }
        }

        private static DateTime? ParseDateSafe(object value)
        {
            if (value == null || value == DBNull.Value) return null;
            string str = Convert.ToString(value);
            if (string.IsNullOrWhiteSpace(str) || str == "None") return null;
            if (DateTime.TryParse(str, out DateTime result)) return result;
            return null;
        }

        public SqliteResult<int> UpdateVinculacion(Vinculacion vinculacion)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = $@"UPDATE {_modo.TablaVinculaciones} SET
                                   idsede = @IdSede,
                                   idperiodo = @IdPeriodo,
                                   idtercero = @IdTercero,
                                   idprograma = @IdPrograma,
                                   grupo = @Grupo,
                                   semestre = @Semestre,
                                   idasignatura = @IdAsignatura,
                                   total_horas = @TotalHoras,
                                   dia_horas_semanales = @DiaHorasSemanales,
                                   fecha_inicio = @FechaInicio,
                                   fecha_fin = @FechaFin,
                                   modalidadVinculacion = @Modalidad,
                                   isPernoctaUltimoDia = @IsPernoctaUltimoDia,
                                   isTerminalAereo = @IsTerminalAereo,
                                   idTerminales = @IdTerminales,
                                   gastoTerminalAerea = @GastoTerminalAerea,
                                   es_vacante = @EsVacante,
                                   motivo_vacante = @MotivoVacante
                               WHERE id = @Id";
                int rows = _connection.Execute(sql, new
                {
                    vinculacion.IdSede,
                    vinculacion.IdPeriodo,
                    vinculacion.IdTercero,
                    vinculacion.IdPrograma,
                    vinculacion.Grupo,
                    vinculacion.Semestre,
                    vinculacion.IdAsignatura,
                    vinculacion.TotalHoras,
                    vinculacion.DiaHorasSemanales,
                    vinculacion.FechaInicio,
                    vinculacion.FechaFin,
                    vinculacion.Modalidad,
                    vinculacion.IsPernoctaUltimoDia,
                    vinculacion.IsTerminalAereo,
                    vinculacion.IdTerminales,
                    vinculacion.GastoTerminalAerea,
                    vinculacion.EsVacante,
                    MotivoVacante = vinculacion.MotivoVacante ?? string.Empty,
                    vinculacion.Id
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en UpdateVinculacion Id={vinculacion.Id}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeleteVinculacion(int idVinculacion)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int filas = _connection.Execute(
                    $"DELETE FROM {_modo.TablaVinculaciones} WHERE id = @id",
                    new { id = idVinculacion }, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] DeleteVinculacion: Vinculación eliminada con Id={idVinculacion}, subunidad={_subunidad}");
                NotifyMarkDirty(); 
                return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en DeleteVinculacion Id={idVinculacion}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> SetPernoctaUltimoDia(int idVinculacion, int valor)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int rows = _connection.Execute(
                    $"UPDATE {_modo.TablaVinculaciones} SET isPernoctaUltimoDia = @valor WHERE id = @id",
                    new { id = idVinculacion, valor = valor == 1 ? 1 : 0 }, transaction: _transaction);
                if (rows > 0)
                {
                    NotifyMarkDirty();
                }
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en SetPernoctaUltimoDia Id={idVinculacion}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }


        public SqliteResult<IEnumerable<Sesion>> GetSesionesByVinculacion(int idVinculacion)
        {
            try
            {
                var result = _connection.Query<Sesion>(
                    // Fases 11-12 FIX: COALESCE en viatica para evitar que NULLs de PULL/Migración se mapeen a false
                    $"SELECT id, idvinculacion AS IdVinculacion, idtercero AS IdTercero, fechasesion AS FechaSesion, horariosesion AS HorarioSesion, horas, COALESCE(viatica, 0) AS Viatica FROM {_modo.TablaSesiones} WHERE idvinculacion = @idVinculacion",
                    new { idVinculacion }, transaction: _transaction).ToList();
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
            EnsureCrSqliteEnabled();
            try
            {
                string sql = $@"INSERT INTO {_modo.TablaSesiones} (idvinculacion, idtercero, fechasesion, horariosesion, horas, viatica, subunidad, es_vacante, motivo_vacante)
                               VALUES (@IdVinculacion, @IdTercero, @FechaSesion, @HorarioSesion, @Horas, @Viatica, @Subunidad, @EsVacante, @MotivoVacante);
                               SELECT last_insert_rowid();"; 
                int id = _connection.QuerySingle<int>(sql, new
                {
                    sesion.IdVinculacion,
                    sesion.IdTercero,
                    sesion.FechaSesion,
                    sesion.HorarioSesion,
                    sesion.Horas,
                    Viatica = sesion.Viatica ? 1 : 0, // Fase 8 FIX: Mapeo estricto a INTEGER
                    Subunidad = _subunidad,
                    sesion.EsVacante,
                    MotivoVacante = sesion.MotivoVacante ?? string.Empty
                }, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] InsertSesion: Sesión guardada con ID {id}.");
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en InsertSesion", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertSesiones(IEnumerable<Sesion> sesiones)
        {
            EnsureCrSqliteEnabled();
            if (sesiones == null || !sesiones.Any()) { NotifyMarkDirty(); return SqliteResult<int>.Success(0); }

            try
            {
                string sql = $@"INSERT INTO {_modo.TablaSesiones} (idvinculacion, idtercero, fechasesion, horariosesion, horas, viatica, subunidad, es_vacante, motivo_vacante)
                                     VALUES (@IdVinculacion, @IdTercero, @FechaSesion, @HorarioSesion, @Horas, @Viatica, @Subunidad, @EsVacante, @MotivoVacante);";

                // Dapper ejecuta batch insert cuando recibe IEnumerable como parámetro
                var parametros = sesiones.Select(s => new {
                    s.IdVinculacion,
                    s.IdTercero,
                    s.FechaSesion,
                    s.HorarioSesion,
                    s.Horas,
                    Viatica = s.Viatica ? 1 : 0, // Fase 8 FIX: Mapeo estricto a INTEGER
                    Subunidad = _subunidad,
                    s.EsVacante,
                    MotivoVacante = s.MotivoVacante ?? string.Empty
                });
                int count = _connection.Execute(sql, parametros, transaction: _transaction);
                
                LoggerHelper.Info($"[AcademicRepository] InsertSesiones: {count} sesiones insertadas vía batch (Dapper).");
                NotifyMarkDirty(); return SqliteResult<int>.Success(count);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en InsertSesiones (Batch)", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
        public SqliteResult<int> DeleteSesionesByVinculacion(int idVinculacion)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int rows = _connection.Execute(
                    $"DELETE FROM {_modo.TablaSesiones} WHERE idvinculacion = @idVinculacion AND (subunidad = @Subunidad OR subunidad = '' OR subunidad IS NULL)",
                    new { idVinculacion, Subunidad = _subunidad }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] Error en DeleteSesionesByVinculacion para VincID: {idVinculacion}", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> ExecuteRawSql(string sql)
        {
            EnsureCrSqliteEnabled();
            if (string.IsNullOrWhiteSpace(sql)) { NotifyMarkDirty(); return SqliteResult<int>.Success(0); }

            try
            {
                int rows = _connection.Execute(sql, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] ExecuteRawSql: {rows} filas afectadas.");
                NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en ExecuteRawSql", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> SeedCatalogosTransaccional()
        {
            EnsureCrSqliteEnabled();
            try
            {
                string resourceName = "ConversorPDF.XLL.Academic.Infrastructure.Persistence.Sql.AcademicSeedData.sql";
                var assembly = Assembly.GetExecutingAssembly();
                
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        LoggerHelper.Warning($"[AcademicRepository] No se encontró el recurso embebido: {resourceName}");
                        NotifyMarkDirty(); return SqliteResult<int>.Success(0);
                    }

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string sql = reader.ReadToEnd();
                        
                        int rows = _connection.Execute(sql, transaction: _transaction);
                        LoggerHelper.Info($"[AcademicRepository] SeedCatalogosTransaccional: {rows} filas afectadas vía Recurso Embebido (Idempotente).");
                        
                        // Reparar subunidad en registros locales insertados por el seed
                        RepairSubunidadAfterSeed();
                        
                        NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AcademicRepository] Error en SeedCatalogosTransaccional", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        /// <summary>
        /// Repara registros huérfanos sin subunidad estableciéndolos a la subunidad actual,
        /// pero SOLO si no existen ya en la tabla registros que pertenezcan a OTRA subunidad.
        /// 
        /// Lógica del NOT EXISTS: Es conservadora intencionalmente. Si por ejemplo, tras un PULL,
        /// recibimos datos de "BOGOTA" en la base local de "VALLE", el NOT EXISTS fallará y no 
        /// repararemos los registros vacíos. Esto previene que datos sin subunidad sean reclamados
        /// erróneamente cuando hay bases de datos mixtas. Asumimos que los registros sin subunidad
        /// son del "Seed" inicial antes de la mezcla.
        /// </summary>
        private void RepairSubunidadAfterSeed()
        {
            try
            {
                if (string.IsNullOrEmpty(_subunidad)) return;

                // FASE 20: Incluir cdpDependencias en las tablas a reparar tras un seed.
                var localTables = new[] { "sedes", "festivos", "vinculaciones", "sesiones", "salvedades_conflictos", "cdpDependencias" };
                foreach (var table in localTables)
                {
                    int needsRepair = _connection.ExecuteScalar<int>($@"
                        SELECT COUNT(1) FROM [{table}] 
                        WHERE (subunidad = '' OR subunidad IS NULL)
                          AND NOT EXISTS (SELECT 1 FROM [{table}] WHERE subunidad != '' AND subunidad != @Subunidad)",
                        new { Subunidad = _subunidad }, transaction: _transaction);

                    if (needsRepair > 0)
                    {
                        int repaired = _connection.Execute($@"
                            UPDATE [{table}] SET subunidad = @Subunidad
                            WHERE (subunidad = '' OR subunidad IS NULL)
                              AND NOT EXISTS (SELECT 1 FROM [{table}] WHERE subunidad != '' AND subunidad != @Subunidad)",
                            new { Subunidad = _subunidad }, transaction: _transaction);
                        if (repaired > 0)
                        {
                            LoggerHelper.Info($"[AcademicRepository] SeedCatalogosTransaccional: reparados {repaired} registros en '{table}' con subunidad.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AcademicRepository] No se pudo reparar subunidad tras seed: {ex.Message}");
            }
        }

        // --- Festivos ---

        public SqliteResult<IEnumerable<Festivo>> GetFestivos()
        {
            try
            {
                var result = _connection.Query<Festivo>(@"
                    SELECT idfestivo as IdFestivo, idperiodo as IdPeriodo, 
                           fechafestivoinicio as FechaFestivoInicio, 
                           fechafestivofinal as FechaFestivoFinal, 
                           detallefestivo as DetalleFestivo,
                           tipofestivo as TipoFestivo
                    FROM festivos", transaction: _transaction).ToList();
                LoggerHelper.Info($"[AcademicRepository] GetFestivos: {result.Count} festivos encontrados.");
                return SqliteResult<IEnumerable<Festivo>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Festivo>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Festivo> GetFestivoById(int id)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Festivo>(@"
                    SELECT idfestivo as IdFestivo, idperiodo as IdPeriodo, 
                           fechafestivoinicio as FechaFestivoInicio, 
                           fechafestivofinal as FechaFestivoFinal, 
                           detallefestivo as DetalleFestivo,
                           tipofestivo as TipoFestivo
                    FROM festivos WHERE idfestivo = @id", new { id }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Festivo>.Success(result)
                    : SqliteResult<Festivo>.Failure("Festivo no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Festivo>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> InsertFestivo(Festivo festivo)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"
                    INSERT INTO festivos (idperiodo, fechafestivoinicio, fechafestivofinal, detallefestivo, tipofestivo, subunidad)
                    VALUES (@IdPeriodo, @FechaFestivoInicio, @FechaFestivoFinal, @DetalleFestivo, @TipoFestivo, @Subunidad);
                    SELECT last_insert_rowid();";
                int id = _connection.QuerySingle<int>(sql, new
                {
                    festivo.IdPeriodo,
                    festivo.FechaFestivoInicio,
                    festivo.FechaFestivoFinal,
                    festivo.DetalleFestivo,
                    festivo.TipoFestivo,
                    Subunidad = _subunidad
                }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(id);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdateFestivo(Festivo festivo)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"
                    UPDATE festivos SET 
                        idperiodo = @IdPeriodo, 
                        fechafestivoinicio = @FechaFestivoInicio, 
                        fechafestivofinal = @FechaFestivoFinal, 
                        detallefestivo = @DetalleFestivo,
                        tipofestivo = @TipoFestivo
                    WHERE idfestivo = @IdFestivo";
                int rows = _connection.Execute(sql, festivo, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeleteFestivo(int id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int rows = _connection.Execute(
                    "DELETE FROM festivos WHERE idfestivo = @id AND (subunidad = @Subunidad OR subunidad = '' OR subunidad IS NULL)",
                    new { id, Subunidad = _subunidad }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
        // --- Usuarios ---

        public SqliteResult<IEnumerable<Usuario>> GetUsuarios()
        {
            try
            {
                var result = _connection.Query<Usuario>("SELECT id as Id, nombre as Nombre, nombreusuario as NombreUsuario, subunidad as Subunidad, activo as Activo, perfil as Perfil, firma as Firma FROM usuarios", transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Usuario>>.Success(result);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Usuario>>.Failure(ex.Message);
            }
        }

        public SqliteResult<Usuario> GetUsuarioByNombreUsuario(string nombreUsuario)
        {
            try
            {
                var result = _connection.QueryFirstOrDefault<Usuario>("SELECT id as Id, nombre as Nombre, nombreusuario as NombreUsuario, subunidad as Subunidad, activo as Activo, perfil as Perfil, firma as Firma FROM usuarios WHERE nombreusuario = @nombreUsuario COLLATE NOCASE", new { nombreUsuario }, transaction: _transaction);
                return result != null
                    ? SqliteResult<Usuario>.Success(result)
                    : SqliteResult<Usuario>.Failure("Usuario no encontrado");
            }
            catch (Exception ex)
            {
                return SqliteResult<Usuario>.Failure(ex.Message);
            }
        }

        public SqliteResult<string> InsertUsuario(Usuario usuario)
        {
            EnsureCrSqliteEnabled();
            try
            {
                // Validación: Verificar que el usuario no exista previamente para evitar duplicidad o fallas de restricción
                string checkSql = "SELECT COUNT(1) FROM usuarios WHERE nombreusuario = @NombreUsuario COLLATE NOCASE";
                int count = _connection.ExecuteScalar<int>(checkSql, new { NombreUsuario = usuario.NombreUsuario }, transaction: _transaction);
                if (count > 0)
                {
                    return SqliteResult<string>.Failure($"El usuario '{usuario.NombreUsuario}' ya se encuentra registrado.");
                }

                usuario.Id = Guid.NewGuid().ToString("N").ToLowerInvariant();
                string sql = @"INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, perfil, firma) 
                               VALUES (@Id, @Nombre, @NombreUsuario, @Subunidad, @Activo, @Perfil, @Firma);";
                _connection.Execute(sql, usuario, transaction: _transaction);
                
                NotifyMarkDirty(); 
                return SqliteResult<string>.Success(usuario.Id);
            }
            catch (Exception ex)
            {
                return SqliteResult<string>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> UpdateUsuario(Usuario usuario)
        {
            EnsureCrSqliteEnabled();
            try
            {
                LoggerHelper.Info($"[AcademicRepository] UpdateUsuario START - Id={usuario.Id}, NombreUsuario={usuario.NombreUsuario}, FirmaLength={usuario.Firma?.Length ?? 0}");
                string sql = @"UPDATE usuarios SET 
                    nombre = @Nombre, nombreusuario = @NombreUsuario, subunidad = @Subunidad, activo = @Activo, perfil = @Perfil, firma = @Firma 
                    WHERE id = @Id";
                int filas = _connection.Execute(sql, usuario, transaction: _transaction);
                LoggerHelper.Info($"[AcademicRepository] UpdateUsuario Execute Result: Filas={filas}");
                NotifyMarkDirty(); 
                return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AcademicRepository] UpdateUsuario Error", ex);
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeleteUsuario(string id)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int filas = _connection.Execute("DELETE FROM usuarios WHERE id = @id", new { id }, transaction: _transaction);
                NotifyMarkDirty(); return SqliteResult<int>.Success(filas);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<string> GetPermisosPorUsuario(string nombreUsuario)
        {
            try
            {
                string sql = @"
                    SELECT p.opcionespermitidas 
                    FROM usuarios u 
                    INNER JOIN permisos p ON u.perfil = p.perfil 
                    WHERE u.nombreusuario = @nombreUsuario COLLATE NOCASE";

                string permisos = _connection.QueryFirstOrDefault<string>(sql, new { nombreUsuario }, transaction: _transaction);
                
                return SqliteResult<string>.Success(permisos);
            }
            catch (Exception ex)
            {
                return SqliteResult<string>.Failure(ex.Message);
            }
        }

        // CRUD Permisos
        public SqliteResult<IEnumerable<Permiso>> GetTodosLosPermisos()
        {
            try
            {
                string sql = "SELECT perfil as Perfil, opcionespermitidas as Botones FROM permisos ORDER BY perfil";
                var list = _connection.Query<Permiso>(sql, transaction: _transaction).ToList();
                return SqliteResult<IEnumerable<Permiso>>.Success(list);
            }
            catch (Exception ex)
            {
                return SqliteResult<IEnumerable<Permiso>>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> SavePermiso(Permiso permiso)
        {
            EnsureCrSqliteEnabled();
            try
            {
                string sql = @"
                    INSERT INTO permisos (perfil, opcionespermitidas) 
                    VALUES (@Perfil, @Botones)
                    ON CONFLICT(perfil) DO UPDATE SET 
                        opcionespermitidas = excluded.opcionespermitidas";

                int rows = _connection.Execute(sql, permiso, transaction: _transaction);
                if (rows > 0)
                    NotifyMarkDirty();
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }

        public SqliteResult<int> DeletePermiso(string perfil)
        {
            EnsureCrSqliteEnabled();
            try
            {
                // Prevent deleting 'Admin' or 'Usuario' to avoid system lockout? Let's just delete
                string sql = "DELETE FROM permisos WHERE perfil = @Perfil";
                int rows = _connection.Execute(sql, new { Perfil = perfil }, transaction: _transaction);
                if (rows > 0)
                    NotifyMarkDirty();
                return SqliteResult<int>.Success(rows);
            }
            catch (Exception ex)
            {
                return SqliteResult<int>.Failure(ex.Message);
            }
        }
    }
}

