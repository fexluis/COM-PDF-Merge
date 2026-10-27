using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using Dapper;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Repositorio para la persistencia transaccional de perfiles de selección del árbol académico y sus nodos asociados (Fase 238).
    /// </summary>
    public class AgArbolRepository : IAgArbolRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgArbolRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        private static void NotifyMarkDirty()
        {
            try { SyncContext.Scheduler?.MarkDirty(); } catch { }
        }

        public async Task<Result<long>> GuardarAsync(AgArbol perfil, IReadOnlyCollection<AgArbolNodo> nodos, CancellationToken ct)
        {
            if (perfil == null) return Result.Failure<long>("El perfil no puede ser nulo.");
            if (string.IsNullOrWhiteSpace(perfil.Nombre)) return Result.Failure<long>("El nombre del perfil es obligatorio.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    using (var tx = db.BeginTransaction())
                    {
                        ct.ThrowIfCancellationRequested();

                        long idArbol = perfil.Id;
                        string ahora = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        if (idArbol <= 0)
                        {
                            // Verificar si ya existe por nombre y subunidad para sobrescritura transaccional (Fase 248)
                            const string sqlExiste = @"
SELECT id FROM ag_arbol 
WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(@Nombre))
  AND COALESCE(subunidad, '') = COALESCE(@Subunidad, '') LIMIT 1;";
                            
                            var idExistente = await db.QueryFirstOrDefaultAsync<long?>(
                                new CommandDefinition(sqlExiste, new { perfil.Nombre, perfil.Subunidad }, transaction: tx, cancellationToken: ct));

                            if (idExistente.HasValue && idExistente.Value > 0)
                            {
                                idArbol = idExistente.Value;
                            }
                        }

                        if (idArbol > 0)
                        {
                            const string sqlUpdate = @"
UPDATE ag_arbol SET
    nombre = @Nombre,
    descripcion = @Descripcion,
    idperiodo = @IdPeriodo,
    fecha_inicio = @FechaInicio,
    fecha_fin = @FechaFin,
    ventana_origen = @VentanaOrigen,
    manual_ventana = @ManualVentana,
    poblacion = @Poblacion,
    generaciones = @Generaciones,
    tasa_cruce = @TasaCruce,
    tasa_mutacion = @TasaMutacion,
    elitismo = @Elitismo,
    estancamiento_max = @EstancamientoMax,
    semilla = @Semilla,
    usuario = @Usuario,
    subunidad = @Subunidad,
    modificado_en = @ModificadoEn,
    activo = @Activo
WHERE id = @Id;";

                            int filas = await db.ExecuteAsync(new CommandDefinition(sqlUpdate, new
                            {
                                Id = idArbol,
                                perfil.Nombre,
                                Descripcion = perfil.Descripcion ?? string.Empty,
                                perfil.IdPeriodo,
                                FechaInicio = perfil.FechaInicio ?? string.Empty,
                                FechaFin = perfil.FechaFin ?? string.Empty,
                                VentanaOrigen = perfil.VentanaOrigen ?? "Periodos",
                                perfil.ManualVentana,
                                perfil.Poblacion,
                                perfil.Generaciones,
                                perfil.TasaCruce,
                                perfil.TasaMutacion,
                                perfil.Elitismo,
                                perfil.EstancamientoMax,
                                perfil.Semilla,
                                Usuario = perfil.Usuario ?? string.Empty,
                                Subunidad = perfil.Subunidad ?? string.Empty,
                                ModificadoEn = ahora,
                                perfil.Activo
                            }, transaction: tx, cancellationToken: ct));

                            // Fase 247: Si el perfil especificado no existe, abortar para no dejar nodos huérfanos
                            if (filas <= 0)
                            {
                                tx.Rollback();
                                return Result.Failure<long>($"No se pudo actualizar el perfil con ID {idArbol}: el registro no existe en la base de datos.");
                            }
                        }
                        else
                        {
                            const string sqlInsert = @"
INSERT INTO ag_arbol (
    nombre, descripcion, idperiodo, fecha_inicio, fecha_fin, ventana_origen, manual_ventana,
    poblacion, generaciones, tasa_cruce, tasa_mutacion, elitismo, estancamiento_max, semilla,
    usuario, subunidad, creado_en, modificado_en, activo
) VALUES (
    @Nombre, @Descripcion, @IdPeriodo, @FechaInicio, @FechaFin, @VentanaOrigen, @ManualVentana,
    @Poblacion, @Generaciones, @TasaCruce, @TasaMutacion, @Elitismo, @EstancamientoMax, @Semilla,
    @Usuario, @Subunidad, @CreadoEn, @ModificadoEn, @Activo
);
SELECT last_insert_rowid();";

                            idArbol = await db.ExecuteScalarAsync<long>(new CommandDefinition(sqlInsert, new
                            {
                                perfil.Nombre,
                                Descripcion = perfil.Descripcion ?? string.Empty,
                                perfil.IdPeriodo,
                                FechaInicio = perfil.FechaInicio ?? string.Empty,
                                FechaFin = perfil.FechaFin ?? string.Empty,
                                VentanaOrigen = perfil.VentanaOrigen ?? "Periodos",
                                perfil.ManualVentana,
                                perfil.Poblacion,
                                perfil.Generaciones,
                                perfil.TasaCruce,
                                perfil.TasaMutacion,
                                perfil.Elitismo,
                                perfil.EstancamientoMax,
                                perfil.Semilla,
                                Usuario = perfil.Usuario ?? string.Empty,
                                Subunidad = perfil.Subunidad ?? string.Empty,
                                CreadoEn = ahora,
                                ModificadoEn = ahora,
                                perfil.Activo
                            }, transaction: tx, cancellationToken: ct));
                        }

                        // Reemplazar nodos en bloque (DELETE + INSERT)
                        const string sqlDeleteNodos = "DELETE FROM ag_arbol_nodo WHERE idarbol = @IdArbol;";
                        await db.ExecuteAsync(new CommandDefinition(sqlDeleteNodos, new { IdArbol = idArbol }, transaction: tx, cancellationToken: ct));

                        if (nodos != null && nodos.Count > 0)
                        {
                            const string sqlInsertNodo = @"
INSERT INTO ag_arbol_nodo (
    idarbol, idperiodo, idsede, idprograma, semestre, grupo, modalidad, aula_ref,
    capacidad, tipo_aula, caracteristicas_json, codigo, nombre, es_virtual, habilitado, subunidad,
    jornada, horas_dia, dias_ventana
) VALUES (
    @IdArbol, @IdPeriodo, @IdSede, @IdPrograma, @Semestre, @Grupo, @Modalidad, @AulaRef,
    @Capacidad, @TipoAula, @CaracteristicasJson, @Codigo, @Nombre, @EsVirtual, @Habilitado, @Subunidad,
    @Jornada, @HorasDia, @DiasVentana
);";

                            var nodosParams = nodos.Select(n => new
                            {
                                IdArbol = idArbol,
                                n.IdPeriodo,
                                n.IdSede,
                                n.IdPrograma,
                                Semestre = n.Semestre ?? string.Empty,
                                Grupo = n.Grupo ?? "A",
                                Modalidad = n.Modalidad ?? "Presencial",
                                AulaRef = n.AulaRef ?? string.Empty,
                                n.Capacidad,
                                TipoAula = n.TipoAula ?? "ESTANDAR",
                                CaracteristicasJson = n.CaracteristicasJson ?? "{}",
                                Codigo = n.Codigo ?? string.Empty,
                                Nombre = n.Nombre ?? string.Empty,
                                n.EsVirtual,
                                Habilitado = n.Habilitado > 0 ? 1 : 0,
                                Subunidad = n.Subunidad ?? string.Empty,
                                Jornada = n.Jornada ?? string.Empty,
                                n.HorasDia,
                                DiasVentana = n.DiasVentana ?? string.Empty
                            });

                            await db.ExecuteAsync(new CommandDefinition(sqlInsertNodo, nodosParams, transaction: tx, cancellationToken: ct));
                        }

                        tx.Commit();
                        NotifyMarkDirty();
                        return Result.Success(idArbol);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<long>("Operación de guardado cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al guardar perfil '{perfil?.Nombre}'", ex);
                return Result.Failure<long>($"Error al guardar perfil: {ex.Message}");
            }
        }

        public async Task<Result<AgArbol>> ObtenerAsync(long idArbol, CancellationToken ct)
        {
            if (idArbol <= 0) return Result.Failure<AgArbol>("Identificador de perfil inválido.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sql = @"
SELECT id as Id, nombre as Nombre, descripcion as Descripcion, idperiodo as IdPeriodo,
       fecha_inicio as FechaInicio, fecha_fin as FechaFin, ventana_origen as VentanaOrigen,
       manual_ventana as ManualVentana, poblacion as Poblacion, generaciones as Generaciones,
       tasa_cruce as TasaCruce, tasa_mutacion as TasaMutacion, elitismo as Elitismo,
       estancamiento_max as EstancamientoMax, semilla as Semilla, usuario as Usuario,
       subunidad as Subunidad, creado_en as CreadoEn, modificado_en as ModificadoEn, activo as Activo
FROM ag_arbol
WHERE id = @idArbol;";

                    var item = await db.QueryFirstOrDefaultAsync<AgArbol>(new CommandDefinition(sql, new { idArbol }, cancellationToken: ct));
                    if (item == null) return Result.Failure<AgArbol>($"No se encontró el perfil con id {idArbol}.");

                    return Result.Success(item);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgArbol>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al obtener perfil {idArbol}", ex);
                return Result.Failure<AgArbol>($"Error al leer perfil: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgArbolNodo>>> ObtenerNodosAsync(long idArbol, CancellationToken ct)
        {
            if (idArbol <= 0) return Result.Failure<IReadOnlyCollection<AgArbolNodo>>("Identificador de perfil inválido.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sql = @"
SELECT id as Id, idarbol as IdArbol, idperiodo as IdPeriodo, idsede as IdSede, idprograma as IdPrograma,
       semestre as Semestre, grupo as Grupo, modalidad as Modalidad, aula_ref as AulaRef,
       capacidad as Capacidad, tipo_aula as TipoAula, caracteristicas_json as CaracteristicasJson,
       codigo as Codigo, nombre as Nombre, es_virtual as EsVirtual, habilitado as Habilitado, subunidad as Subunidad,
       jornada as Jornada, horas_dia as HorasDia, dias_ventana as DiasVentana
FROM ag_arbol_nodo
WHERE idarbol = @idArbol AND habilitado = 1
ORDER BY idsede, idprograma, semestre, grupo, modalidad;";

                    var items = (await db.QueryAsync<AgArbolNodo>(new CommandDefinition(sql, new { idArbol }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgArbolNodo>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgArbolNodo>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al obtener nodos del perfil {idArbol}", ex);
                return Result.Failure<IReadOnlyCollection<AgArbolNodo>>($"Error al leer nodos del perfil: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgArbol>>> ListarAsync(string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sql = @"
SELECT id as Id, nombre as Nombre, descripcion as Descripcion, idperiodo as IdPeriodo,
       fecha_inicio as FechaInicio, fecha_fin as FechaFin, ventana_origen as VentanaOrigen,
       manual_ventana as ManualVentana, poblacion as Poblacion, generaciones as Generaciones,
       tasa_cruce as TasaCruce, tasa_mutacion as TasaMutacion, elitismo as Elitismo,
       estancamiento_max as EstancamientoMax, semilla as Semilla, usuario as Usuario,
       subunidad as Subunidad, creado_en as CreadoEn, modificado_en as ModificadoEn, activo as Activo
FROM ag_arbol
WHERE (subunidad = @subunidad OR subunidad = '' OR @subunidad = '')
  AND activo = 1
ORDER BY nombre;";

                    var items = (await db.QueryAsync<AgArbol>(new CommandDefinition(sql, new { subunidad = subunidad ?? string.Empty }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgArbol>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgArbol>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgArbolRepository] Error al listar perfiles", ex);
                return Result.Failure<IReadOnlyCollection<AgArbol>>($"Error al listar perfiles: {ex.Message}");
            }
        }

        public async Task<Result<bool>> ExisteNombreAsync(string nombre, string subunidad, long excluirId, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(nombre)) return Result.Success(false);

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sql = @"
SELECT COUNT(1) 
FROM ag_arbol 
WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(@nombre))
  AND COALESCE(subunidad, '') = COALESCE(@subunidad, '')
  AND (@excluirId <= 0 OR id <> @excluirId);";

                    int count = await db.ExecuteScalarAsync<int>(new CommandDefinition(sql, new
                    {
                        nombre,
                        subunidad = subunidad ?? string.Empty,
                        excluirId
                    }, cancellationToken: ct));

                    return Result.Success(count > 0);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<bool>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al verificar existencia de nombre '{nombre}'", ex);
                return Result.Failure<bool>($"Error al validar nombre: {ex.Message}");
            }
        }

        public async Task<Result> RenombrarAsync(long idArbol, string nuevoNombre, CancellationToken ct)
        {
            if (idArbol <= 0) return Result.Failure("Identificador de perfil inválido.");
            if (string.IsNullOrWhiteSpace(nuevoNombre)) return Result.Failure("El nuevo nombre no puede estar vacío.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sql = @"
UPDATE ag_arbol 
SET nombre = @nuevoNombre, modificado_en = @ahora 
WHERE id = @idArbol;";

                    int afectadas = await db.ExecuteAsync(new CommandDefinition(sql, new
                    {
                        idArbol,
                        nuevoNombre = nuevoNombre.Trim(),
                        ahora = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    }, cancellationToken: ct));

                    if (afectadas > 0) NotifyMarkDirty();
                    return afectadas > 0 ? Result.Success() : Result.Failure($"No se encontró el perfil {idArbol}.");
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al renombrar perfil {idArbol}", ex);
                return Result.Failure($"Error al renombrar: {ex.Message}");
            }
        }

        public async Task<Result> EliminarAsync(long idArbol, CancellationToken ct)
        {
            if (idArbol <= 0) return Result.Failure("Identificador de perfil inválido.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    using (var tx = db.BeginTransaction())
                    {
                        ct.ThrowIfCancellationRequested();

                        const string sqlDeleteNodos = "DELETE FROM ag_arbol_nodo WHERE idarbol = @idArbol;";
                        await db.ExecuteAsync(new CommandDefinition(sqlDeleteNodos, new { idArbol }, transaction: tx, cancellationToken: ct));

                        const string sqlDeletePerfil = "DELETE FROM ag_arbol WHERE id = @idArbol;";
                        int afectadas = await db.ExecuteAsync(new CommandDefinition(sqlDeletePerfil, new { idArbol }, transaction: tx, cancellationToken: ct));

                        tx.Commit();
                        NotifyMarkDirty();
                        return afectadas > 0 ? Result.Success() : Result.Failure($"No se encontró el perfil con id {idArbol}.");
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgArbolRepository] Error al eliminar perfil {idArbol}", ex);
                return Result.Failure($"Error al eliminar: {ex.Message}");
            }
        }
    }
}
