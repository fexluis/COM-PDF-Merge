using System;
using System.Data;
using System.Linq;
using CSharpFunctionalExtensions;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Determina a qué subunidad (sede territorial) pertenece el usuario de Windows
    ///   que está ejecutando el add-in en este momento.
    ///
    /// MECANISMO:
    ///   Consulta la tabla `usuarios` por Environment.UserName.
    ///   Si el usuario existe y está activo, retorna su subunidad.
    ///   Si no existe o está inactivo, retorna Result.Failure con un mensaje descriptivo.
    ///
    /// PATRÓN RESULT:
    ///   Usa CSharpFunctionalExtensions.Result&lt;T&gt; para comunicar errores de negocio
    ///   sin excepciones, siguiendo las reglas de arquitectura del proyecto.
    ///
    /// USO:
    ///   var result = _resolver.Resolve(connection);
    ///   if (result.IsFailure) { // mostrar FormRegistroSubunidad }
    ///   string subunidad = result.Value; // "VALLE", "ANTIOQUIA", etc.
    ///
    /// RETROCOMPATIBILIDAD:
    ///   Si la tabla usuarios aún no existe (DB sin migración V24), retorna Failure
    ///   con mensaje claro. No lanza excepción.
    /// </summary>
    internal sealed class SubunidadResolver
    {
        // ====================================================================
        // CONSTANTES
        // ====================================================================

        /// <summary>
        /// SQL para buscar la subunidad del usuario actual.
        /// Se filtra por activo = 1 para excluir usuarios dados de baja temporalmente.
        /// La columna nombreusuario almacena el valor de Environment.UserName (lowercase).
        /// </summary>
        private const string SqlGetSubunidad = @"
            SELECT subunidad
            FROM   usuarios
            WHERE  LOWER(nombreusuario) = LOWER(@UserName)
              AND  activo = 1
            ORDER BY activo DESC, updated_at DESC, id ASC
            LIMIT 1;";

        private const string SqlGetCandidatos = @"
            SELECT id, subunidad, updated_at
            FROM   usuarios
            WHERE  LOWER(nombreusuario) = LOWER(@UserName)
            ORDER BY activo DESC, updated_at DESC, id ASC;";

        /// <summary>
        /// SQL para verificar si la tabla usuarios existe en la DB actual.
        /// Necesario para manejar el caso de DB sin migración V24.
        /// </summary>
        private const string SqlCheckTableExists = @"
            SELECT COUNT(*)
            FROM   sqlite_master
            WHERE  type = 'table'
              AND  name = 'usuarios';";

        /// <summary>
        /// SQL fallback para leer la subunidad global almacenada en config.
        /// Se usa cuando el usuario de Windows no está explícitamente en la tabla usuarios.
        /// </summary>
        private const string SqlGetSubunidadFromConfig = @"
            SELECT subunidad
            FROM   config
            LIMIT 1;";

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Resuelve la subunidad del usuario de Windows actual.
        ///
        /// CASOS DE RETORNO:
        ///   - Result.Success("VALLE")    → usuario encontrado y activo.
        ///   - Result.Failure("...")      → usuario no existe, inactivo, o tabla ausente.
        ///
        /// NOTA: Este método hace consultas DB síncronas porque se ejecuta durante el
        /// arranque del add-in (STA thread) donde async/await es problemático.
        /// Las consultas son rápidas (índice en nombreusuario).
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta (de EncryptedSQLiteContext).</param>
        internal Result<string> Resolve(IDbConnection connection)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            string currentUser = Environment.UserName;
            LoggerHelper.Info($"[SubunidadResolver] Resolviendo subunidad para usuario Windows: '{currentUser}'");

            try
            {
                // Paso 1: Verificar que la tabla usuarios existe (migración V24 aplicada)
                int tableCount = connection.QueryFirstOrDefault<int>(SqlCheckTableExists);
                if (tableCount == 0)
                {
                    string msg = "[SubunidadResolver] La tabla 'usuarios' no existe. " +
                                 "La migración V24 no se ha ejecutado. " +
                                 "Ejecute la migración antes de usar el sistema de sync.";
                    LoggerHelper.Warning(msg);
                    return Result.Failure<string>(
                        "La base de datos no está preparada para sincronización. " +
                        "Contacte al administrador del sistema.");
                }

                // Paso 2: Buscar el usuario actual en la tabla usuarios
                var candidatos = connection.Query(
                    SqlGetCandidatos,
                    new { UserName = currentUser }).ToList();

                if (candidatos.Count > 1)
                {
                    string candInfo = string.Join("; ", candidatos.Select(c => $"{c.id}|{c.subunidad}|{c.updated_at}"));
                    LoggerHelper.Warning($"[SubunidadResolver] Múltiples registros encontrados para usuario '{currentUser}': {candInfo}. Seleccionando el más reciente por updated_at.");
                }

                string subunidad = connection.QueryFirstOrDefault<string>(
                    SqlGetSubunidad,
                    new { UserName = currentUser });

                if (string.IsNullOrWhiteSpace(subunidad))
                {
                    string msg = $"[SubunidadResolver] Usuario '{currentUser}' no encontrado en la tabla usuarios o está inactivo. Intentando lectura desde config...";
                    LoggerHelper.Warning(msg);

                    // Logging de diagnóstico para depuración ciega
                    try
                    {
                        int userCount = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM usuarios;");
                        LoggerHelper.Info($"[SubunidadResolver] Total usuarios en tabla: {userCount}");

                        var allUsers = connection.Query("SELECT nombreusuario, subunidad, activo FROM usuarios;");
                        string allUsersInfo = string.Join("; ", allUsers.Select(u => $"{u.nombreusuario}|{u.subunidad}|{u.activo}"));
                        LoggerHelper.Info($"[SubunidadResolver] Registros: {allUsersInfo}");
                    }
                    catch (Exception diagnosticEx)
                    {
                        LoggerHelper.Warning($"[SubunidadResolver] Error recolectando diagnóstico: {diagnosticEx.Message}");
                    }

                    // Paso 3: Fallback a tabla config
                    try
                    {
                        string configSubunidad = connection.QueryFirstOrDefault<string>(SqlGetSubunidadFromConfig);
                        if (!string.IsNullOrWhiteSpace(configSubunidad))
                        {
                            string normalizedConfigSubunidad = configSubunidad.ToUpperInvariant().Trim();
                            LoggerHelper.Info($"[SubunidadResolver] Subunidad resuelta vía fallback (tabla config): '{normalizedConfigSubunidad}'");
                            return Result.Success(normalizedConfigSubunidad);
                        }
                    }
                    catch (Exception configEx)
                    {
                        LoggerHelper.Warning($"[SubunidadResolver] Error leyendo tabla config (fallback): {configEx.Message}");
                    }

                    return Result.Failure<string>(
                        $"El usuario de Windows '{currentUser}' no está registrado en el sistema.\n\n" +
                        "Debe registrarse para habilitar la sincronización entre sedes.");
                }

                string normalizedSubunidad = subunidad.ToUpperInvariant().Trim();
                LoggerHelper.Info($"[SubunidadResolver] Subunidad resuelta (vía usuarios): '{normalizedSubunidad}'");
                return Result.Success(normalizedSubunidad);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error consultando la subunidad del usuario: {ex.Message}";
                LoggerHelper.Error("[SubunidadResolver] Error inesperado al resolver subunidad.", ex);
                return Result.Failure<string>(errorMsg);
            }
        }
    }
}
