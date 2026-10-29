using System;
using System.Data;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence
{
    public class UsuarioVaultMover : IUsuarioVaultMover
    {
        private readonly Func<string, string> _resolvePath;
        private readonly Func<string, EncryptedSQLiteContext> _openVault;
        private readonly Action<byte[], string> _persist;

        public UsuarioVaultMover()
            : this(
                SqliteHelper.GetDbPathForSubunidad,
                SqliteHelper.GetContextForPath,
                (bytes, path) => SqliteHelper.UpdateVaultDirectly(bytes, path))
        {
        }

        internal UsuarioVaultMover(
            Func<string, string> resolvePath,
            Func<string, EncryptedSQLiteContext> openVault,
            Action<byte[], string> persist)
        {
            _resolvePath = resolvePath;
            _openVault = openVault;
            _persist = persist;
        }

        public Result<MovimientoVaultResultado> Mover(Usuario usuario, string subunidadDestino)
        {
            if (usuario == null)
            {
                return Result.Failure<MovimientoVaultResultado>("El usuario es requerido.");
            }

            if (string.IsNullOrWhiteSpace(usuario.NombreUsuario))
            {
                return Result.Failure<MovimientoVaultResultado>("El nombre de usuario es requerido.");
            }

            if (string.IsNullOrWhiteSpace(subunidadDestino))
            {
                return Result.Failure<MovimientoVaultResultado>("La subunidad destino es requerida.");
            }

            string destinoNormalizado = subunidadDestino.Trim();

            try
            {
                string destinoPath = _resolvePath(destinoNormalizado);
                string origenPath = SqliteHelper.GetVaultPath();

                // R-MUV-04: Si el destino es igual al origen o la subunidad no cambia, es un no-op sin pérdida
                if (string.Equals(destinoPath, origenPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(usuario.Subunidad?.Trim(), destinoNormalizado, StringComparison.OrdinalIgnoreCase))
                {
                    return Result.Success(new MovimientoVaultResultado
                    {
                        SubunidadDestino = destinoNormalizado,
                        RutaVaultDestino = destinoPath,
                        Insertado = false,
                        Actualizado = false
                    });
                }

                using (var ctx = _openVault(destinoPath))
                {
                    var conn = ctx.Open();
                    var upsertResult = UpsertEnVault(conn, usuario, destinoNormalizado);
                    if (upsertResult.IsFailure)
                    {
                        return upsertResult;
                    }

                    upsertResult.Value.RutaVaultDestino = destinoPath;

                    // Persistir el destino de forma resiliente
                    ctx.CloseConnectionResiliently();
                    byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(ctx.TempDbPath);
                    if (plainBytes == null || plainBytes.Length == 0)
                    {
                        return Result.Failure<MovimientoVaultResultado>("No se pudieron leer los bytes temporales del vault destino.");
                    }

                    _persist(plainBytes, destinoPath);
                    ctx.MarkChangesPersisted();

                    return upsertResult;
                }
            }
            catch (Exception ex)
            {
                return Result.Failure<MovimientoVaultResultado>($"Error al mover usuario al vault destino '{destinoNormalizado}': {ex.Message}");
            }
        }

        internal static Result<MovimientoVaultResultado> UpsertEnVault(IDbConnection connection, Usuario usuario, string subunidadDestino)
        {
            if (connection == null)
            {
                return Result.Failure<MovimientoVaultResultado>("Conexión de base de datos es requerida.");
            }

            if (usuario == null)
            {
                return Result.Failure<MovimientoVaultResultado>("El usuario es requerido.");
            }

            if (string.IsNullOrWhiteSpace(usuario.NombreUsuario))
            {
                return Result.Failure<MovimientoVaultResultado>("El nombre de usuario es requerido.");
            }

            try
            {
                string normalizedUser = usuario.NombreUsuario.Trim().ToLowerInvariant();
                string deterministicId = "usr_" + normalizedUser;

                // Verificar existencia previa para determinar si es INSERT o UPDATE
                var existingId = connection.QueryFirstOrDefault<string>(
                    "SELECT id FROM usuarios WHERE id = @id OR LOWER(TRIM(nombreusuario)) = @usr LIMIT 1;",
                    new { id = deterministicId, usr = normalizedUser });

                bool isUpdate = existingId != null;

                // Si existía con id legacy, unificar al id determinista
                if (existingId != null && !string.Equals(existingId, deterministicId, StringComparison.OrdinalIgnoreCase))
                {
                    connection.Execute(
                        "UPDATE usuarios SET id = @deterministicId WHERE id = @existingId;",
                        new { deterministicId, existingId });
                }

                string perfil = string.IsNullOrWhiteSpace(usuario.Perfil) ? "Academica" : usuario.Perfil;

                connection.Execute(@"
                    INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at)
                    VALUES (@id, @nombre, @nombreusuario, @subunidad, @perfil, @activo, @firma, datetime('now'))
                    ON CONFLICT(id) DO UPDATE SET
                        nombre = excluded.nombre,
                        nombreusuario = excluded.nombreusuario,
                        subunidad = excluded.subunidad,
                        perfil = excluded.perfil,
                        activo = excluded.activo,
                        firma = excluded.firma,
                        updated_at = datetime('now');
                ", new
                {
                    id = deterministicId,
                    nombre = usuario.Nombre ?? usuario.NombreUsuario,
                    nombreusuario = usuario.NombreUsuario.Trim(),
                    subunidad = subunidadDestino,
                    perfil,
                    activo = usuario.Activo,
                    firma = usuario.Firma
                });

                return Result.Success(new MovimientoVaultResultado
                {
                    SubunidadDestino = subunidadDestino,
                    Insertado = !isUpdate,
                    Actualizado = isUpdate
                });
            }
            catch (Exception ex)
            {
                return Result.Failure<MovimientoVaultResultado>($"Error al ejecutar upsert en vault: {ex.Message}");
            }
        }
    }
}
