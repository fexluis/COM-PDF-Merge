using System;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// Garantiza que la fila del usuario exista en el vault de su subunidad, con
    /// identidad determinista (<c>usr_&lt;nombreusuario&gt;</c>) y la subunidad resuelta.
    /// Se invoca al conmutar de vault en el arranque (<see cref="SyncContext"/>) y en el registro.
    /// </summary>
    internal static class UsuarioVaultProvisioner
    {
        /// <summary>
        /// Normaliza duplicados por nombre de usuario y hace upsert de la fila con la
        /// subunidad indicada. Idempotente.
        /// </summary>
        internal static void EnsureUsuarioEnVault(IDbConnection connection, string nombreUsuario, string nombre, string subunidad)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            if (string.IsNullOrWhiteSpace(nombreUsuario))
            {
                throw new ArgumentException("El nombre de usuario es obligatorio.", nameof(nombreUsuario));
            }

            if (string.IsNullOrWhiteSpace(subunidad))
            {
                throw new ArgumentException("La subunidad es obligatoria.", nameof(subunidad));
            }

            string usuario = nombreUsuario.Trim();
            string idDeterminista = "usr_" + usuario.ToLowerInvariant();
            string subunidadNormalizada = subunidad.Trim().ToUpperInvariant();
            string nombreNormalizado = (nombre ?? string.Empty).Trim();

            // 1) Eliminar duplicados legacy de la misma cuenta: sobrevive la identidad determinista.
            connection.Execute(
                @"DELETE FROM usuarios
                  WHERE LOWER(TRIM(nombreusuario)) = LOWER(TRIM(@NombreUsuario))
                    AND id <> @Id;",
                new { NombreUsuario = usuario, Id = idDeterminista });

            // 2) Upsert con la subunidad resuelta (no pisa el nombre si llega vacio).
            connection.Execute(
                @"INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, updated_at)
                  VALUES (@Id, @Nombre, @NombreUsuario, @Subunidad, 'Academica', 1, datetime('now'))
                  ON CONFLICT(id) DO UPDATE SET
                      nombre = CASE
                                   WHEN TRIM(COALESCE(excluded.nombre, '')) <> '' THEN excluded.nombre
                                   ELSE usuarios.nombre
                               END,
                      subunidad = excluded.subunidad,
                      updated_at = datetime('now');",
                new
                {
                    Id = idDeterminista,
                    Nombre = nombreNormalizado,
                    NombreUsuario = usuario,
                    Subunidad = subunidadNormalizada
                });
        }
    }
}
