using System;
using System.Data;
using System.Linq;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Agrega la columna 'perfil' a la tabla 'usuarios' y carga datos semilla.
    /// FASE 33: Incluye defensa en profundidad para crear la tabla desde cero
    /// si migraciones anteriores (como V24) fallaron silenciosamente o si
    /// el entorno de pruebas inicializó la base de datos sin ella.
    /// </summary>
    public class MigrationV29_FixUsuariosPerfil : IMigration
    {
        public int Version => 29;
        public string Name => "Fix Usuarios Perfil and Seed Data";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            LoggerHelper.Info("[MigrationV29] Verificando columna 'perfil' en tabla usuarios...");
            
            // Guarda de existencia: si usuarios no existe, crearla con schema completo incluyendo perfil
            int tableExists = db.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='usuarios';",
                transaction: transaction);

            if (tableExists == 0)
            {
                LoggerHelper.Warning("[MigrationV29] Tabla 'usuarios' no existe. Creando con schema completo...");
                db.Execute(@"
                    CREATE TABLE IF NOT EXISTS usuarios (
                        id            TEXT PRIMARY KEY DEFAULT (lower(hex(randomblob(16)))),
                        nombre        TEXT NOT NULL,
                        nombreusuario TEXT NOT NULL,
                        subunidad     TEXT NOT NULL,
                        perfil        TEXT DEFAULT 'Admin',
                        activo        INTEGER DEFAULT 1,
                        created_at    TEXT DEFAULT (datetime('now')),
                        updated_at    TEXT DEFAULT (datetime('now'))
                    );", transaction: transaction);

                LoggerHelper.Info("[MigrationV29] Tabla 'usuarios' creada con columna 'perfil' incluida.");
            }
            else
            {
                // La tabla existe: solo agregar perfil si falta
                var columns = db.Query<dynamic>("PRAGMA table_info(usuarios)", transaction: transaction);
                bool hasPerfil = columns.Any(c => string.Equals(c.name.ToString(), "perfil", StringComparison.OrdinalIgnoreCase));

                if (!hasPerfil)
                {
                    db.Execute("ALTER TABLE usuarios ADD COLUMN perfil TEXT DEFAULT 'Admin';", transaction: transaction);
                    LoggerHelper.Info("[MigrationV29] Columna 'perfil' agregada a usuarios.");
                }
                else
                {
                    LoggerHelper.Info("[MigrationV29] Columna 'perfil' ya existe. Omitiendo ALTER TABLE.");
                }
            }

            LoggerHelper.Info("[MigrationV29] Insertando seed data de usuarios (si no existen)...");

            // 2. Insertar los usuarios semilla (solo si no existen, basado en nombreusuario para evitar duplicados UUID)
            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Usuario Predeterminado', 'Usuario', 'VALLE', 'Admin', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'Usuario');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Luis', 'luis.bechara', 'VALLE', 'Admin', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'luis.bechara');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Montano', 'alexander.montano', 'VALLE', 'Contador', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'alexander.montano');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'carmen', 'carmroja', 'VALLE', 'Academica', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'carmroja');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Eduin', 'eduin.palacios', 'VALLE', 'Admin', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'eduin.palacios');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Isadora', 'isadora.chavez', 'VALLE', 'Academica', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'isadora.chavez');", transaction: transaction);

            db.Execute(@"INSERT INTO usuarios (nombre, nombreusuario, subunidad, perfil, activo)
                SELECT 'Elizabeth', 'elizabeth.cordobao', 'VALLE', 'Pagador', 1
                WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE nombreusuario = 'elizabeth.cordobao');", transaction: transaction);
            
            LoggerHelper.Info("[MigrationV29] Migración completada.");
        }
    }
}
