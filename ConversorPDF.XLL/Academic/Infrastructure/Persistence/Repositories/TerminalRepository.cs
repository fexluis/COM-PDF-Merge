using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class TerminalRepository : ITerminalRepository
    {
        private readonly EncryptedSQLiteContext _context;

        /// <summary>
        /// Inicializa el repositorio con el contexto encriptado de SQLite.
        /// Cambio (Fase 7): Uso de EncryptedSQLiteContext en lugar de Func de IDbConnection para asegurar persistencia en Vault.
        /// </summary>
        public TerminalRepository(EncryptedSQLiteContext context)
        {
            _context = context;
        }

        public Result<IReadOnlyCollection<Terminal>> GetAll()
        {
            try
            {
                var db = _context.Open();
                const string sql = "SELECT id as Id, terminal as TerminalName, maxvalor as MaxValor FROM terminales WHERE id > 1 ORDER BY id";
                var items = db.Query<Terminal>(sql).ToList();
                return Result.Success((IReadOnlyCollection<Terminal>)items);
            }
            catch (Exception ex)
            {
                return Result.Failure<IReadOnlyCollection<Terminal>>($"Error al obtener terminales: {ex.Message}");
            }
        }

        public Result Add(Terminal terminal)
        {
            try
            {
                var db = _context.Open();
                const string sql = "INSERT INTO terminales (terminal, maxvalor) VALUES (@TerminalName, @MaxValor)";
                int rows = db.Execute(sql, terminal);
                if (rows == 0) return Result.Failure("No se pudo insertar la terminal.");
                
                // Cambio (Fase 7): Notificar cambios al Vault para encriptación al cerrar
                _context.MarkChangesPersisted();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al insertar terminal: {ex.Message}");
            }
        }

        public Result Update(Terminal terminal)
        {
            try
            {
                var db = _context.Open();
                const string sql = "UPDATE terminales SET terminal = @TerminalName, maxvalor = @MaxValor WHERE id = @Id";
                int rows = db.Execute(sql, terminal);
                if (rows == 0) return Result.Failure("No se encontró la terminal a actualizar.");
                
                // Cambio (Fase 7): Notificar cambios al Vault para encriptación al cerrar
                _context.MarkChangesPersisted();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al actualizar terminal: {ex.Message}");
            }
        }

        public Result Delete(int id)
        {
            try
            {
                var db = _context.Open();
                const string sql = "DELETE FROM terminales WHERE id = @Id";
                int rows = db.Execute(sql, new { Id = id });
                if (rows == 0) return Result.Failure("No se encontró la terminal a eliminar.");
                
                // Cambio (Fase 7): Notificar cambios al Vault para encriptación al cerrar
                _context.MarkChangesPersisted();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al eliminar terminal: {ex.Message}");
            }
        }
    }
}
