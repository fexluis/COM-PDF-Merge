using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

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
            return Result.Failure<MovimientoVaultResultado>("Mover entre vaults no implementado");
        }

        internal static Result<MovimientoVaultResultado> UpsertEnVault(System.Data.IDbConnection connection, Usuario usuario, string subunidadDestino)
        {
            return Result.Failure<MovimientoVaultResultado>("Upsert en vault no implementado");
        }
    }
}
