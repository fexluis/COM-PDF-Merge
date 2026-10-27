using System;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    internal class VaultPersister : IVaultPersister
    {
        public void Persist(byte[] plainDatabaseBytes, string vaultPath, IProgress<string> progress = null)
        {
            SqliteHelper.UpdateVaultDirectly(plainDatabaseBytes, vaultPath, progress);
        }
    }
}
