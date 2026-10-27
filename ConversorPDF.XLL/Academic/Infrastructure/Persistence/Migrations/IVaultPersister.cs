using System;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public interface IVaultPersister
    {
        void Persist(byte[] plainDatabaseBytes, string vaultPath, IProgress<string> progress = null);
    }
}
