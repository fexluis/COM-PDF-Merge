using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public interface ISchemaMigrator
    {
        void EnsureSchemaUpToDate(EncryptedSQLiteContext context);
    }
}
