using System.Data;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public interface IMigration
    {
        int Version { get; }
        string Name { get; }
        void Up(IDbConnection connection, IDbTransaction transaction);
    }
}
