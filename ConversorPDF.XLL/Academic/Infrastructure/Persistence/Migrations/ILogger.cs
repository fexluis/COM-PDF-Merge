using System;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public interface ILogger
    {
        void Info(string message);
        void Warning(string message);
        void Error(string message, Exception ex = null);
    }
}
