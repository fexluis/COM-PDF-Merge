using System;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    internal class LoggerAdapter : ILogger
    {
        public void Info(string message) => LoggerHelper.Info(message);
        public void Warning(string message) => LoggerHelper.Warning(message);
        public void Error(string message, Exception ex = null) => LoggerHelper.Error(message, ex);
    }
}
