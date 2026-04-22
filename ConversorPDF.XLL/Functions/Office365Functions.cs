using ExcelDna.Integration;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Functions
{
    public static class WindowsUserFunctions
    {
        [ExcelFunction(
            Name = "USUARIO_WINDOWS",
            Description = "Obtiene el nombre del usuario actual de Windows",
            Category = "ESAP - Windows"
        )]
        public static string UsuarioWindows() => 
            System.Security.Principal.WindowsIdentity.GetCurrent().Name;
    }
}