using System;
namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Provee metodos para la optimizacion del estado de Excel durante operaciones intensivas.
    /// </summary>
    public static class ExcelOptimizeHelper
    {
        /// <summary>
        /// Aplica optimizaciones a la aplicacion de Excel y devuelve un objeto IDisposable para restaurar el estado.
        /// </summary>
        public static IDisposable Optimize(dynamic app) => 
            new ExcelAppStateHelper(app).ApplyOptimizedAndReturn();
    }
}