using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa una terminal (terrestre o aérea) para viáticos.
    /// </summary>
    public class Terminal
    {
        public int Id { get; set; }
        public string TerminalName { get; set; } // Map to 'terminal' column
        public double MaxValor { get; set; }
    }
}
