using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Builder mutable de corta vida para construir instancias inmutables de <see cref="AcademicDateAdjustment"/>.
    /// Solo debe ser usado dentro de <see cref="ConversorPDF.XLL.Academic.Application.Services.AcademicDateSyncService"/>.
    /// </summary>
    internal class AcademicDateAdjustmentBuilder
    {
        private readonly List<string> _messages = new List<string>();

        public bool NeedsAdjustment { get; set; }
        public DateTime CalculatedStart { get; set; }
        public DateTime CalculatedEnd { get; set; }
        public bool ExceededOriginalLimit { get; set; }

        public void AddMessage(string message)
        {
            if (string.IsNullOrEmpty(message)) return;
            _messages.Add(message);
        }

        public AcademicDateAdjustment Build()
        {
            return new AcademicDateAdjustment(
                needsAdjustment: NeedsAdjustment,
                calculatedStart: CalculatedStart,
                calculatedEnd: CalculatedEnd,
                exceededOriginalLimit: ExceededOriginalLimit,
                messages: _messages.AsReadOnly());
        }
    }
}
