using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Modelo inmutable de patrón semanal emergente derivado del recorrido calendario real (Fase 381 / Fase 391).
    /// </summary>
    public sealed class AgPatronSemanal
    {
        public IReadOnlyList<int> DiasSemana { get; }
        public IReadOnlyList<int> SesionesPorSemana { get; }
        public IReadOnlyList<DateTime> FechasValidas { get; }

        public AgPatronSemanal(
            IReadOnlyList<int> diasSemana,
            IReadOnlyList<int> sesionesPorSemana,
            IReadOnlyList<DateTime> fechasValidas)
        {
            DiasSemana = diasSemana ?? new int[0];
            SesionesPorSemana = sesionesPorSemana ?? new int[0];
            FechasValidas = fechasValidas ?? new DateTime[0];
        }
    }
}
