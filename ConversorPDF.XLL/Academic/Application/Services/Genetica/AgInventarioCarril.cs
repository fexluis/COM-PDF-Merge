using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Inventario de ranuras por carril (aula + patrón semanal) para la reparación R5 (Fase 2366).
    /// Separa la disponibilidad real del carril de la ocupación agregada de la cohorte.
    /// </summary>
    internal sealed class AgInventarioCarril
    {
        public string Clave { get; set; } = string.Empty;

        public int IdAula { get; set; }

        public string DiasPatron { get; set; } = string.Empty;

        public IReadOnlyList<int> Dias { get; set; } = new List<int>();

        public IReadOnlyList<DateTime> FechasValidas { get; set; } = new List<DateTime>();

        public IReadOnlyList<DateTime> FechasOcupadas { get; set; } = new List<DateTime>();

        public IReadOnlyList<DateTime> FechasLibres { get; set; } = new List<DateTime>();
    }
}
