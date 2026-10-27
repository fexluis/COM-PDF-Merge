using System;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Fila de saneamiento V76: dimensiones actuales de ag_aula y de ag_arbol_nodo (Fase 1213).
    /// Evita el mapeo posicional frágil de ValueTuple de 8 elementos en Dapper.
    /// </summary>
    internal sealed class AgAulaDimensionesCandidato
    {
        public long Id { get; set; }
        public string AulaRef { get; set; }
        public string JornadaActual { get; set; }
        public int HorasDiaActual { get; set; }
        public string DiasVentanaActual { get; set; }
        public string JornadaNodo { get; set; }
        public int HorasDiaNodo { get; set; }
        public string DiasVentanaNodo { get; set; }
    }
}
