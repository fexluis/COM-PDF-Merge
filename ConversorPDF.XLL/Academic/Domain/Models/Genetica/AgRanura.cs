using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa una ranura temporal elemental de asignación: día de la semana ISO (1=LUN..7=DOM),
    /// franja horaria canónica (MANANA, TARDE, NOCHE) y duración de sesión en horas (Fase 1154).
    /// </summary>
    internal sealed class AgRanura : IEquatable<AgRanura>
    {
        public int Dia { get; }
        public string Franja { get; }
        public double Horas { get; }

        public AgRanura(int dia, string franja, double horas = 4.0)
        {
            Dia = dia;
            Franja = franja ?? string.Empty;
            Horas = horas > 0 ? horas : 4.0;
        }

        public bool Equals(AgRanura other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Dia == other.Dia
                   && string.Equals(Franja, other.Franja, StringComparison.OrdinalIgnoreCase)
                   && Math.Abs(Horas - other.Horas) < 0.001;
        }

        public override bool Equals(object obj)
        {
            return ReferenceEquals(this, obj) || (obj is AgRanura other && Equals(other));
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = Dia;
                hashCode = (hashCode * 397) ^ (Franja != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(Franja) : 0);
                hashCode = (hashCode * 397) ^ Horas.GetHashCode();
                return hashCode;
            }
        }

        public override string ToString()
        {
            return $"{Dia}:{Franja}({Horas:0.#}h)";
        }
    }
}
