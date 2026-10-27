using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Value Object para representar una franja de tiempo dentro de un día.
    /// </summary>
    public struct FranjaHoraria
    {
        public TimeSpan Inicio { get; }
        public TimeSpan Fin { get; }

        public FranjaHoraria(TimeSpan inicio, TimeSpan fin)
        {
            if (inicio >= fin)
                throw new ArgumentException("El inicio de la franja debe ser menor al fin.", nameof(inicio));

            Inicio = inicio;
            Fin = fin;
        }

        /// <summary>
        /// Determina si esta franja se solapa en tiempo con otra franja.
        /// Dos franjas [A_inicio, A_fin] y [B_inicio, B_fin] se solapan (intersección mayor a 0) si:
        /// A_inicio es menor a B_fin Y B_inicio es menor a A_fin
        /// </summary>
        public bool SeSolapaCon(FranjaHoraria otra)
        {
            return Inicio < otra.Fin && otra.Inicio < Fin;
        }

        public override string ToString()
        {
            // En TimeSpan 'hh' representa horas de 00 a 23 (formato 24h militar)
            return $"{Inicio:hh\\:mm} - {Fin:hh\\:mm}";
        }
    }
}
