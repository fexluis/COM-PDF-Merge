using System;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// Identificador tipado para nodos del árbol de selección académica de 6 niveles.
    /// Niveles: 1=Periodo, 2=Sede, 3=Programa, 4=Semestre, 5=Grupo, 6=Modalidad.
    /// </summary>
    public class AgNodoArbolKey
    {
        public int Nivel { get; set; }
        public int Id { get; set; }
        public string Codigo { get; set; }
        public string Descripcion { get; set; }
        public string Jornada { get; set; } = string.Empty;
        public int HorasDia { get; set; }
        public string DiasVentana { get; set; } = string.Empty;

        public AgNodoArbolKey()
        {
            Codigo = string.Empty;
            Descripcion = string.Empty;
            Jornada = string.Empty;
            HorasDia = 0;
            DiasVentana = string.Empty;
        }

        public AgNodoArbolKey(int nivel, int id, string codigo, string descripcion)
        {
            Nivel = nivel;
            Id = id;
            Codigo = codigo ?? string.Empty;
            Descripcion = descripcion ?? string.Empty;
        }

        public override string ToString() => $"[Nivel {Nivel}] {Codigo} - {Descripcion}";
    }
}
