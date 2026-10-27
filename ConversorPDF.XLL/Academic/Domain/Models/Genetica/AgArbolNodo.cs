using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Entidad de dominio que representa una hoja seleccionada del árbol académico vinculada a un perfil (ag_arbol_nodo).
    /// Almacena la tupla canónica S (AulaRef) junto con los atributos de aula necesarios para su materialización (Fase 237).
    /// </summary>
    public class AgArbolNodo
    {
        public long Id { get; set; }
        public long IdArbol { get; set; }
        public int IdPeriodo { get; set; }
        public int IdSede { get; set; }
        public int IdPrograma { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public string Modalidad { get; set; }
        public string AulaRef { get; set; }
        public int Capacidad { get; set; }
        public string TipoAula { get; set; }
        public string CaracteristicasJson { get; set; }
        public string Codigo { get; set; }
        public string Nombre { get; set; }
        public int EsVirtual { get; set; }
        public int Habilitado { get; set; }
        public string Subunidad { get; set; }
        public string Jornada { get; set; } = string.Empty;
        public int HorasDia { get; set; }
        public string DiasVentana { get; set; } = string.Empty;

        public AgArbolNodo()
        {
            Semestre = string.Empty;
            Grupo = "A";
            Modalidad = "Presencial";
            AulaRef = string.Empty;
            TipoAula = "ESTANDAR";
            CaracteristicasJson = "{}";
            Codigo = string.Empty;
            Nombre = string.Empty;
            Habilitado = 1;
            Subunidad = string.Empty;
            Jornada = string.Empty;
            HorasDia = 0;
            DiasVentana = string.Empty;
        }

        public override string ToString() => $"{AulaRef} (Arbol={IdArbol})";
    }
}
