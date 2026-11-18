using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Entidad de dominio que representa un perfil guardado de selección de árbol académico (ag_arbol).
    /// Contiene metadatos, ventana temporal y parámetros evolutivos asociados (Fase 237).
    /// </summary>
    public class AgArbol
    {
        public long Id { get; set; }
        public string Nombre { get; set; }
        public string Descripcion { get; set; }
        public int IdPeriodo { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string VentanaOrigen { get; set; }
        public int ManualVentana { get; set; }
        public int Poblacion { get; set; }
        public int Generaciones { get; set; }
        public double TasaCruce { get; set; }
        public double TasaMutacion { get; set; }
        public int Elitismo { get; set; }
        public int EstancamientoMax { get; set; }
        public int Semilla { get; set; }
        public string Usuario { get; set; }
        public string Subunidad { get; set; }
        public string CreadoEn { get; set; }
        public string ModificadoEn { get; set; }
        public int Activo { get; set; }

        public AgArbol()
        {
            Nombre = string.Empty;
            Descripcion = string.Empty;
            FechaInicio = string.Empty;
            FechaFin = string.Empty;
            VentanaOrigen = "Periodos";
            Poblacion = 50;
            Generaciones = 200;
            TasaCruce = 0.85;
            TasaMutacion = 0.10;
            Elitismo = 2;
            EstancamientoMax = AgParametrosPorDefecto.EstancamientoMax;
            Semilla = 0;
            Usuario = string.Empty;
            Subunidad = string.Empty;
            CreadoEn = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            ModificadoEn = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Activo = 1;
        }

        public override string ToString() => $"{Nombre} (P={Poblacion}, G={Generaciones})";
    }
}
