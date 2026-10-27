using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Modo de operación para conmutar transparentemente entre el entorno de PRODUCCIÓN y el STAGING del Algoritmo Genético.
    /// </summary>
    public class AgModoDatos
    {
        public static AgModoDatos Determinista { get; } = new AgModoDatos(false, "PRODUCCIÓN");
        public static AgModoDatos Produccion => Determinista;
        public static AgModoDatos Ag { get; } = new AgModoDatos(true, "MODO AG");
        public static AgModoDatos Staging => Ag;

        public bool EsModoAg { get; }
        public string Nombre { get; }

        public string TablaVinculaciones => EsModoAg ? "ag_vinculaciones" : "vinculaciones";
        public string TablaSesiones => EsModoAg ? "ag_sesiones" : "sesiones";
        public string TablaSalvedades => EsModoAg ? "ag_salvedades_conflictos" : "salvedades_conflictos";
        public string VistaVinculacionesDetalle => EsModoAg ? "vw_AgVinculacionesDetalle" : "vw_VinculacionesDetalle";
        public string VistaSesionesDetalle => EsModoAg ? "vw_AgSesionesDetalle" : "vw_SesionesDetalle";

        public AgModoDatos(bool esModoAg, string nombre = null)
        {
            EsModoAg = esModoAg;
            Nombre = nombre ?? (esModoAg ? "MODO AG" : "PRODUCCIÓN");
        }

        public override string ToString() => Nombre;
    }
}
