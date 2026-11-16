using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Capacidad fechada informativa, sin docente asignado ni autorización de escritura.</summary>
    internal sealed class AgSesionPropuestaAprobada
    {
        public DateTime Fecha { get; }
        public string Franja { get; }
        public double Horas { get; }
        internal AgSesionPropuestaAprobada(AgSesionPlanificada s)
        { Fecha = s.Fecha; Franja = s.Ranura.Franja; Horas = s.Ranura.Horas; }
    }

    /// <summary>Comparación por identidad completa. Recuperadas son ranuras nuevas respecto a la simulación
    /// actual, no sesiones productivas ni garantía docente; el déficit canónico se informa por separado.</summary>
    internal sealed class AgDetalleEscaladoAprobada
    {
        public AgClaveRelacionCobertura Clave { get; }
        public double DeficitCanonicoHoras { get; }
        public string OrigenPatron { get; }
        public string PatronAntes { get; }
        public string PatronSinRedistribucion { get; }
        public string PatronConRedistribucion { get; }
        public bool RedistribuyeLegado { get; }
        public IReadOnlyList<AgSesionCanonica> SesionesOrigen { get; }
        public IReadOnlyList<AgSesionPropuestaAprobada> Antes { get; }
        public IReadOnlyList<AgSesionPropuestaAprobada> SinRedistribucion { get; }
        public IReadOnlyList<AgSesionPropuestaAprobada> ConRedistribucion { get; }
        public IReadOnlyList<AgSesionPropuestaAprobada> RecuperadasSinRedistribucion { get; }
        public IReadOnlyList<AgSesionPropuestaAprobada> RecuperadasConRedistribucion { get; }
        public int ResidualSinRedistribucion { get; }
        public int ResidualConRedistribucion { get; }
        public int GananciaNetaSinRedistribucion => SinRedistribucion.Count - Antes.Count;
        public int GananciaNetaConRedistribucion => ConRedistribucion.Count - Antes.Count;
        internal AgDetalleEscaladoAprobada(AgValidacionRelacionPlan validacion, AgAulaAsignatura relacion,
            IEnumerable<AgSesionCanonica> origen, AgPlanOleadasResultado antes, AgPlanOleadasResultado sin,
            AgPlanOleadasResultado con, double unidad, bool redistribuye)
        {
            Clave = validacion.Clave; DeficitCanonicoHoras = validacion.Deficit;
            OrigenPatron = relacion.OrigenPatron; PatronAntes = relacion.DiasPatron; RedistribuyeLegado = redistribuye;
            SesionesOrigen = origen.ToList().AsReadOnly();
            Antes = Sesiones(antes); SinRedistribucion = Sesiones(sin); ConRedistribucion = Sesiones(con);
            PatronSinRedistribucion = Patron(sin); PatronConRedistribucion = Patron(con);
            RecuperadasSinRedistribucion = Nuevas(SinRedistribucion); RecuperadasConRedistribucion = Nuevas(ConRedistribucion);
            int demanda = (int)Math.Ceiling(validacion.HorasEsperadas / unidad);
            ResidualSinRedistribucion = Math.Max(0, demanda - SinRedistribucion.Count);
            ResidualConRedistribucion = Math.Max(0, demanda - ConRedistribucion.Count);
        }
        private IReadOnlyList<AgSesionPropuestaAprobada> Sesiones(AgPlanOleadasResultado plan) => plan.Planes
            .Where(p => p.IdAsignatura == Clave.IdAsignatura).SelectMany(p => p.Sesiones).OrderBy(s => s.Fecha)
            .ThenBy(s => s.Ranura.Franja).Select(s => new AgSesionPropuestaAprobada(s)).ToList().AsReadOnly();
        private string Patron(AgPlanOleadasResultado plan) => string.Join(",", plan.Planes
            .Where(p => p.IdAsignatura == Clave.IdAsignatura).SelectMany(p => p.DiasPatron).Distinct().OrderBy(d => d));
        private IReadOnlyList<AgSesionPropuestaAprobada> Nuevas(IEnumerable<AgSesionPropuestaAprobada> sesiones) => sesiones
            .Where(s => !Antes.Any(a => a.Fecha == s.Fecha && a.Franja == s.Franja && a.Horas == s.Horas)).ToList().AsReadOnly();
    }

    /// <summary>Un solo paso ligado a origen y configuración. No es una Decision aplicable ni contiene permisos.</summary>
    internal sealed class AgPropuestaEscaladoAprobada
    {
        public int IdAula { get; }
        public string AulaRef { get; }
        public string HuellaOrigen { get; }
        public string HuellaConfiguracion { get; }
        public string Desde { get; }
        public string Hacia { get; }
        public IReadOnlyList<int> DiasAnadidos { get; }
        public DateTime FechaLimite { get; }
        public bool Terminal { get; }
        public bool BusquedaAgotada { get; }
        public IReadOnlyList<AgDetalleEscaladoAprobada> Relaciones { get; }
        internal AgPropuestaEscaladoAprobada(AgAula aula, string origen, string configuracion, string siguiente,
            DateTime limite, bool agotada, IEnumerable<AgDetalleEscaladoAprobada> relaciones)
        {
            IdAula = aula.Id; AulaRef = aula.AulaRef; HuellaOrigen = origen; HuellaConfiguracion = configuracion;
            Desde = aula.DiasVentana; Hacia = siguiente; FechaLimite = limite;
            Terminal = string.IsNullOrEmpty(siguiente) || string.IsNullOrEmpty(AgEscaleraVentanas.Siguiente(siguiente));
            DiasAnadidos = AgDiasVentana.Dias(siguiente).Except(AgDiasVentana.Dias(Desde)).OrderBy(d => d).ToList().AsReadOnly();
            BusquedaAgotada = agotada; Relaciones = relaciones.ToList().AsReadOnly();
        }
    }
}
