using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    // S8: pruebas escritas, no ejecutadas; ninguna conexión ni almacenamiento.
    public class AgSimuladorEscaladoAprobadasTests
    {
        [Fact]
        public void AprobadaMarVie_OfreceUnPasoFechadoSinMutarNiAutorizar()
        {
            var c = Contexto();
            var huella = AgHuellaPlanVentana.Calcular(c, c.Aulas.Single());
            var propuesta = new AgSimuladorEscaladoAprobadas().Evaluar(c).Single();
            Assert.Equal("MAR-VIE", propuesta.Desde);
            Assert.Equal("LUN-VIE", propuesta.Hacia);
            Assert.Equal(new[] { 1 }, propuesta.DiasAnadidos);
            Assert.True(propuesta.Terminal);
            Assert.Equal(new DateTime(2027, 5, 30), propuesta.FechaLimite);
            var r = propuesta.Relaciones.Single();
            Assert.Equal(4, r.SesionesOrigen.Count);
            Assert.Equal(4, r.Antes.Count);
            Assert.Equal(5, r.SinRedistribucion.Count);
            Assert.Single(r.RecuperadasSinRedistribucion);
            Assert.Equal(new DateTime(2027, 5, 24), r.RecuperadasSinRedistribucion.Single().Fecha);
            Assert.Equal(0, r.ResidualSinRedistribucion);
            Assert.Equal(huella, AgHuellaPlanVentana.Calcular(c, c.Aulas.Single()));
            Assert.Equal(AgEstadoAula.Aprobado, c.Aulas.Single().Estado);
            Assert.Equal("MAR-VIE", c.Aulas.Single().DiasVentana);
            Assert.False(AgEscaladorRanuraAula.PuedeAplicarPropuesta(c.Aulas.Single()));
            Assert.True(c.ValidacionPlanAprobado.BloqueaPreparacion);
            Assert.Equal(0, c.ValidacionPlanAprobado.Aulas.Single().HorasPreservadas);
        }

        [Fact]
        public void CompletaCanonica_NoOfreceAunqueCapacidadActualDesdeCeroNoAlcance()
        {
            var c = Contexto(completa: true);
            Assert.Empty(new AgSimuladorEscaladoAprobadas().Evaluar(c));
        }

        [Theory]
        [InlineData("ausente")]
        [InlineData("ambigua")]
        [InlineData("extra")]
        [InlineData("desactualizada")]
        public void OrigenDesconocidoODesligado_NoSimula(string caso)
        {
            var c = Contexto(); var s = c.PlanAprobadoSnapshot;
            if (caso == "desactualizada") c.AulaAsignaturas.Single().HorasObjetivo += 4;
            else c.PlanAprobadoSnapshot = new AgPlanAprobadoSnapshot(s.Manifiesto,
                caso == "ausente" ? new AgRelacionPlanLeida[0] : s.Relaciones.Select(r =>
                    new AgRelacionPlanLeida(r.IdVinculacion, caso == "ambigua"
                        ? new[] { r.ClavesCandidatas.Single(), r.ClavesCandidatas.Single() } : r.ClavesCandidatas,
                        r.Sesiones, r.OrigenCanonico)), caso == "extra" ? new[] { "RELACION_EXTRA:99" } : new string[0], "catalogo");
            Assert.False(AgEscaladorRanuraAula.PuedeEvaluarPropuesta(c, c.Aulas.Single()));
            Assert.Empty(new AgSimuladorEscaladoAprobadas().Evaluar(c));
            Assert.Equal("MAR-VIE", c.Aulas.Single().DiasVentana);
        }

        [Theory]
        [InlineData("FIJO", 1)]
        [InlineData("LEGADO", 0)]
        public void PatronFijoNoSeLibera_LegadoSeComparaSinConsentimiento(string origen, int residual)
        {
            var c = Contexto(); var aa = c.AulaAsignaturas.Single();
            aa.OrigenPatron = origen; aa.DiasPatron = "2,3,4,5";
            var p = new AgSimuladorEscaladoAprobadas().Evaluar(c).Single();
            var r = p.Relaciones.Single();
            Assert.Equal(1, r.ResidualSinRedistribucion);
            Assert.Equal(residual, r.ResidualConRedistribucion);
            Assert.Equal(origen == "LEGADO", r.RedistribuyeLegado);
            Assert.Equal("2,3,4,5", r.PatronAntes);
            Assert.Equal("2,3,4,5", aa.DiasPatron);
            Assert.Equal(origen, aa.OrigenPatron);
            Assert.All(r.ConRedistribucion, x => Assert.InRange(x.Fecha, c.Ventana.FechaInicio, c.Ventana.FechaFin));
        }

        [Theory]
        [InlineData("VIE-DOM", "JUE-DOM", false)]
        [InlineData("JUE-DOM", "MIE-DOM", true)]
        [InlineData("VIE-SAB", "JUE-SAB", false)]
        [InlineData("JUE-SAB", "4,5,6,7", true)]
        [InlineData("4,5,6,7", "", true)]
        [InlineData("LUN-VIE", "", true)]
        public void EscaleraHistoricaYTerminal_NoEncadenaNiNormalizaOrigen(string desde, string hacia, bool terminal)
        {
            var c = Contexto(); c.Aulas.Single().DiasVentana = desde;
            var p = new AgSimuladorEscaladoAprobadas().Evaluar(c).Single();
            Assert.Equal(desde, p.Desde); Assert.Equal(hacia, p.Hacia); Assert.Equal(terminal, p.Terminal);
            Assert.Equal(desde, c.Aulas.Single().DiasVentana);
            if (hacia == "") Assert.Equal(p.Relaciones.Single().Antes.Count, p.Relaciones.Single().SinRedistribucion.Count);
        }

        [Fact]
        public void DeficitTemporalConIdentidadValida_NoSeConfundeConRatificacionNiOrigenDesconocido()
        {
            var c = Contexto(); var s = c.PlanAprobadoSnapshot; var r = s.Relaciones.Single();
            c.PlanAprobadoSnapshot = new AgPlanAprobadoSnapshot(s.Manifiesto, new[] {
                new AgRelacionPlanLeida(r.IdVinculacion, r.ClavesCandidatas, r.Sesiones.Select(x =>
                    new AgSesionCanonica(x.IdSesion, x.IdVinculacion, x.IdTercero, x.Fecha.Value.AddDays(7),
                        x.HoraInicio, x.HoraFin, x.HorasLectivas, x.Subunidad, x.OrigenCanonico)), r.OrigenCanonico)
            }, new string[0], "catalogo");
            c.ValidacionPlanAprobado = new AgValidadorPlanAprobado().Validar(c.PlanAprobadoSnapshot, c);
            Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, c.ValidacionPlanAprobado.Aulas.Single().Estado);
            Assert.True(AgEscaladorRanuraAula.PuedeEvaluarPropuesta(c, c.Aulas.Single()));
            Assert.Single(new AgSimuladorEscaladoAprobadas().Evaluar(c));
            Assert.True(c.ValidacionPlanAprobado.BloqueaPreparacion);
        }

        [Fact]
        public void Disponible_ConservaEvaluacionHeredadaSinExigirSnapshot()
        {
            var c = Contexto(); c.Aulas.Single().Estado = AgEstadoAula.Disponible;
            c.PlanAprobadoSnapshot = null; c.ManifiestoCobertura = null;
            Assert.True(AgEscaladorRanuraAula.PuedeEvaluarPropuesta(c, c.Aulas.Single()));
            Assert.True(AgEscaladorRanuraAula.PuedeAplicarPropuesta(c.Aulas.Single()));
            var d = new AgEscaladorRanuraAula(() => null).Evaluar(c).Single();
            Assert.Equal("LUN-VIE", d.Hacia); Assert.True(d.Aplica);
            Assert.Empty(new AgSimuladorEscaladoAprobadas().Evaluar(c));
        }

        private static AgContexto Contexto(bool completa = false)
        {
            var c = new AgContexto { Subunidad = "VALLE", IdPeriodo = 3, IdSede = 1, IdPrograma = 1,
                Ventana = VentanaSemestre.Crear(3, "2027", new DateTime(2027, 5, 24), new DateTime(2027, 5, 30), AgOrigenVentana.Manual).Value };
            c.Programas.Add(new Programa { Id = 1, Modalidad = "PRESENCIAL" });
            c.Aulas.Add(new AgAula { Id = 14, IdPeriodo = 3, IdSede = 1, IdPrograma = 1, AulaRef = "CALI I-B",
                Subunidad = "VALLE", Estado = AgEstadoAula.Aprobado, DiasVentana = "MAR-VIE", Jornada = "MANANA",
                HorasDia = 4, Semestre = "I", Grupo = "B" });
            c.Asignaturas.Add(new Asignatura { Id = 7, IdPrograma = 1, Semestre = "I", Horas = 20 });
            c.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 14, IdAsignatura = 7, HorasObjetivo = 20,
                Subunidad = "VALLE", OrigenPatron = "AUTOMATICO" });
            for (int d = 1; d <= 7; d++)
                c.Bloques.Add(new AgBloqueHorario { Id = d, DiaSemanaNumero = d, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            c.Bloques.Add(new AgBloqueHorario { Id = 8, DiaSemanaNumero = 6, Franja = "TARDE", DuracionHoras = 4, Activo = 1 });
            c.ConstruirIndices();
            c.ManifiestoCobertura = new AgConstructorManifiestoCobertura().Construir(c, new[] { 14 }, 16, "VALLE", c.Ventana.FechaInicio, c.Ventana.FechaFin).Value;
            var k = c.ManifiestoCobertura.Relaciones.Single().Clave;
            var sesiones = Enumerable.Range(completa ? 0 : 1, completa ? 5 : 4).Select(i =>
                new AgSesionCanonica(i + 1, 10, 9, c.Ventana.FechaInicio.AddDays(i), TimeSpan.FromHours(8), TimeSpan.FromHours(12), 4, "VALLE", "s:" + i));
            c.PlanAprobadoSnapshot = new AgPlanAprobadoSnapshot(c.ManifiestoCobertura,
                new[] { new AgRelacionPlanLeida(10, new[] { k }, sesiones, "v:10") }, new string[0], "catalogo");
            c.ValidacionPlanAprobado = new AgValidadorPlanAprobado().Validar(c.PlanAprobadoSnapshot, c);
            return c;
        }
    }
}
