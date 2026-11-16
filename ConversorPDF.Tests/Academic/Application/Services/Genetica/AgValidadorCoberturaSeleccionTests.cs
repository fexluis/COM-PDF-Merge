using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgValidadorCoberturaSeleccionTests
    {
        private static AgManifiestoCobertura Manifiesto(int aulas = 1, DateTime? fin = null)
        {
            var relaciones = Enumerable.Range(1, aulas).Select(a => new AgRelacionCobertura(
                new AgClaveRelacionCobertura(1, "VALLE", 1, 1, 1, a, a, "I", "A", "PRESENCIAL"), 4, 4, 1));
            return new AgManifiestoCobertura(new DateTime(2027, 2, 5), fin ?? new DateTime(2027, 5, 30), relaciones);
        }

        private static AgSesionCobertura Sesion(AgManifiestoCobertura m, int aula = 1,
            string origen = "NUEVA", long id = 1, int docente = 9,
            DateTime? fecha = null, int inicio = 8) => new AgSesionCobertura(
                m.Relaciones.Single(r => r.Clave.IdAula == aula).Clave, origen, id, docente,
                fecha ?? new DateTime(2027, 2, 5), TimeSpan.FromHours(inicio), TimeSpan.FromHours(inicio + 4), 4);

        [Fact]
        public void DisponibleSinManifiestoConservaElegibilidadAnterior()
        {
            var c = new AgContexto();
            c.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible });
            c.Aulas.Add(new AgAula { Id = 2, Estado = AgEstadoAula.Aprobado });
            Assert.True(c.PuedeGenerarAula(1));
            Assert.False(c.PuedeGenerarAula(2));
            c.ManifiestoCobertura = Manifiesto();
            Assert.True(c.PuedeGenerarAula(1));
            Assert.False(c.PuedeGenerarAula(2));
        }

        [Fact]
        public void RelacionEliminadaFallaAunqueGenesYStagingCoincidan()
        {
            var m = Manifiesto(2);
            var resultado = new AgValidadorCoberturaSeleccion().Validar(m, Array.Empty<AgSesionCobertura>(),
                new[] { Sesion(m) });
            Assert.True(resultado.IsFailure);
            Assert.StartsWith("COBERTURA_INCOMPLETA:", resultado.Error);
        }

        [Fact]
        public void SoloConservacionNoNecesitaStagingNiPoblacion()
        {
            var m = Manifiesto();
            Assert.True(new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { Sesion(m, origen: "CONSERVADA") }, Array.Empty<AgSesionCobertura>()).IsSuccess);
        }

        [Theory]
        [InlineData(1, 10, "CRUCE_AULA")]
        [InlineData(2, 9, "CRUCE_DOCENTE")]
        public void ConservadaBloqueaIntervaloDeNueva(int aula, int docente, string error)
        {
            var m = Manifiesto(2);
            var conservada = Sesion(m, origen: "CONSERVADA");
            var nueva = Sesion(m, aula, id: 2, docente: docente, inicio: aula == 1 ? 9 : 8);
            var r = new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { conservada }, new[] { nueva });
            Assert.True(r.IsFailure);
            Assert.Equal(error, r.Error);
        }

        [Fact]
        public void ReemplazoAutorizadoNoSeAutocruzaConSesionAnteriorExcluida()
        {
            var m = Manifiesto();
            Assert.True(new AgValidadorCoberturaSeleccion().Validar(m,
                Array.Empty<AgSesionCobertura>(), new[] { Sesion(m) }).IsSuccess);
        }

        [Theory]
        [InlineData(2027, 5, 31)]
        [InlineData(2027, 2, 4)]
        public void FechaExteriorFalla(int year, int month, int day)
        {
            var m = Manifiesto();
            var r = new AgValidadorCoberturaSeleccion().Validar(m, Array.Empty<AgSesionCobertura>(),
                new[] { Sesion(m, fecha: new DateTime(year, month, day)) });
            Assert.Equal("SESION_FUERA_DE_VENTANA", r.Error);
        }

        [Fact]
        public void PlazoMasCortoTambienBloquea()
        {
            var m = Manifiesto(fin: new DateTime(2027, 5, 20));
            var r = new AgValidadorCoberturaSeleccion().Validar(m, Array.Empty<AgSesionCobertura>(),
                new[] { Sesion(m, fecha: new DateTime(2027, 5, 21)) });
            Assert.Equal("SESION_FUERA_DE_VENTANA", r.Error);
        }

        [Fact]
        public void IdCanonicoDuplicadoNoCuentaDosVeces()
        {
            var m = Manifiesto(); var s = Sesion(m, origen: "CONSERVADA");
            var r = new AgValidadorCoberturaSeleccion().Validar(m, new[] { s, s }, Array.Empty<AgSesionCobertura>());
            Assert.Equal("SESION_DUPLICADA", r.Error);
        }

        [Fact]
        public void RelacionNoSeleccionadaEsInesperada()
        {
            var m = Manifiesto(); var otra = Manifiesto(2).Relaciones[1].Clave;
            var s = new AgSesionCobertura(otra, "NUEVA", 1, 9, new DateTime(2027, 2, 5),
                TimeSpan.FromHours(8), TimeSpan.FromHours(12), 4);
            Assert.Equal("RELACION_INESPERADA", new AgValidadorCoberturaSeleccion().Validar(m,
                Array.Empty<AgSesionCobertura>(), new[] { s }).Error);
        }

        [Fact]
        public void ConservadaUnaVezNoProduceAutocruce()
        {
            var m = Manifiesto();
            var conservada = Sesion(m, origen: "CONSERVADA");
            Assert.True(new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { conservada }, Array.Empty<AgSesionCobertura>()).IsSuccess);
        }

        [Fact]
        public void FestivoBloqueaInclusoCuandoLasHorasCoinciden()
        {
            var m = Manifiesto();
            var fecha = new DateTime(2027, 2, 5);
            var r = new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { Sesion(m, origen: "CONSERVADA") }, Array.Empty<AgSesionCobertura>(),
                festivos: new[] { (fecha, fecha) });
            Assert.Equal("SESION_EN_FESTIVO", r.Error);
        }

        [Fact]
        public void VacanteNoAutorizaSesionAusente()
        {
            var m = Manifiesto(2);
            Assert.StartsWith("COBERTURA_INCOMPLETA:", new AgValidadorCoberturaSeleccion().Validar(m,
                Array.Empty<AgSesionCobertura>(), new[] { Sesion(m) }).Error);
            Assert.Equal("DOCENTE_VACANTE", new AgValidadorCoberturaSeleccion().Validar(Manifiesto(),
                Array.Empty<AgSesionCobertura>(), new[] { Sesion(Manifiesto(), docente: 0) }).Error);
        }
    }
}
