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
        public void VentanaHastaJunio_NoEsBloqueadaPorUnaFechaFija()
        {
            // Requisito: el límite es la fecha de fin de la ventana semestral, no una fecha fija.
            var m = Manifiesto(fin: new DateTime(2027, 6, 30));
            var sesion = Sesion(m, fecha: new DateTime(2027, 6, 20));
            var resultado = new AgValidadorCoberturaSeleccion().Validar(m, Array.Empty<AgSesionCobertura>(),
                new[] { sesion });
            Assert.True(resultado.IsSuccess, resultado.IsFailure ? resultado.Error : string.Empty);
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

        // Decisión humana 2026-11-15 «B: extras como ocupación fija»: una relación productiva declarada
        // como extra no requerida se conserva inmutable, ocupa aula/docente y no suma demanda requerida.
        private static AgClaveRelacionCobertura Clave(int aula, int asignatura)
            => new AgClaveRelacionCobertura(1, "VALLE", 1, 1, 1, aula, asignatura, "I", "A", "PRESENCIAL");

        private static AgManifiestoCobertura ManifiestoConExtras(int aulas, params AgClaveRelacionCobertura[] extras)
            => new AgManifiestoCobertura(new DateTime(2027, 2, 5), new DateTime(2027, 5, 30),
                Enumerable.Range(1, aulas).Select(a => new AgRelacionCobertura(Clave(a, a), 4, 4, 1)), extras);

        private static AgSesionCobertura SesionDe(AgClaveRelacionCobertura clave, string origen = "NUEVA",
            long id = 1, int docente = 9, int inicio = 8) => new AgSesionCobertura(clave, origen, id, docente,
                new DateTime(2027, 2, 5), TimeSpan.FromHours(inicio), TimeSpan.FromHours(inicio + 4), 4);

        [Fact]
        public void ExtraDeclaradaConservadaNoBloqueaNiSumaDemandaRequerida()
        {
            var extra = Clave(1, 2);
            var m = ManifiestoConExtras(1, extra);
            var r = new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { SesionDe(extra, "CONSERVADA", 77, 9, 14), Sesion(m, origen: "CONSERVADA") },
                Array.Empty<AgSesionCobertura>());
            Assert.True(r.IsSuccess, r.IsFailure ? r.Error : string.Empty);
        }

        [Fact]
        public void ExtraNoCubreLaDemandaRequerida()
        {
            var extra = Clave(1, 2);
            var m = ManifiestoConExtras(1, extra);
            var r = new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { SesionDe(extra, "CONSERVADA", 77, 9, 14) }, Array.Empty<AgSesionCobertura>());
            Assert.StartsWith("COBERTURA_INCOMPLETA:", r.Error);
        }

        [Fact]
        public void ExtraDeclaradaComoNuevaEsInesperada()
        {
            var extra = Clave(1, 2);
            var m = ManifiestoConExtras(1, extra);
            Assert.Equal("RELACION_INESPERADA", new AgValidadorCoberturaSeleccion().Validar(m,
                Array.Empty<AgSesionCobertura>(), new[] { SesionDe(extra, "NUEVA", 77, 9, 14) }).Error);
        }

        [Fact]
        public void ExtraConservadaBloqueaCruceDeAulaYDeDocente()
        {
            var extra = Clave(1, 2);
            var m = ManifiestoConExtras(2, extra);
            Assert.Equal("CRUCE_AULA", new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { SesionDe(extra, "CONSERVADA", 77, 9, 8) },
                new[] { Sesion(m, aula: 1, id: 2, docente: 8, inicio: 9) }).Error);
            Assert.Equal("CRUCE_DOCENTE", new AgValidadorCoberturaSeleccion().Validar(m,
                new[] { SesionDe(extra, "CONSERVADA", 77, 9, 8) },
                new[] { Sesion(m, aula: 2, id: 2, docente: 9, inicio: 8) }).Error);
        }

        [Fact]
        public void ProyectorConservaExtrasInmutablesSinReplanificarlas()
        {
            // Requisito 5: el proyector proyecta las extras como conservadas inmutables (ocupan aula y
            // docente) y nunca las inserta en staging ni producción. Se cubre aquí porque el proyector no
            // tiene archivo de pruebas entre las superficies autorizadas.
            var extra = Clave(1, 2);
            var m = new AgManifiestoCobertura(new DateTime(2027, 2, 5), new DateTime(2027, 5, 30),
                new[] { new AgRelacionCobertura(Clave(1, 1), 4, 4, (int)AgEstadoAula.Aprobado) }, new[] { extra });
            var requerida = m.Relaciones.Single().Clave;
            var snapshot = new AgPlanAprobadoSnapshot(m,
                new[] { new AgRelacionPlanLeida(10, new[] { requerida }, new[] { Canonica(1, 10, 9, 8) }, "v10") },
                new string[0], "catalogo",
                new[] { new AgRelacionPlanLeida(11, new[] { extra }, new[] { Canonica(2, 11, 9, 14) }, "v11") });
            var contexto = new AgContexto
            {
                ManifiestoCobertura = m,
                PlanAprobadoSnapshot = snapshot,
                ValidacionPlanAprobado = new AgValidacionPlanAprobado(snapshot.HuellaOrigen,
                    new[] { new AgValidacionAulaPlan(1, AgEstadoValidacionPlan.RequiereRatificacion,
                        new[] { new AgValidacionRelacionPlan(m.Relaciones.Single(), 4, 0, new string[0]) }) },
                    new string[0]),
                Ventana = new VentanaSemestre { FechaInicio = m.Inicio, FechaFin = m.Fin }
            };
            contexto.RegistrarCandidatoCobertura(new AgCandidatoPlanAprobado(Guid.NewGuid(), Guid.NewGuid(), 1,
                "actor", "motivo", m, snapshot.HuellaOrigen, "previa", "MAR-VIE", "MAR-VIE", new int[0]));
            var r = new AgProyectorSolucionCobertura().Conservar(contexto);
            Assert.True(r.IsSuccess, r.IsFailure ? r.Error : string.Empty);
            Assert.Equal(2, r.Value.Count);
            var conservadaExtra = r.Value.Single(s => s.Clave.Valor == extra.Valor);
            Assert.Equal("CONSERVADA", conservadaExtra.Origen);
            Assert.Equal(2, conservadaExtra.Id);
            Assert.Equal(14, conservadaExtra.Inicio.TotalHours);
        }

        private static AgSesionCanonica Canonica(long id, long vinculacion, long docente, int inicio)
            => new AgSesionCanonica(id, vinculacion, docente, new DateTime(2027, 2, 5),
                TimeSpan.FromHours(inicio), TimeSpan.FromHours(inicio + 4), 4, "VALLE", "sesion:" + id);
    }
}
