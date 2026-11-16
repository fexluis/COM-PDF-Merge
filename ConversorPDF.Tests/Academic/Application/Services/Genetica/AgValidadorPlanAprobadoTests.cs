using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    // Escritas bajo S8: no ejecutadas. API pública, sin SQL ni acceso a almacenamiento.
    public class AgValidadorPlanAprobadoTests
    {
        private static readonly DateTime Inicio = new DateTime(2027, 2, 5);
        private static readonly DateTime Fin = new DateTime(2027, 5, 30);
        private static AgClaveRelacionCobertura Clave(int aula = 14, int asignatura = 7)
            => new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, aula, asignatura, "I", "B", "PRESENCIAL");
        private static AgManifiestoCobertura Manifiesto(params AgClaveRelacionCobertura[] claves)
            => new AgManifiestoCobertura(Inicio, Fin, claves.Select(k =>
                new AgRelacionCobertura(k, 4, 4, (int)AgEstadoAula.Aprobado)));
        private static AgSesionCanonica Sesion(long id = 1, DateTime? fecha = null, double horas = 4,
            double inicio = 8, double fin = 12, long docente = 9, long vinculacion = 10)
            => new AgSesionCanonica(id, vinculacion, docente, fecha ?? Inicio, TimeSpan.FromHours(inicio),
                TimeSpan.FromHours(fin), horas, "VALLE", "sesion:" + id);
        private static AgRelacionPlanLeida Relacion(AgClaveRelacionCobertura clave, params AgSesionCanonica[] sesiones)
            => new AgRelacionPlanLeida(10, new[] { clave }, sesiones, "vinculacion");
        private static AgValidacionPlanAprobado Validar(AgManifiestoCobertura m, AgRelacionPlanLeida[] r,
            string[] diagnosticos = null, AgContexto contexto = null)
            => new AgValidadorPlanAprobado().Validar(new AgPlanAprobadoSnapshot(m, r,
                diagnosticos ?? new string[0], "catalogo"), contexto ?? new AgContexto());

        [Fact]
        public void CompletaDentroDelLimite_NoAutorizaConservarNiReplanificar()
        {
            var k = Clave(); var s = Sesion(fecha: Fin);
            var r = Validar(Manifiesto(k), new[] { Relacion(k, s) });
            Assert.Equal(AgEstadoValidacionPlan.RequiereRatificacion, r.Aulas.Single().Estado);
            var relacion = r.Aulas.Single().Relaciones.Single();
            Assert.Equal(4, relacion.HorasEsperadas);
            Assert.Equal(4, relacion.HorasValidasEnVentana);
            Assert.Equal(0, relacion.HorasPreservadas);
            Assert.Equal(0, relacion.Deficit);
            Assert.Equal(0, relacion.Exceso);
            Assert.True(r.BloqueaPreparacion);
            Assert.Equal(Fin, s.Fecha);
            Assert.Equal(9, s.IdTercero);
        }

        [Fact]
        public void CeroSesionesONingunaVinculacion_NoEsPlanCompleto()
        {
            var k = Clave(); var m = Manifiesto(k);
            foreach (var relaciones in new[] { new AgRelacionPlanLeida[0], new[] { Relacion(k) } })
            {
                var r = Validar(m, relaciones).Aulas.Single();
                Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, r.Estado);
                Assert.Equal(4, r.Relaciones.Single().Deficit);
                Assert.Equal(0, r.Relaciones.Single().HorasPreservadas);
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(115)] // 31/05/2027: no extender hasta junio.
        public void FueraDeVentana_NoCuentaCobertura(int dias)
        {
            var k = Clave(); var r = Validar(Manifiesto(k), new[] { Relacion(k, Sesion(fecha: Inicio.AddDays(dias))) });
            var d = r.Aulas.Single().Relaciones.Single();
            Assert.Equal(4, d.HorasFueraVentana);
            Assert.Equal(4, d.Deficit);
            Assert.Contains("SESION_FUERA_DE_VENTANA:1", d.Diagnosticos);
            Assert.Equal(Fin, Manifiesto(k).Fin);
        }

        [Theory]
        [InlineData("duracion", "DURACION_INVALIDA:1")]
        [InlineData("festivo", "SESION_EN_FESTIVO:1")]
        [InlineData("duplicado", "SESION_DUPLICADA:1")]
        [InlineData("vinculo", "IDENTIDAD_SESION_INVALIDA:1")]
        [InlineData("docente", "SESION_NO_VERIFICABLE:1")]
        public void SesionesInvalidas_NoSumanHoras(string caso, string codigo)
        {
            var k = Clave(); var c = new AgContexto();
            if (caso == "festivo") c.Festivos.Add((Inicio, Inicio));
            var s = Sesion(fin: caso == "duracion" ? 10 : 12, docente: caso == "docente" ? 0 : 9,
                vinculacion: caso == "vinculo" ? 99 : 10);
            var sesiones = caso == "duplicado" ? new[] { s, s } : new[] { s };
            var r = Validar(Manifiesto(k), new[] { Relacion(k, sesiones) }, contexto: c).Aulas.Single();
            Assert.Contains(codigo, r.Relaciones.Single().Diagnosticos);
            Assert.Equal(0, r.Relaciones.Single().HorasValidasEnVentana);
            Assert.Equal(4, r.Relaciones.Single().Deficit);
            Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, r.Estado);
        }

        [Fact]
        public void ExcesoYDeficit_SeInformanPorRelacionSinCompensar()
        {
            var k = Clave(); var otra = Clave(asignatura: 8);
            var r = Validar(Manifiesto(k, otra), new[] { Relacion(k, Sesion(), Sesion(2, Inicio.AddDays(1))) });
            Assert.Equal(4, r.Aulas.Single().Relaciones.Single(x => x.Clave.Valor == k.Valor).Exceso);
            Assert.Equal(4, r.Aulas.Single().Relaciones.Single(x => x.Clave.Valor == otra.Valor).Deficit);
            Assert.True(r.BloqueaPreparacion);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void CrucesExactosDeAulaODocente_BloqueanAmbasRelaciones(bool mismaAula)
        {
            var k = Clave(); var otra = Clave(mismaAula ? 14 : 13, 8);
            var r = Validar(Manifiesto(k, otra), new[] { Relacion(k, Sesion()),
                new AgRelacionPlanLeida(11, new[] { otra }, new[] { Sesion(2, inicio: 11, fin: 15,
                    docente: mismaAula ? 8 : 9, vinculacion: 11) }, "otra") });
            Assert.All(r.Aulas.SelectMany(a => a.Relaciones), d =>
                Assert.Contains(d.Diagnosticos, x => x.StartsWith(mismaAula ? "CRUCE_AULA:" : "CRUCE_DOCENTE:")));
            Assert.All(r.Aulas, a => Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, a.Estado));
        }

        [Fact]
        public void IntervalosContiguos_NoSeSolapan()
        {
            var k = Clave();
            var r = Validar(Manifiesto(k), new[] { Relacion(k, Sesion(), Sesion(2, inicio: 12, fin: 16)) });
            Assert.DoesNotContain(r.Aulas.Single().Relaciones.Single().Diagnosticos, x => x.StartsWith("CRUCE_"));
        }

        [Theory]
        [InlineData("RELACION_EXTRA:99")]
        [InlineData("SESION_HUERFANA_ORIGEN_DESCONOCIDO:99")]
        [InlineData("IDENTIDAD_AMBIGUA:10")]
        public void IncertidumbreDelLector_NoSeOcultaAunqueHorasCoincidan(string codigo)
        {
            var k = Clave(); var r = Validar(Manifiesto(k), new[] { Relacion(k, Sesion()) }, new[] { codigo });
            Assert.Contains(codigo, r.Diagnosticos);
            Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, r.Aulas.Single().Estado);
        }

        [Fact]
        public void ExtrasYClavesDesconocidasSinDiagnosticoDelLector_TambienBloquean()
        {
            var k = Clave();
            foreach (var claves in new[] { new AgClaveRelacionCobertura[0], new[] { Clave(asignatura: 99) }, new[] { k, k } })
            {
                var r = Validar(Manifiesto(k), new[] { new AgRelacionPlanLeida(10, claves, new[] { Sesion() }, "extra") });
                Assert.Equal(AgEstadoValidacionPlan.PlanNoVerificable, r.Aulas.Single().Estado);
                Assert.Contains(r.Diagnosticos, d => d.StartsWith("RELACION_NO_RECONCILIADA:"));
            }
        }

        [Fact]
        public void Disponible_NoNecesitaCertificarOrigenCanonico()
        {
            var m = new AgManifiestoCobertura(Inicio, Fin, new[] { new AgRelacionCobertura(Clave(), 4, 4, 1) });
            var r = Validar(m, new AgRelacionPlanLeida[0], new[] { "INCERTIDUMBRE_AJENA" });
            Assert.False(r.BloqueaPreparacion);
            Assert.Equal(AgEstadoValidacionPlan.NoRequerida, r.Aulas.Single().Estado);
        }
    }
}
