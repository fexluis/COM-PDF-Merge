using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>Regla general: toda aula con relaciones requeridas y horas es demanda obligatoria.
    /// Una aprobada sin fila productiva para su cohorte se planifica como nueva; nunca se elimina,
    /// se inactiva ni se bloquea por el estado del aula.</summary>
    public class AgAprobadaSinOrigenTests
    {
        private static readonly DateTime Inicio = new DateTime(2027, 2, 5);
        private static readonly DateTime Fin = new DateTime(2027, 5, 30);
        private const int AulaId = 13;

        private static AgClaveRelacionCobertura Clave(int asignatura = 1)
            => new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, AulaId, asignatura, "I", "A", "PRESENCIAL");

        private static AgAula Aula(AgEstadoAula estado = AgEstadoAula.Aprobado)
            => new AgAula
            {
                Id = AulaId,
                IdPeriodo = 3,
                IdSede = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                ModalidadVinculacion = "Presencial",
                AulaRef = "3|1|1|I|A|PRESENCIAL",
                Estado = estado,
                Jornada = "MANANA",
                HorasDia = 4,
                DiasVentana = "MAR-VIE",
                Subunidad = "VALLE"
            };

        private static AgManifiestoCobertura Manifiesto(AgEstadoAula estado = AgEstadoAula.Aprobado, int asignatura = 1)
            => new AgManifiestoCobertura(Inicio, Fin,
                new[] { new AgRelacionCobertura(Clave(asignatura), 4, 4, (int)estado) });

        private static AgSesionCanonica Sesion(long id = 1)
            => new AgSesionCanonica(id, 10, 9, Inicio, TimeSpan.FromHours(8), TimeSpan.FromHours(12), 4,
                "VALLE", "s:" + id);

        private static AgPlanAprobadoSnapshot Snapshot(AgManifiestoCobertura manifiesto,
            params AgRelacionPlanLeida[] relaciones)
            => new AgPlanAprobadoSnapshot(manifiesto, relaciones, new string[0], "catalogo");

        private static AgContexto Contexto(AgManifiestoCobertura manifiesto, AgPlanAprobadoSnapshot snapshot,
            AgEstadoAula estado = AgEstadoAula.Aprobado)
            => new AgContexto
            {
                Aulas = new List<AgAula> { Aula(estado) },
                ManifiestoCobertura = manifiesto,
                PlanAprobadoSnapshot = snapshot
            };

        [Fact]
        public void TieneOrigenConservable_WhenAprobadaSinFilasProductivas_ShouldSerFalso()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));

            sut.TieneOrigenConservable(AulaId).Should().BeFalse();
        }

        [Fact]
        public void TieneOrigenConservable_WhenExisteVinculacionDeLaCohorte_ShouldSerVerdadero()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m, new AgRelacionPlanLeida(10, new[] { Clave() }, new[] { Sesion() }, "v")));

            sut.TieneOrigenConservable(AulaId).Should().BeTrue();
        }

        [Fact]
        public void PuedeGenerarAula_WhenAprobadaSinOrigen_ShouldPermitirGeneracion()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));

            sut.PuedeGenerarAula(AulaId).Should().BeTrue("no existe origen que conservar y la demanda es obligatoria");
        }

        [Fact]
        public void PuedeGenerarAula_WhenAprobadaConOrigenSinCandidato_ShouldDenegarGeneracion()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m, new AgRelacionPlanLeida(10, new[] { Clave() }, new[] { Sesion() }, "v")));

            sut.PuedeGenerarAula(AulaId).Should().BeFalse("un origen existente solo se replanifica con ratificación");
        }

        [Fact]
        public void PuedeGenerarAula_WhenDisponible_ShouldPermitirGeneracion()
        {
            var m = Manifiesto(AgEstadoAula.Disponible);
            var sut = Contexto(m, Snapshot(m), AgEstadoAula.Disponible);

            sut.PuedeGenerarAula(AulaId).Should().BeTrue();
        }

        [Fact]
        public void Validar_WhenAprobadaSinOrigen_ShouldSerPendienteDeProgramar()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));

            var result = new AgValidadorPlanAprobado().Validar(sut.PlanAprobadoSnapshot, sut);

            result.Aulas.Single().Estado.Should().Be(AgEstadoValidacionPlan.PendienteDeProgramar);
            result.BloqueaPreparacion.Should().BeFalse();
            result.Aulas.Single().Relaciones.Single().Diagnosticos
                .Should().Contain(d => d.StartsWith("RELACION_AUSENTE:"));
        }

        [Fact]
        public void Validar_WhenAprobadaConOrigenParcial_ShouldSeguirPlanNoVerificable()
        {
            var m = new AgManifiestoCobertura(Inicio, Fin, new[]
            {
                new AgRelacionCobertura(Clave(1), 4, 4, (int)AgEstadoAula.Aprobado),
                new AgRelacionCobertura(Clave(2), 4, 4, (int)AgEstadoAula.Aprobado)
            });
            var sut = Contexto(m, Snapshot(m,
                new AgRelacionPlanLeida(10, new[] { Clave(1) }, new[] { Sesion() }, "v")));

            var result = new AgValidadorPlanAprobado().Validar(sut.PlanAprobadoSnapshot, sut);

            result.Aulas.Single().Estado.Should().Be(AgEstadoValidacionPlan.PlanNoVerificable);
            result.BloqueaPreparacion.Should().BeTrue();
        }

        [Fact]
        public void Conservar_WhenSoloHayAprobadasSinOrigen_ShouldNoBloquearNiConservar()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));
            sut.ValidacionPlanAprobado = new AgValidadorPlanAprobado().Validar(sut.PlanAprobadoSnapshot, sut);

            var result = new AgProyectorSolucionCobertura().Conservar(sut);

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }

        [Fact]
        public void ValidarAprobadas_WhenAprobadaSinOrigen_ShouldPermitirLaCorrida()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));

            AgGeneradorHorarios.ValidarAprobadas(sut).IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void ValidarAprobadas_WhenAprobadaConOrigenSinRatificacion_ShouldFallarCerrado()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m, new AgRelacionPlanLeida(10, new[] { Clave() }, new[] { Sesion() }, "v")));

            AgGeneradorHorarios.ValidarAprobadas(sut).IsFailure.Should().BeTrue();
        }
    }
}
