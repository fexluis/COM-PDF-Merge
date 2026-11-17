using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
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

        [Fact]
        public void PuedeGenerarAula_WhenSnapshotNuloConManifiestoAprobado_ShouldRetornarFalso()
        {
            var m = Manifiesto();
            var sut = Contexto(m, null);

            sut.PuedeGenerarAula(AulaId).Should().BeFalse("snapshot nulo no acredita ausencia de origen");
        }

        [Fact]
        public void ValidarAprobadas_WhenSnapshotNuloConManifiestoAprobado_ShouldFallarCerrado()
        {
            var m = Manifiesto();
            var sut = Contexto(m, null);

            var res = AgGeneradorHorarios.ValidarAprobadas(sut);
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SNAPSHOT_APROBADO_REQUERIDO");
        }

        [Fact]
        public void ValidarAprobadas_WhenSnapshotDivergenteConManifiesto_ShouldFallarCerrado()
        {
            var m1 = Manifiesto();
            var m2 = new AgManifiestoCobertura(Inicio, Fin,
                new[] { new AgRelacionCobertura(new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, AulaId, 99, "I", "A", "PRESENCIAL"), 4, 4, (int)AgEstadoAula.Aprobado) });
            var sut = Contexto(m1, Snapshot(m2));

            var res = AgGeneradorHorarios.ValidarAprobadas(sut);
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SNAPSHOT_DIVERGENTE");
        }

        [Fact]
        public async Task ValidarFactibilidad_WhenSoloAprobadaSinOrigen_ShouldCalcularOfertaYSerFactible()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));
            sut.Ventana = VentanaSemestre.Crear(3, "2027-1", Inicio, Fin, AgOrigenVentana.Presencial).Value;
            sut.Bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 }
            };
            sut.Asignaturas = new List<Asignatura>
            {
                new Asignatura { Id = 1, Nombre = "Fundamentos", Horas = 16, Semestre = "I" }
            };
            sut.Docentes = new List<Tercero>
            {
                new Tercero { Id = 9, Nombre = "Docente 1", Activo = 1 }
            };
            sut.AulaAsignaturas = new List<AgAulaAsignatura>
            {
                new AgAulaAsignatura { IdAula = AulaId, IdAsignatura = 1, Requerida = 1, HorasObjetivo = 16 }
            };
            sut.ConstruirIndices();

            var validador = new AgValidadorVentana();
            var res = await validador.ValidarFactibilidadAsync(sut, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.BloqueosCriticos.Should().BeEmpty();
            int ocurrenciasMartes = AgPlanificadorTemporal.ContarOcurrencias(DayOfWeek.Tuesday, Inicio, Fin, sut.EsFestivo);
            res.Value.TotalCapacidadHorariaAulas.Should().Be(ocurrenciasMartes * 4);
        }

        [Fact]
        public void EscaladorRanura_WhenAprobadaSinOrigen_ShouldSerElegibleParaPropuesta()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));
            var aula = sut.Aulas.Single();

            AgEscaladorRanuraAula.PuedeAplicarPropuesta(sut, aula).Should().BeTrue("aprobada sin origen es demanda nueva elegible");

            var sutSinSnap = Contexto(m, null);
            AgEscaladorRanuraAula.PuedeAplicarPropuesta(sutSinSnap, aula).Should().BeFalse("snapshot nulo no acredita ausencia de origen");

            var sutConOrigen = Contexto(m, Snapshot(m, new AgRelacionPlanLeida(10, new[] { Clave() }, new[] { Sesion() }, "v")));
            AgEscaladorRanuraAula.PuedeAplicarPropuesta(sutConOrigen, aula).Should().BeFalse("aprobada con origen no admite escalado de escritura");
        }

        [Fact]
        public async Task OptimizacionCompleta_WhenSoloAprobadaSinOrigen_ShouldEjecutarOptimizacionConExito()
        {
            var m = Manifiesto();
            var sut = Contexto(m, Snapshot(m));
            sut.Ventana = VentanaSemestre.Crear(3, "2027-1", Inicio, Fin, AgOrigenVentana.Presencial).Value;
            var aula = sut.Aulas.Single();
            aula.Capacidad = 30;
            sut.AulasPorId[aula.Id] = aula;

            var asig = new Asignatura { Id = 1, Nombre = "Fundamentos", Horas = 4, Semestre = "I" };
            sut.Asignaturas.Add(asig);
            sut.AsignaturasPorId[asig.Id] = asig;

            var bloque = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 };
            sut.Bloques.Add(bloque);
            sut.BloquesPorId[bloque.Id] = bloque;

            var doc = new Tercero { Id = 9, Nombre = "Docente 1", Activo = 1 };
            sut.Docentes.Add(doc);
            sut.DocentesPorId[doc.Id] = doc;

            sut.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = AulaId, IdAsignatura = 1, Requerida = 1, HorasObjetivo = 4 });
            sut.ConstruirIndices();

            var mockVentana = new Mock<IAgValidadorVentana>();
            var informe = new InformeFactibilidad(); // EsFactible = true
            mockVentana.Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(Result.Success(informe));

            var mockAptitud = new Mock<IAgEvaluadorAptitud>();
            mockAptitud.Setup(a => a.EvaluarAptitud(It.IsAny<AgIndividuo>(), It.IsAny<AgContexto>()))
                       .Callback<AgIndividuo, AgContexto>((ind, ctx) =>
                       {
                           ind.Fitness = 1.0;
                           ind.HardViolations = 0;
                       });

            var sutGenerador = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);
            var resultado = await sutGenerador.EjecutarOptimizacionAsync(sut, 4, 2, 0.8, 0.1, 1, null, CancellationToken.None);

            resultado.IsSuccess.Should().BeTrue();
            resultado.Value.EsFactible.Should().BeTrue();
            resultado.Value.MejorIndividuo.Should().NotBeNull();
            resultado.Value.MejorIndividuo.Genes.Should().NotBeEmpty()
                .And.OnlyContain(g => g.IdAula == AulaId);
        }
    }
}
