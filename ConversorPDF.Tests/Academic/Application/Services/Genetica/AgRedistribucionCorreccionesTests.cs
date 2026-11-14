using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgRedistribucionCorreccionesTests
    {
        [Theory]
        [InlineData("LEGADO", "AUTOMATICO", "", true)]
        [InlineData("LEGADO", "FIJO", "5", true)]
        [InlineData("LEGADO", "FIJO", "7", false)]
        [InlineData("LEGADO", "AUTOMATICO", "5", false)]
        [InlineData("AUTOMATICO", "FIJO", "5", false)]
        [InlineData("FIJO", "AUTOMATICO", "", false)]
        public void Reclasificacion_ValidaOrigenYConservacionDePatron(string antes, string despues, string dias, bool valida)
        {
            var rec = new AgReclasificacionPatron
            {
                IdAula = 1, AulaRef = "A1", IdAsignatura = 1,
                OrigenAnterior = antes, OrigenPropuesto = despues,
                DiasPatronAnterior = "5", DiasPatronPropuesto = dias
            };
            rec.EsTransicionValida(out _).Should().Be(valida);
        }

        [Theory]
        [InlineData("duplicada")]
        [InlineData("referencia")]
        [InlineData("subunidad")]
        [InlineData("origen")]
        [InlineData("patron")]
        [InlineData("relacion")]
        [InlineData("sinIdentidad")]
        public void Evaluar_RechazaAutorizacionInvalidaSinAlterarContexto(string caso)
        {
            var ctx = Contexto();
            var sut = new AgEscaladorRanuraAula(() => null);
            var rec = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]).Single();
            rec.Autorizado = true;
            var seleccion = new List<AgReclasificacionPatron> { rec };
            switch (caso)
            {
                case "duplicada": seleccion.Add(rec.Clone()); break;
                case "referencia": rec.AulaRef = "OTRA"; break;
                case "subunidad": rec.Subunidad = "OTRA"; break;
                case "origen": rec.OrigenAnterior = "FIJO"; break;
                case "patron": rec.DiasPatronAnterior = "7"; break;
                case "relacion": rec.IdAsignatura = 999; break;
                case "sinIdentidad": rec.IdAula = 0; break;
            }
            Action evaluar = () => sut.Evaluar(ctx, 0, seleccion);
            evaluar.Should().Throw<ArgumentException>();
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron.Should().Be("LEGADO");
        }

        [Fact]
        public void Evaluar_AutorizaSoloAulaSeleccionadaYCopiaLaSeleccion()
        {
            var ctx = Contexto();
            ctx.Aulas.Add(new AgAula
            {
                Id = 2, AulaRef = "A2", Subunidad = "VALLE", IdPrograma = 1,
                Semestre = "I", DiasVentana = "LUN-VIE", Jornada = "MANANA",
                HorasDia = 4, Estado = AgEstadoAula.Disponible
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 2, IdAsignatura = 1, HorasObjetivo = 16,
                OrigenPatron = "LEGADO", DiasPatron = "5"
            });
            ctx.ConstruirIndices();
            var sut = new AgEscaladorRanuraAula(() => null);
            var rec = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]).Single();
            rec.Autorizado = true;
            string huellaOtra = AgHuellaPlanVentana.Calcular(ctx, ctx.Aulas[1]);
            AgHuellaPlanVentana.Calcular(ctx, ctx.Aulas[1], reclasificaciones: new[] { rec })
                .Should().Be(huellaOtra);

            var decisiones = sut.Evaluar(ctx, 0, new[] { rec });
            var elegida = decisiones.Single(d => d.AulaRef == "A1");
            elegida.Aplica.Should().BeTrue();
            elegida.Hacia.Should().Be("LUN-VIE");
            elegida.DemandaPendiente.Should().Be(0);
            decisiones.Single(d => d.AulaRef == "A2").ReclasificacionesAutorizadas.Should().BeEmpty();
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].OrigenPatron.Should().Be("LEGADO");
            ctx.AulaAsignaturaPorAulaAsig[(1, 1)].DiasPatron.Should().Be("5");
            rec.Autorizado = false;
            elegida.ReclasificacionesAutorizadas.Single().Autorizado.Should().BeTrue();
        }

        [Fact]
        public async Task Aplicar_NoOmiteUnaDecisionManipuladaAunqueDejeDeSerAplicable()
        {
            var ctx = Contexto();
            var sut = new AgEscaladorRanuraAula(() => null);
            var rec = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]).Single();
            rec.Autorizado = true;
            var decisiones = sut.Evaluar(ctx, 0, new[] { rec });
            decisiones.Single().ReclasificacionesAutorizadas.Clear();
            decisiones.Single().Aplica.Should().BeFalse();
            var resultado = await sut.AplicarDetalladoAsync(decisiones, 0, CancellationToken.None);
            resultado.Exito.Should().BeFalse();
            resultado.Error.Should().Contain("modificada");
            resultado.Modificadas.Should().Be(0);
        }

        [Fact]
        public void Sellar_CompletaHuecosAnterioresALaUltimaFecha()
        {
            var ctx = Contexto();
            var g = Gen(new DateTime(2027, 2, 15));
            g.FechaFinEstimada = new DateTime(2027, 2, 15);
            var ind = new AgIndividuo();
            ind.Genes.Add(g);
            AgCoherenciaGrupo.Sellar(ind, ctx);
            g.FechasGen.Should().Equal(new DateTime(2027, 2, 8), new DateTime(2027, 2, 15));
            g.HorasGen.Should().Be(8);
        }

        [Theory]
        [InlineData("10:00", "14:00", 1, 22)] // Solapamiento aunque cambie la etiqueta de franja.
        [InlineData("12:00", "16:00", 1, 15)] // Intervalos adyacentes.
        [InlineData("08:00", "12:00", 2, 15)] // Otra aula.
        public void Sellar_RespetaIntervalosRealesDeOtrasAsignaturas(string inicio, string fin, int otraAula, int diaEsperado)
        {
            var ctx = Contexto();
            ctx.Bloques.Add(new AgBloqueHorario
            {
                Id = 2, DiaSemanaNumero = 1, HoraInicio = inicio, HoraFin = fin,
                Franja = "OTRA_ETIQUETA", DuracionHoras = 4
            });
            ctx.ConstruirIndices();
            var g = Gen(new DateTime(2027, 2, 8));
            var otro = new AgGen
            {
                IdAula = otraAula, IdAsignatura = 2, IdBloqueHorario = 2, HorasGen = 4,
                FechasGen = new List<DateTime> { new DateTime(2027, 2, 15) }
            };
            var grupo = new AgIndividuo();
            grupo.Genes.Add(g);
            // Igual que staging: reparar un grupo consultando ocupación de todo el individuo.
            AgCoherenciaGrupo.Sellar(grupo, ctx, new[] { g, otro });
            g.FechasGen.Should().Equal(new DateTime(2027, 2, 8), new DateTime(2027, 2, diaEsperado));
            otro.FechasGen.Should().Equal(new DateTime(2027, 2, 15));
        }

        [Fact]
        public void Sellar_SinCapacidadConservaHorasYMarcaInsuficiencia()
        {
            var ctx = Contexto();
            var g = Gen(new DateTime(2027, 2, 8));
            g.FechaFinEstimada = new DateTime(2027, 2, 8);
            var ind = new AgIndividuo();
            ind.Genes.Add(g);
            AgCoherenciaGrupo.Sellar(ind, ctx);
            g.HorasGen.Should().Be(8);
            g.FechasGen.Should().HaveCount(1);
            g.Estado.Should().Be(AgEstadoGen.Incompleto);
            g.MotivoVacante.Should().Be(AgMotivoVacante.SIN_PLAN_OLEADA.Codigo());
        }

        private static AgGen Gen(DateTime primera)
        {
            return new AgGen
            {
                IdAula = 1, IdAsignatura = 1, IdBloqueHorario = 1, DiaSemana = 1,
                IdTercero = 10, HorasGen = 4, TotalHorasAsignatura = 8, HorasSemanales = 4,
                FechaInicioEstimada = new DateTime(2027, 2, 8),
                FechaFinEstimada = new DateTime(2027, 2, 22),
                FechasGen = new List<DateTime> { primera }, ClaveGrupo = "1|1|A"
            };
        }

        private static AgContexto Contexto()
        {
            var ctx = new AgContexto
            {
                Ventana = VentanaSemestre.Crear(1, "2027-01", new DateTime(2027, 2, 1),
                    new DateTime(2027, 5, 30), AgOrigenVentana.Presencial).Value
            };
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            ctx.Asignaturas.Add(new Asignatura { Id = 1, IdPrograma = 1, Semestre = "I", IntHoraria = 16 });
            ctx.Aulas.Add(new AgAula
            {
                Id = 1, AulaRef = "A1", Subunidad = "VALLE", IdPrograma = 1, Semestre = "I",
                DiasVentana = "LUN-VIE", Jornada = "MANANA", HorasDia = 4, Estado = AgEstadoAula.Disponible
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 1, IdAsignatura = 1, HorasObjetivo = 16, DiasPatron = "5", OrigenPatron = "LEGADO"
            });
            ctx.Bloques.Add(new AgBloqueHorario
            {
                Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4
            });
            for (int dia = 2; dia <= 5; dia++)
            {
                ctx.Bloques.Add(new AgBloqueHorario
                {
                    Id = 10 + dia, DiaSemanaNumero = dia,
                    HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4
                });
            }
            ctx.ConstruirIndices();
            return ctx;
        }
    }
}
