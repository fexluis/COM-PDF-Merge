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
            var ctx = CrearContextoConPatrones();
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
            var ctx = CrearContextoConPatrones();
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
            var ctx = CrearContextoConPatrones();
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

        [Theory]
        [InlineData(AgEstadoAula.Aprobado)]
        [InlineData(AgEstadoAula.Inactiva)]
        [InlineData(AgEstadoAula.Mantenimiento)]
        [InlineData(AgEstadoAula.Reservada)]
        public void IdentificarPatrones_WhenAulaNoElegible_ShouldNoGenerarCandidatos(AgEstadoAula estadoNoElegible)
        {
            // Arrange
            var ctx = CrearContextoConPatrones();
            var aula = ctx.Aulas[0];
            aula.Estado = estadoNoElegible;
            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var candidatos = sut.IdentificarPatronesRedistribuibles(ctx, aula);
            var candidatosTodos = sut.IdentificarPatronesRedistribuibles(ctx);

            // Assert
            candidatos.Should().BeEmpty();
            candidatosTodos.Should().BeEmpty();
        }

        [Fact]
        public void ValidarSelecciones_WhenAulaEstadoAprobado_ShouldRetornarRechazoAulaNoElegible()
        {
            // Arrange
            var ctx = CrearContextoConPatrones();
            var aula = ctx.Aulas[0];
            aula.Estado = AgEstadoAula.Aprobado;
            var sut = new AgEscaladorRanuraAula(() => null);

            var rec = new AgReclasificacionPatron
            {
                IdAula = aula.Id,
                AulaRef = aula.AulaRef,
                Subunidad = aula.Subunidad ?? string.Empty,
                IdAsignatura = 1,
                NombreAsignatura = "Asig 1",
                OrigenAnterior = "LEGADO",
                OrigenPropuesto = "AUTOMATICO",
                DiasPatronAnterior = "5",
                DiasPatronPropuesto = string.Empty,
                SesionesPendientes = 2,
                HorasPendientes = 8,
                Autorizado = true
            };

            // Act
            var validacion = sut.ValidarSelecciones(ctx, new[] { rec });
            Action act = () => sut.Evaluar(ctx, 0, new[] { rec });

            // Assert
            validacion.EsValido.Should().BeFalse();
            validacion.Motivo.Should().Be("AULA_NO_ELEGIBLE");
            validacion.IdAulaRechazada.Should().Be(aula.Id);
            act.Should().Throw<ArgumentException>()
                .WithMessage($"*{aula.AulaRef}*");
        }

        [Fact]
        public void IdentificarYValidar_WhenDosAulasMismaMateriaDistintoEstado_ShouldAislarDisponibleDeAprobada()
        {
            // Arrange
            var ctx = CrearContextoConPatrones();
            var aulaDisponible = ctx.Aulas[0];
            aulaDisponible.Estado = AgEstadoAula.Disponible;

            var aulaAprobada = new AgAula
            {
                Id = 2,
                AulaRef = "A2",
                Subunidad = "VALLE",
                IdPrograma = 1,
                Semestre = "I",
                DiasVentana = "LUN-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Aprobado
            };
            ctx.Aulas.Add(aulaAprobada);
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 2,
                IdAsignatura = 1,
                HorasObjetivo = 16,
                OrigenPatron = "LEGADO",
                DiasPatron = "5"
            });
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act
            var candidatosDisponible = sut.IdentificarPatronesRedistribuibles(ctx, aulaDisponible);
            var candidatosAprobada = sut.IdentificarPatronesRedistribuibles(ctx, aulaAprobada);
            var todosCandidatos = sut.IdentificarPatronesRedistribuibles(ctx);

            // Assert candidatos
            candidatosDisponible.Should().HaveCount(1);
            candidatosDisponible[0].IdAula.Should().Be(aulaDisponible.Id);
            candidatosAprobada.Should().BeEmpty();
            todosCandidatos.Should().HaveCount(1);
            todosCandidatos[0].IdAula.Should().Be(aulaDisponible.Id);

            // Autorizar candidato de aula disponible
            var rec = candidatosDisponible[0].Clone();
            rec.Autorizado = true;

            var validacion = sut.ValidarSelecciones(ctx, new[] { rec });
            validacion.EsValido.Should().BeTrue();

            string huellaAprobadaAntes = AgHuellaPlanVentana.Calcular(ctx, aulaAprobada);
            var decisiones = sut.Evaluar(ctx, 0, new[] { rec });

            // Assert decisiones y aislamiento
            var decisionDisponible = decisiones.SingleOrDefault(d => d.IdAula == aulaDisponible.Id);
            decisionDisponible.Should().NotBeNull();
            decisionDisponible.ReclasificacionesAutorizadas.Should().HaveCount(1);

            decisiones.Any(d => d.IdAula == aulaAprobada.Id).Should().BeFalse();
            ctx.AulaAsignaturaPorAulaAsig[(2, 1)].OrigenPatron.Should().Be("LEGADO");
            ctx.AulaAsignaturaPorAulaAsig[(2, 1)].DiasPatron.Should().Be("5");
            AgHuellaPlanVentana.Calcular(ctx, aulaAprobada).Should().Be(huellaAprobadaAntes);
        }

        [Fact]
        public void ValidarSelecciones_WhenDosProgramasHomonimosMismaSede_ShouldMantenerIdentidadesAisladas()
        {
            // Arrange
            var fechaFin = new DateTime(2027, 2, 15);
            var ctx = new AgContexto
            {
                IdSede = 1,
                Subunidad = "SEDE_CENTRAL",
                Ventana = VentanaSemestre.Crear(1, "2027-01", new DateTime(2027, 2, 1),
                    fechaFin, AgOrigenVentana.Presencial).Value
            };
            ctx.Sedes.Add(new Sede { Id = 1, Ciudad = "SEDE CENTRAL" });

            // Dos programas con el mismo nombre pero diferente IdPrograma
            ctx.Programas.Add(new Programa { Id = 1, Nombre = "ADMINISTRACION PUBLICA", Modalidad = "PREGRADO" });
            ctx.Programas.Add(new Programa { Id = 2, Nombre = "ADMINISTRACION PUBLICA", Modalidad = "PREGRADO" });

            // Asignatura para cada programa con Id distinto
            ctx.Asignaturas.Add(new Asignatura { Id = 10, IdPrograma = 1, Nombre = "MATEMATICAS", Semestre = "I", IntHoraria = 16 });
            ctx.Asignaturas.Add(new Asignatura { Id = 20, IdPrograma = 2, Nombre = "MATEMATICAS", Semestre = "I", IntHoraria = 16 });

            var aula1 = new AgAula
            {
                Id = 1,
                AulaRef = "A1",
                Subunidad = "SEDE_CENTRAL",
                IdSede = 1,
                IdPrograma = 1,
                Semestre = "I",
                DiasVentana = "LUN-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            };
            var aula2 = new AgAula
            {
                Id = 2,
                AulaRef = "A2",
                Subunidad = "SEDE_CENTRAL",
                IdSede = 1,
                IdPrograma = 2,
                Semestre = "I",
                DiasVentana = "LUN-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible
            };
            ctx.Aulas.Add(aula1);
            ctx.Aulas.Add(aula2);

            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 10,
                HorasObjetivo = 16,
                DiasPatron = "5",
                OrigenPatron = "LEGADO"
            });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 2,
                IdAsignatura = 20,
                HorasObjetivo = 16,
                DiasPatron = "5",
                OrigenPatron = "LEGADO"
            });

            for (int dia = 1; dia <= 5; dia++)
            {
                ctx.Bloques.Add(new AgBloqueHorario
                {
                    Id = dia,
                    DiaSemanaNumero = dia,
                    HoraInicio = "08:00",
                    HoraFin = "12:00",
                    DuracionHoras = 4
                });
            }
            ctx.ConstruirIndices();

            var sut = new AgEscaladorRanuraAula(() => null);

            // Act 1: Obtener candidatas
            var candidatosAula1 = sut.IdentificarPatronesRedistribuibles(ctx, aula1);
            var candidatosAula2 = sut.IdentificarPatronesRedistribuibles(ctx, aula2);

            // Assert 1: Identidades estrictas
            candidatosAula1.Should().HaveCount(1);
            candidatosAula1[0].IdAula.Should().Be(1);
            candidatosAula1[0].IdAsignatura.Should().Be(10);
            candidatosAula1[0].Programa.Should().Be("ADMINISTRACION PUBLICA");

            candidatosAula2.Should().HaveCount(1);
            candidatosAula2[0].IdAula.Should().Be(2);
            candidatosAula2[0].IdAsignatura.Should().Be(20);
            candidatosAula2[0].Programa.Should().Be("ADMINISTRACION PUBLICA");

            // Act 2: Cruzar selección (asignar Asignatura 20 a Aula 1)
            var recCruzado = candidatosAula1[0].Clone();
            recCruzado.IdAsignatura = 20;
            recCruzado.Autorizado = true;

            var valCruzado = sut.ValidarSelecciones(ctx, new[] { recCruzado });

            // Assert 2: Rechazado por CONFIGURACION_MODIFICADA
            valCruzado.EsValido.Should().BeFalse();
            valCruzado.Motivo.Should().Be("CONFIGURACION_MODIFICADA");

            // Act 3: Intentar asociar Aula 1 con AulaRef ajeno de Aula 2
            var recRefCruzada = candidatosAula1[0].Clone();
            recRefCruzada.AulaRef = aula2.AulaRef;
            recRefCruzada.Autorizado = true;

            var valRefCruzada = sut.ValidarSelecciones(ctx, new[] { recRefCruzada });

            // Assert 3: Rechazado por AULA_NO_ELEGIBLE (no existe combinación)
            valRefCruzada.EsValido.Should().BeFalse();
            valRefCruzada.Motivo.Should().Be("AULA_NO_ELEGIBLE");

            // Act 4: Autorización legítima individual de Aula 1
            var recAula1 = candidatosAula1[0].Clone();
            recAula1.Autorizado = true;
            var valValida = sut.ValidarSelecciones(ctx, new[] { recAula1 });
            var decisiones = sut.Evaluar(ctx, 0, new[] { recAula1 });

            // Assert 4: Válido y decisiones aisladas
            valValida.EsValido.Should().BeTrue();
            var decAula1 = decisiones.Single(d => d.IdAula == 1);
            var decAula2 = decisiones.Single(d => d.IdAula == 2);
            decAula1.ReclasificacionesAutorizadas.Should().HaveCount(1);
            decAula1.ReclasificacionesAutorizadas[0].IdAsignatura.Should().Be(10);
            decAula2.ReclasificacionesAutorizadas.Should().BeEmpty();
        }

        [Fact]
        public async Task Aplicar_WhenCanceladoORechazado_ShouldTenerCeroEscrituras()
        {
            // Arrange
            var ctx = CrearContextoConPatrones();
            var sut = new AgEscaladorRanuraAula(() => null);
            var rec = sut.IdentificarPatronesRedistribuibles(ctx, ctx.Aulas[0]).Single();
            rec.Autorizado = true;
            var decisiones = sut.Evaluar(ctx, 0, new[] { rec });

            // SUT con factory que fallaría si se intentara abrir conexión a BD
            var sutSinDb = new AgEscaladorRanuraAula(() => throw new InvalidOperationException("No debe abrir conexion a BD"));

            // Act 1: Cancelación vía CancellationToken
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                var resultadoCancelado = await sutSinDb.AplicarDetalladoAsync(decisiones, 0, cts.Token);

                // Assert 1: Cancelado y cero escrituras
                resultadoCancelado.Exito.Should().BeFalse();
                resultadoCancelado.Error.Should().Be("Operación cancelada.");
                resultadoCancelado.Modificadas.Should().Be(0);
                resultadoCancelado.Confirmadas.Should().Be(0);
            }

            // Act 2: Rechazo por huella corrupta o manipulada
            var decisionManipulada = decisiones[0];
            decisionManipulada.HuellaDecision = "HUELLA_ALTERADA";
            var resultadoRechazado = await sutSinDb.AplicarDetalladoAsync(new[] { decisionManipulada }, 0, CancellationToken.None);

            // Assert 2: Falla controlada y cero escrituras
            resultadoRechazado.Exito.Should().BeFalse();
            resultadoRechazado.Error.Should().Contain("modificada");
            resultadoRechazado.Modificadas.Should().Be(0);
            resultadoRechazado.Confirmadas.Should().Be(0);
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

        private static AgContexto Contexto(DateTime? fin = null)
        {
            var fechaFin = fin ?? new DateTime(2027, 5, 30);
            var ctx = new AgContexto
            {
                Ventana = VentanaSemestre.Crear(1, "2027-01", new DateTime(2027, 2, 1),
                    fechaFin, AgOrigenVentana.Presencial).Value
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

        private static AgContexto CrearContextoConPatrones(DateTime? fin = null)
        {
            // Ventana acotada a 2 semanas: 2 viernes disponibles vs 4 sesiones requeridas (16h) = 2 sesiones pendientes (déficit 8h)
            return Contexto(fin ?? new DateTime(2027, 2, 15));
        }
    }
}
