using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPlanificadorRanurasTests
    {
        [Theory]
        [InlineData(32, 64)]
        [InlineData(48, 68)]
        public void Construir_WhenDemandaCaliSinRestriccionesDocentes_ShouldCubrirHastaMayo(int quinta, int sesiones)
        {
            // Arrange
            var ctx = Crear(new[] { 32, 48, 48, 48, quinta, 48 });
            // Act
            var result = Plan(ctx);
            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();
            var fechas = result.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            fechas.Should().HaveCount(sesiones);
            fechas.Should().OnlyContain(s => s.Fecha <= new DateTime(2027, 5, 30));
            fechas.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja).Should().OnlyHaveUniqueItems();
        }

        [Fact]
        public void Construir_WhenReferenciaCabeEnVentanaAmpliada_ShouldBuscarFinMasTemprano()
        {
            // Arrange
            var ctx = Crear(new[] { 32, 48, 48, 48, 32, 48 });
            ctx.Ventana.FechaFin = new DateTime(2027, 7, 30);
            // Act
            var result = Plan(ctx);
            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();
            result.Value.Planes.SelectMany(p => p.Sesiones).Max(s => s.Fecha)
                .Should().BeOnOrBefore(new DateTime(2027, 5, 30));
            ctx.Ventana.FechaFin.Should().Be(new DateTime(2027, 7, 30));
        }

        [Fact]
        public void CrearAleatorio_WhenPlanFechado_ShouldConservarLasMismasFechas()
        {
            // Arrange
            var ctx = Crear(new[] { 28, 28, 24 });
            ctx.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx.Ventana.FechaFin = new DateTime(2027, 2, 28);
            var esperado = Plan(ctx).Value.Planes.SelectMany(p => p.Sesiones)
                .Select(s => s.IdAsignatura + "|" + s.Fecha.ToString("yyyyMMdd") + "|" + s.Ranura.Franja).ToList();
            // Act
            var individuo = new AgFabricaIndividuos(new Random(42)).CrearAleatorio(ctx, 1, 0);
            // Assert
            individuo.Genes.SelectMany(g => g.FechasGen.Select(f =>
                g.IdAsignatura + "|" + f.ToString("yyyyMMdd") + "|" + ctx.BloquesPorId[g.IdBloqueHorario].Franja))
                .Should().BeEquivalentTo(esperado);
            individuo.Genes.Sum(g => g.HorasGen).Should().Be(80);
        }

        [Theory]
        [InlineData("LUN-JUE", "LUN-VIE")]
        [InlineData("MAR-VIE", "LUN-VIE")]
        [InlineData("VIE-DOM", "JUE-DOM")]
        public void Evaluar_WhenDemanda5W_ShouldContarSesionesReales(string desde, string hacia)
        {
            // Arrange
            var ctx = Crear(new[] { 16, 16, 16, 16, 16 });
            ctx.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx.Ventana.FechaFin = new DateTime(2027, 2, 28);
            // Act
            var antes = AgEvaluadorDemandaVentana.Evaluar(desde, "TARDE", 4, ctx, ctx.Aulas[0], ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin);
            var despues = AgEvaluadorDemandaVentana.Evaluar(hacia, "TARDE", 4, ctx, ctx.Aulas[0], ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin);
            // Assert
            antes.DemandaPendiente.Should().Be(4);
            despues.CabeEnPlazo.Should().BeTrue();
            despues.DemandaPendiente.Should().Be(0);
            despues.Sesiones.Should().HaveCount(20);
        }

        [Fact]
        public void Construir_WhenFijosSoloViernes_ShouldConservarRestriccionYPendientes()
        {
            // Arrange
            var ctx = Crear(new[] { 48, 48 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var patrones = new Dictionary<int, string> { { 1, "5" }, { 2, "5" } };
            // Act
            var result = new AgPlanificadorOleadas().ConstruirEvaluacion(demandas, "LUN-VIE", ctx.Ventana.FechaInicio,
                ctx.Ventana.FechaFin, 4, ctx.EsFestivo, ctx.Bloques, "TARDE", patrones);
            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().NotBeEmpty();
            result.Value.Planes.SelectMany(p => p.Sesiones).Should().OnlyContain(s => s.Ranura.Dia == 5);
        }

        [Fact]
        public void Huella_WhenContenidoCambiaConMismoConteo_ShouldInvalidarse()
        {
            // Arrange
            var ctx = Crear(new[] { 28, 28, 24 });
            string original = AgHuellaPlanVentana.Calcular(ctx, ctx.Aulas[0]);
            // Act
            ctx.AulaAsignaturas[0].HorasObjetivo += 4;
            ctx.AulaAsignaturas[1].HorasObjetivo -= 4;
            // Assert
            AgHuellaPlanVentana.Calcular(ctx, ctx.Aulas[0]).Should().NotBe(original);
        }

        [Fact]
        public void Construir_WhenPresupuestoAgotado_ShouldDeclararBusquedaIncompleta()
        {
            // Arrange
            var ctx = Crear(new[] { 48, 48 });
            // Act
            var result = AgPlanificadorRanuras.Construir(ctx.AulaAsignaturas.Select(a =>
                (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList(),
                AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4),
                ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null, presupuesto: 1);
            // Assert
            result.BusquedaAgotada.Should().BeTrue();
            result.Infactibles.Should().NotBeEmpty();
        }

        [Theory]
        [InlineData(4, 16)]
        [InlineData(8, 12)]
        public void Evaluar_WhenUnidad4u8_ShouldDistinguirFranjasDeSabado(int unidad, int capacidad)
        {
            // Arrange
            var ctx = Crear(new[] { unidad * 4, unidad * 4, unidad * 4 });
            ctx.Programas[0].Modalidad = unidad == 8 ? "POSGRADO" : "PRESENCIAL";
            ctx.Aulas[0].HorasDia = unidad;
            ctx.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx.Ventana.FechaFin = new DateTime(2027, 2, 28);
            foreach (var bloque in ctx.Bloques) { bloque.DuracionHoras = unidad; }
            // Act
            var result = AgEvaluadorDemandaVentana.Evaluar("VIE-DOM", "TARDE", unidad, ctx,
                ctx.Aulas[0], ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin);
            // Assert
            result.CapacidadTotal.Should().Be(capacidad);
            result.Sesiones.Select(s => s.Fecha.ToString("yyyyMMdd") + s.Ranura.Franja).Should().OnlyHaveUniqueItems();
            result.Sesiones.Should().OnlyContain(s => Math.Abs(s.Ranura.Horas - unidad) < 0.001);
        }

        [Fact]
        public void CrearAleatorio_WhenDemandaCambioConMismosConteos_ShouldInvalidarCache()
        {
            // Arrange
            var ctx = Crear(new[] { 28, 28, 24 });
            var planificador = new AgPlanificadorOleadas();
            var fabrica = new AgFabricaIndividuos(new Random(42), null, planificador);
            fabrica.CrearAleatorio(ctx, 7, 0);
            int construcciones = planificador.Construcciones;
            fabrica.CrearAleatorio(ctx, 7, 1);
            planificador.Construcciones.Should().Be(construcciones);
            // Act
            ctx.AulaAsignaturas[0].HorasObjetivo += 4;
            ctx.AulaAsignaturas[1].HorasObjetivo -= 4;
            fabrica.CrearAleatorio(ctx, 7, 2);
            // Assert
            planificador.Construcciones.Should().BeGreaterThan(construcciones);
        }

        [Fact]
        public void Construir_WhenCancelado_ShouldInterrumpirBusqueda()
        {
            // Arrange
            var ctx = Crear(new[] { 48, 48 });
            using (var cancelacion = new System.Threading.CancellationTokenSource())
            {
                cancelacion.Cancel();
                // Act
                Action construir = () => new AgPlanificadorOleadas().ConstruirEvaluacion(
                    ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList(),
                    "LUN-VIE", ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, ctx.Bloques,
                    "TARDE", ct: cancelacion.Token);
                // Assert
                construir.Should().Throw<OperationCanceledException>();
            }
        }

        [Fact]
        public void EvaluarEscalador_WhenDemanda6W_ShouldProponerDosPasosSeparados()
        {
            // Arrange
            var ctx = Crear(new[] { 16, 16, 16, 16, 16, 16 });
            ctx.Ventana.FechaInicio = new DateTime(2027, 2, 1);
            ctx.Ventana.FechaFin = new DateTime(2027, 2, 28);
            ctx.Aulas[0].DiasVentana = "VIE-DOM";
            var escalador = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.AgEscaladorRanuraAula(() => null);
            // Act / Assert: cada evaluación ofrece únicamente el paso inmediato.
            var primero = escalador.Evaluar(ctx).Single();
            primero.Hacia.Should().Be("JUE-DOM");
            primero.DemandaPendiente.Should().Be(4);
            ctx.Aulas[0].DiasVentana = primero.Hacia;
            var segundo = escalador.Evaluar(ctx).Single();
            segundo.Hacia.Should().Be("MIE-DOM");
            segundo.DemandaPendiente.Should().Be(0);
            ctx.Aulas[0].DiasVentana = segundo.Hacia;
            escalador.Evaluar(ctx).Should().BeEmpty();
        }

        [Fact]
        public void Construir_WhenUltimaFechaRepresentable_ShouldNoDesbordarCalendario()
        {
            // Arrange
            DateTime fin = DateTime.MaxValue.Date;
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)> { (1, 4, 1) };
            // Act
            var result = new AgPlanificadorOleadas().Construir(demandas, "LUN-VIE", fin, fin, 4);
            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Planes.SelectMany(p => p.Sesiones).Select(s => s.Fecha).Should().Equal(fin);
        }

        [Fact]
        public void ValidarFechado_WhenPlanCompletoValido_ShouldReturnSuccess()
        {
            // Arrange
            var ctx = Crear(new[] { 16, 16 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);
            var plan = AgPlanificadorRanuras.Construir(demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null);

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            plan.Infactibles.Should().BeEmpty();
        }

        [Fact]
        public void ValidarFechado_WhenPlanParcialConInfactiblesConsistentes_ShouldReturnSuccess()
        {
            // Arrange
            var ctx = Crear(new[] { 48, 48 });
            DateTime ini = new DateTime(2027, 2, 1);
            DateTime fin = new DateTime(2027, 2, 5);
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);
            var plan = AgPlanificadorRanuras.Construir(demandas, ranuras, ini, fin, 4, ctx.EsFestivo, null);

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandas, ranuras, ini, fin, 4, ctx.EsFestivo, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            plan.Infactibles.Should().NotBeEmpty();
        }

        [Fact]
        public void ValidarFechado_WhenPlanNulo_ShouldReturnFailurePlanNulo()
        {
            // Arrange
            var ctx = Crear(new[] { 16 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(null, demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_PLAN_NULO");
        }

        [Fact]
        public void ValidarFechado_WhenHayColisionDeSesiones_ShouldReturnFailureFechaOColision()
        {
            // Arrange
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4.0, 1),
                (2, 4.0, 2)
            };
            var ranura = new AgRanura(1, "TARDE", 4.0);
            var ranuras = new List<AgRanura> { ranura };
            DateTime ini = new DateTime(2027, 2, 1);
            DateTime fin = new DateTime(2027, 2, 1);
            var sesionColisionada1 = new AgSesionPlanificada(1, ini, ranura);
            var sesionColisionada2 = new AgSesionPlanificada(2, ini, ranura);

            var oleada1 = new AgPlanOleada(1, 0, 0, new List<int> { 1 }, ini, fin, 1, 1, 1, ranuras, new List<AgSesionPlanificada> { sesionColisionada1 });
            var oleada2 = new AgPlanOleada(2, 0, 0, new List<int> { 1 }, ini, fin, 1, 1, 1, ranuras, new List<AgSesionPlanificada> { sesionColisionada2 });

            var plan = new AgPlanOleadasResultado(new List<AgPlanOleada> { oleada1, oleada2 }, new List<int>());

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandas, ranuras, ini, fin, 4.0, null, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_PLAN_FECHA_O_COLISION");
        }

        [Fact]
        public void ValidarFechado_WhenViolacionDiasDePatron_ShouldReturnFailureSesionInvalida()
        {
            // Arrange
            var demandas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 4.0, 1)
            };
            var ranuraLunes = new AgRanura(1, "TARDE", 4.0);
            var ranuras = new List<AgRanura> { ranuraLunes };
            DateTime ini = new DateTime(2027, 2, 1);
            DateTime fin = new DateTime(2027, 2, 1);
            var sesion = new AgSesionPlanificada(1, ini, ranuraLunes);
            var oleada = new AgPlanOleada(1, 0, 0, new List<int> { 1 }, ini, fin, 1, 1, 1, ranuras, new List<AgSesionPlanificada> { sesion });
            var plan = new AgPlanOleadasResultado(new List<AgPlanOleada> { oleada }, new List<int>());

            var patrones = new Dictionary<int, string> { { 1, "5" } };

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandas, ranuras, ini, fin, 4.0, null, patrones, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_PLAN_SESION_INVALIDA");
        }

        [Fact]
        public void ValidarFechado_WhenPendientesInconsistentes_ShouldReturnFailurePendientesInconsistentes()
        {
            // Arrange
            var ctx = Crear(new[] { 16 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);
            var planValido = AgPlanificadorRanuras.Construir(demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null);

            var planMentiroso = new AgPlanOleadasResultado(planValido.Planes, new List<int> { 1 });

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(planMentiroso, demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_PLAN_PENDIENTES_INCONSISTENTES");
        }

        [Fact]
        public void ValidarFechado_WhenDemandaInconsistente_ShouldReturnFailureDemandaInconsistente()
        {
            // Arrange
            var ctx = Crear(new[] { 16 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);
            var plan = AgPlanificadorRanuras.Construir(demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null);

            var demandasInconsistentes = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 32.0, 1)
            };

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandasInconsistentes, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, null, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I5_PLAN_DEMANDA_INCONSISTENTE");
        }

        [Fact]
        public void ValidarFechado_WhenPatronInvalido_ShouldReturnFailurePatronInvalido()
        {
            // Arrange
            var ctx = Crear(new[] { 16 });
            var demandas = ctx.AulaAsignaturas.Select(a => (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList();
            var ranuras = AgRanuras.PorVentana("LUN-VIE", ctx.Bloques, "TARDE", 4);
            var plan = new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int> { 1 });
            var patrones = new Dictionary<int, string> { { 1, "PATRON_INVALIDO_XYZ" } };

            // Act
            var result = AgPlanificadorRanuras.ValidarFechado(plan, demandas, ranuras, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, patrones, System.Threading.CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I1_PATRON_INVALIDO");
        }

        private static CSharpFunctionalExtensions.Result<AgPlanOleadasResultado> Plan(AgContexto ctx)
        {
            return new AgPlanificadorOleadas().Construir(ctx.AulaAsignaturas.Select(a =>
                (a.IdAsignatura, a.HorasObjetivo, a.OrdenCurricular)).ToList(), "LUN-VIE",
                ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, 4, ctx.EsFestivo, ctx.Bloques, "TARDE");
        }

        private static AgContexto Crear(int[] horas)
        {
            var ctx = new AgContexto { Ventana = new VentanaSemestre
            { FechaInicio = new DateTime(2027, 2, 5), FechaFin = new DateTime(2027, 5, 30) } };
            ctx.Programas.Add(new Programa { Id = 1, Modalidad = "PRESENCIAL" });
            ctx.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, AulaRef = "CALI-TEST", Semestre = "III",
                Jornada = "TARDE", HorasDia = 4, DiasVentana = "LUN-VIE", Estado = AgEstadoAula.Disponible });
            for (int dia = 1; dia <= 7; dia++)
            {
                ctx.Bloques.Add(new AgBloqueHorario { Id = dia * 10, DiaSemanaNumero = dia, DuracionHoras = 4,
                    Franja = dia >= 6 ? "MANANA" : "TARDE", Activo = 1 });
                if (dia == 6)
                {
                    ctx.Bloques.Add(new AgBloqueHorario { Id = 61, DiaSemanaNumero = 6, DuracionHoras = 4, Franja = "TARDE", Activo = 1 });
                }
            }
            for (int i = 0; i < horas.Length; i++)
            {
                ctx.Asignaturas.Add(new Asignatura { Id = i + 1, IdPrograma = 1, Semestre = "III", Horas = horas[i] });
                ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = i + 1,
                    HorasObjetivo = horas[i], OrdenCurricular = i + 1, OrigenPatron = "AUTOMATICO", DiasPatron = "5" });
            }
            ctx.Festivos.Add((new DateTime(2027, 3, 22), new DateTime(2027, 3, 26)));
            ctx.ConstruirIndices();
            return ctx;
        }
    }
}
