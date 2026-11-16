using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.UI.Forms;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace ConversorPDF.Tests.Academic.UI.Forms
{
    [TestClass]
    public class AgFormGeneradorHorariosResultadosTests
    {
        private sealed class STATestMethodAttribute : TestMethodAttribute
        {
            public override TestResult[] Execute(ITestMethod testMethod)
            {
                if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                {
                    return base.Execute(testMethod);
                }

                TestResult[] results = null;
                var thread = new Thread(() =>
                {
                    results = base.Execute(testMethod);
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                return results;
            }
        }

        private Mock<IAgScheduleRepository> _mockScheduleRepository;

        [TestInitialize]
        public void Setup()
        {
            _mockScheduleRepository = new Mock<IAgScheduleRepository>();

            _mockScheduleRepository
                .Setup(r => r.GetSesionesStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>
                {
                    new AgSesion
                    {
                        Id = 1,
                        IdGen = 10,
                        IdVinculacion = 1,
                        FechaSesion = new DateTime(2026, 4, 20),
                        Horas = 2.0,
                        EsVacante = 0
                    }
                }));

            _mockScheduleRepository
                .Setup(r => r.GetVinculacionesStagingAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<AgVinculacion>>(new List<AgVinculacion>()));

            _mockScheduleRepository
                .Setup(r => r.GetViolacionesByIndividuoAsync(It.IsAny<long>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(Result.Success<IReadOnlyCollection<AgViolacion>>(new List<AgViolacion>()));
        }

        [STATestMethod]
        public async Task CargarModuloResultadosAsync_WhenTieneViolaciones_ShouldSepararConteosRegistrosYPenalizaciones()
        {
            // Arrange
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                SetPrivateField(form, "_scheduleRepository", _mockScheduleRepository.Object);

                var individuo = new AgIndividuo
                {
                    Id = 1,
                    IdCorrida = 1,
                    Fitness = 0.95,
                    HardViolations = 3,
                    SoftViolations = 7.5,
                    Violaciones = new List<AgViolacion>
                    {
                        new AgViolacion { EsDura = true, Mensaje = "Dura 1" },
                        new AgViolacion { EsDura = true, Mensaje = "Dura 2" },
                        new AgViolacion { EsDura = false, Mensaje = "Blanda 1" }
                    },
                    Genes = new List<AgGen>()
                };

                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = 1,
                    Estado = AgEstadoCorrida.Completada,
                    MejorIndividuo = individuo
                };
                SetPrivateField(form, "_ultimoResultado", resultado);

                // Act
                bool cargado = await form.CargarModuloResultadosAsync(1);

                // Assert
                cargado.Should().BeTrue();
                var lblHardViolaciones = GetPrivateField<Label>(form, "lblHardViolaciones");
                lblHardViolaciones.Text.Should().Contain("Registros HARD=2, SOFT=1");
                lblHardViolaciones.Text.Should().Contain("infracciones HARD=3");
                lblHardViolaciones.Text.Should().Contain("penalización SOFT=7");
                lblHardViolaciones.ForeColor.Should().Be(Color.DarkRed);
            }
        }

        [STATestMethod]
        public async Task CargarModuloResultadosAsync_WhenCeroViolacionesCeroVacantes_ShouldHabilitarPromocionYColorVerde()
        {
            // Arrange
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                SetPrivateField(form, "_scheduleRepository", _mockScheduleRepository.Object);

                var individuo = new AgIndividuo
                {
                    Id = 1,
                    IdCorrida = 1,
                    Fitness = 1.0,
                    HardViolations = 0,
                    SoftViolations = 0.0,
                    Violaciones = new List<AgViolacion>(),
                    Genes = new List<AgGen>()
                };

                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = 1,
                    Estado = AgEstadoCorrida.Completada,
                    MejorIndividuo = individuo
                };
                SetPrivateField(form, "_ultimoResultado", resultado);

                // Act
                bool cargado = await form.CargarModuloResultadosAsync(1);

                // Assert
                cargado.Should().BeTrue();
                var lblHardViolaciones = GetPrivateField<Label>(form, "lblHardViolaciones");
                var btnPromoverProduccion = GetPrivateField<Button>(form, "btnPromoverProduccion");

                lblHardViolaciones.Text.Should().Contain("Registros HARD=0, SOFT=0");
                lblHardViolaciones.Text.Should().Contain("infracciones HARD=0");
                lblHardViolaciones.Text.Should().Contain("penalización SOFT=0");
                lblHardViolaciones.ForeColor.Should().Be(Color.DarkGreen);
                btnPromoverProduccion.Enabled.Should().BeTrue();
            }
        }

        [STATestMethod]
        public async Task CargarModuloResultadosAsync_WhenTieneHardViolationsMayorACero_ShouldDeshabilitarPromocion()
        {
            // Arrange
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                SetPrivateField(form, "_scheduleRepository", _mockScheduleRepository.Object);

                var individuo = new AgIndividuo
                {
                    Id = 1,
                    IdCorrida = 1,
                    Fitness = 0.8,
                    HardViolations = 1,
                    SoftViolations = 0.0,
                    Violaciones = new List<AgViolacion>
                    {
                        new AgViolacion { EsDura = true, Mensaje = "Conflicto docente" }
                    },
                    Genes = new List<AgGen>()
                };

                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = 1,
                    Estado = AgEstadoCorrida.Completada,
                    MejorIndividuo = individuo
                };
                SetPrivateField(form, "_ultimoResultado", resultado);

                // Act
                bool cargado = await form.CargarModuloResultadosAsync(1);

                // Assert
                cargado.Should().BeTrue();
                var btnPromoverProduccion = GetPrivateField<Button>(form, "btnPromoverProduccion");
                btnPromoverProduccion.Enabled.Should().BeFalse();
            }
        }

        [STATestMethod]
        public async Task CargarModuloResultadosAsync_WhenTieneVacantesMayorACero_ShouldDeshabilitarPromocion()
        {
            // Arrange
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                SetPrivateField(form, "_scheduleRepository", _mockScheduleRepository.Object);

                var individuo = new AgIndividuo
                {
                    Id = 1,
                    IdCorrida = 1,
                    Fitness = 0.85,
                    HardViolations = 0,
                    SoftViolations = 0.0,
                    Violaciones = new List<AgViolacion>(),
                    Genes = new List<AgGen>
                    {
                        new AgGen
                        {
                            Id = 1,
                            EsVacante = 1,
                            IdTercero = 0
                        }
                    }
                };

                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = 1,
                    Estado = AgEstadoCorrida.Completada,
                    MejorIndividuo = individuo
                };
                SetPrivateField(form, "_ultimoResultado", resultado);

                // Act
                bool cargado = await form.CargarModuloResultadosAsync(1);

                // Assert
                cargado.Should().BeTrue();
                individuo.Vacantes.Should().Be(1);
                var btnPromoverProduccion = GetPrivateField<Button>(form, "btnPromoverProduccion");
                btnPromoverProduccion.Enabled.Should().BeFalse();
            }
        }

        [STATestMethod]
        public async Task CargarModuloResultadosAsync_WhenTieneHardYVacantesMayorACero_ShouldDeshabilitarPromocion()
        {
            // Arrange
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                SetPrivateField(form, "_scheduleRepository", _mockScheduleRepository.Object);

                var individuo = new AgIndividuo
                {
                    Id = 1,
                    IdCorrida = 1,
                    Fitness = 0.7,
                    HardViolations = 2,
                    SoftViolations = 1.0,
                    Violaciones = new List<AgViolacion>
                    {
                        new AgViolacion { EsDura = true, Mensaje = "Hard 1" }
                    },
                    Genes = new List<AgGen>
                    {
                        new AgGen
                        {
                            Id = 1,
                            EsVacante = 1,
                            IdTercero = 0
                        }
                    }
                };

                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = 1,
                    Estado = AgEstadoCorrida.Completada,
                    MejorIndividuo = individuo
                };
                SetPrivateField(form, "_ultimoResultado", resultado);

                // Act
                bool cargado = await form.CargarModuloResultadosAsync(1);

                // Assert
                cargado.Should().BeTrue();
                var btnPromoverProduccion = GetPrivateField<Button>(form, "btnPromoverProduccion");
                btnPromoverProduccion.Enabled.Should().BeFalse();
            }
        }

        [STATestMethod]
        public async Task CerrarCorridaFallidaAsync_ErrorYCancelacion_PersistenSoloCabeceraConTokenNoCancelado()
        {
            foreach (bool cancelada in new[] { false, true })
            {
                using (var form = new AgFormGeneradorHorarios(() => null))
                using (var cts = new CancellationTokenSource())
                {
                    cts.Cancel();
                    SetPrivateField(form, "_ctsEjecucion", cts);
                    var repo = new Mock<IAgScheduleRepository>(MockBehavior.Strict);
                    var corrida = new AgCorrida { Id = 95, ParametrosJson = "snapshot", IdIndividuoOptimo = 12 };
                    var estado = cancelada ? AgEstadoCorrida.Cancelada : AgEstadoCorrida.Error;
                    repo.Setup(r => r.UpdateCorridaAsync(corrida, It.IsAny<CancellationToken>()))
                        .Callback<AgCorrida, CancellationToken>((c, ct) =>
                        {
                            c.Estado.Should().Be(estado);
                            c.FechaFinCorrida.Should().NotBeNullOrWhiteSpace();
                            c.ParametrosJson.Should().Be("snapshot");
                            c.IdIndividuoOptimo.Should().Be(12);
                            ct.CanBeCanceled.Should().BeTrue();
                            ct.Should().NotBe(cts.Token);
                            cts.IsCancellationRequested.Should().BeTrue();
                            ct.IsCancellationRequested.Should().BeFalse();
                        })
                        .ReturnsAsync(Result.Success());
                    SetPrivateField(form, "_scheduleRepository", repo.Object);

                    await (Task)InvokePrivate(form, "CerrarCorridaFallidaAsync", corrida, cancelada);

                    repo.Verify(r => r.UpdateCorridaAsync(corrida,
                        It.Is<CancellationToken>(ct => ct.CanBeCanceled && ct != cts.Token)), Times.Once);
                    repo.VerifyNoOtherCalls(); // No purga global, staging ni producción.
                }
            }
        }

        [STATestMethod]
        public async Task CerrarCorridaFallidaAsync_SinCabecera_NoEscribeYFalloDeRepositorioNoEscapa()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var repo = new Mock<IAgScheduleRepository>(MockBehavior.Strict);
                SetPrivateField(form, "_scheduleRepository", repo.Object);
                await (Task)InvokePrivate(form, "CerrarCorridaFallidaAsync", null, false);
                await (Task)InvokePrivate(form, "CerrarCorridaFallidaAsync", new AgCorrida(), false);
                repo.VerifyNoOtherCalls();

                var corrida = new AgCorrida { Id = 95 };
                repo.Setup(r => r.UpdateCorridaAsync(corrida, It.IsAny<CancellationToken>()))
                    .ReturnsAsync(Result.Failure("Fallo de persistencia controlado"));
                await (Task)InvokePrivate(form, "CerrarCorridaFallidaAsync", corrida, false);
                corrida.Estado.Should().Be(AgEstadoCorrida.Error);
                repo.Verify(r => r.UpdateCorridaAsync(corrida,
                    It.Is<CancellationToken>(ct => ct.CanBeCanceled)), Times.Once);
                repo.VerifyNoOtherCalls();
            }
        }

        [STATestMethod]
        public void ProgresoOptimizacion_ExigeIdentidadVigenteYSeRechazaTrasCierreVisual()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                var intento = Guid.NewGuid();
                SetPrivateField(form, "_intentoOptimizacion", intento);
                SetPrivateField(form, "_aceptaProgresoOptimizacion", true);
                ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento)).Should().BeTrue();
                ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", Guid.NewGuid())).Should().BeFalse();

                foreach (bool cancelada in new[] { false, true })
                {
                    // Cada cierre genérico parte de una fase activa, no del terminal anterior.
                    EstablecerEstado(form, "Optimizando");
                    GetPrivateField<Label>(form, "lblStatusAG").Text = "Estado: Optimizando...";
                    GetPrivateField<Button>(form, "btnPromoverProduccion").Enabled = true;
                    InvokePrivate(form, "CerrarEstadoVisualOptimizacion", cancelada);
                    GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(cancelada
                        ? "Estado: Cancelado por el usuario." : "Estado: Error en ejecución.");
                    GetPrivateField<Label>(form, "lblStatusAG").ForeColor.Should().Be(cancelada
                        ? Color.DarkOrange : Color.DarkRed);
                    GetPrivateField<Button>(form, "btnPromoverProduccion").Enabled.Should().BeFalse();
                    GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().Be(cancelada ? "CANCELACION" : "FALLO_EJECUCION");
                    GetPrivateField<string>(form, "_faseFinOptimizacion").Should().Be("Optimizando");
                    ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento)).Should().BeFalse();
                }

                // Un intento nuevo admite sus callbacks, nunca los del anterior.
                var siguiente = Guid.NewGuid();
                SetPrivateField(form, "_intentoOptimizacion", siguiente);
                SetPrivateField(form, "_aceptaProgresoOptimizacion", true);
                ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento)).Should().BeFalse();
                ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", siguiente)).Should().BeTrue();
                form.Dispose();
                ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", siguiente)).Should().BeFalse();
            }
        }

        [STATestMethod]
        public void CierreGenerico_NoSobrescribeMotivoSnapshotNiTerminalCompletadoOCancelado()
        {
            foreach (string estado in new[] { "Error", "Completado", "Cancelado" })
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                var intento = Guid.NewGuid();
                var repo = new Mock<IAgScheduleRepository>(MockBehavior.Strict);
                SetPrivateField(form, "_scheduleRepository", repo.Object);
                SetPrivateField(form, "_intentoOptimizacion", intento);
                EstablecerEstado(form, "Validando");
                string motivo = estado == "Error" ? "SNAPSHOT_SERIALIZACION_ERROR"
                    : estado == "Completado" ? "HORARIO_FACTIBLE" : "CANCELACION";
                if (estado == "Error") InvokePrivate(form, "MarcarErrorPreparacion", motivo);
                else EstablecerEstado(form, estado, motivo);
                var texto = GetPrivateField<Label>(form, "lblStatusAG").Text;
                var color = GetPrivateField<Label>(form, "lblStatusAG").ForeColor;

                // Misma llamada que usan ambos catch; ni error genérico ni cancelación tardía degradan el terminal.
                foreach (bool cancelada in new[] { false, true })
                {
                    InvokePrivate(form, "CerrarEstadoVisualOptimizacion", cancelada);
                    GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().Be(motivo);
                    GetPrivateField<string>(form, "_faseFinOptimizacion").Should().Be("Validando");
                    GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(texto);
                    GetPrivateField<Label>(form, "lblStatusAG").ForeColor.Should().Be(color);
                    GetPrivateField<bool>(form, "_terminalVisualOptimizacion").Should().BeTrue();
                    ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento)).Should().BeFalse();
                }
                repo.VerifyNoOtherCalls(); // No reabre ni modifica el terminal persistido.
            }
        }

        [STATestMethod]
        public async Task IniciarOptimizacionAsync_IntentoActivo_RechazaReentradaSinMutarControles()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var intento = Guid.NewGuid();
                SetPrivateField(form, "_intentoOptimizacion", intento);
                var estado = GetPrivateField<Label>(form, "lblStatusAG").Text;
                var repo = new Mock<IAgScheduleRepository>(MockBehavior.Strict);
                SetPrivateField(form, "_scheduleRepository", repo.Object);

                await (Task)InvokePrivate(form, "IniciarOptimizacionAsync");

                GetPrivateField<Guid?>(form, "_intentoOptimizacion").Should().Be(intento);
                GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(estado);
                repo.VerifyNoOtherCalls();
            }
        }

        [TestMethod]
        public void IniciarOptimizacionAsync_FalloDePresentacion_NoReabreTerminalPersistido_SourceContract()
        {
            // Fixture estructural: el recorrido UI completo requiere preparación y bucle de mensajes.
            var directorio = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            string fuente = null;
            while (directorio != null)
            {
                var archivo = Path.Combine(directorio.FullName, "ConversorPDF.XLL", "Academic", "UI", "Forms", "AgFormGeneradorHorarios.cs");
                if (File.Exists(archivo))
                {
                    fuente = File.ReadAllText(archivo);
                    break;
                }
                directorio = directorio.Parent;
            }
            fuente.Should().NotBeNull("el fixture requiere el checkout accesible desde la salida de pruebas");
            fuente = fuente.Replace("\r\n", "\n");
            int inicio = fuente.IndexOf("private async Task IniciarOptimizacionAsync()", StringComparison.Ordinal);
            // El método puede presentar resultados en más de una ruta (p. ej. Sólo conservación);
            // el invariante se verifica sobre el método completo y no sobre una llamada concreta.
            int fin = fuente.IndexOf("private bool PuedeRecibirProgresoOptimizacion(", inicio, StringComparison.Ordinal);
            inicio.Should().BeGreaterOrEqualTo(0, "el fixture debe localizar IniciarOptimizacionAsync");
            fin.Should().BeGreaterThan(inicio, "el fixture debe delimitar el método");
            var cuerpo = fuente.Substring(inicio, fin - inicio);
            cuerpo.Should().Contain("if (cierreCorrida.IsFailure) throw new InvalidOperationException(cierreCorrida.Error);");
            cuerpo.Should().Contain("corridaFinalizada = true;");
            cuerpo.Should().Contain("catch (OperationCanceledException)");
            cuerpo.Should().Contain("catch (Exception ex)");
            // Sólo se prohíbe el reinicio del terminal; la declaración inicial (bool corridaFinalizada = false;) es legítima.
            cuerpo.Should().NotMatchRegex(@"\n\s*corridaFinalizada = false;");
            cuerpo.Should().Contain("if (!corridaFinalizada)");
            cuerpo.Should().Contain("await CerrarCorridaFallidaAsync(corrida, cancelada);");
        }

        [STATestMethod]
        public void Preparacion_DistingueFasesYPreguntas_SoloOptimizandoAdmiteProgreso()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var handle = form.Handle;
                var intento = Guid.NewGuid();
                SetPrivateField(form, "_intentoOptimizacion", intento);
                var estados = new[] { "Preparando", "Validando", "RequiereDecision", "Validando", "Optimizando", "Materializando" };
                var textos = new[] { "Estado: Preparando...", "Estado: Validando...",
                    "Estado: Decisión pendiente de confirmación.", "Estado: Validando...",
                    "Estado: Optimizando...", "Estado: Materializando..." };
                for (int i = 0; i < estados.Length; i++)
                {
                    EstablecerEstado(form, estados[i]);
                    GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(textos[i]);
                    GetPrivateField<Label>(form, "lblStatusAG").ForeColor.Should().Be(
                        estados[i] == "RequiereDecision" ? Color.DarkOrange : Color.DarkBlue);
                    GetPrivateField<bool>(form, "_terminalVisualOptimizacion").Should().BeFalse();
                    ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento))
                        .Should().Be(estados[i] == "Optimizando");
                    GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().BeNull();
                }
            }
        }

        [STATestMethod]
        public void Terminales_ConservanFaseYMotivo_DeshabilitanPromocionYProgreso()
        {
            var estados = new[] { "Completado", "Bloqueado", "Cancelado", "Error" };
            var textos = new[] { "Estado: Completado.", "Estado: Bloqueado. Revise el diagnóstico.",
                "Estado: Cancelado por el usuario.", "Estado: Error en ejecución." };
            var colores = new[] { Color.DarkGreen, Color.DarkOrange, Color.DarkOrange, Color.DarkRed };
            for (int i = 0; i < estados.Length; i++)
            {
                using (var form = new AgFormGeneradorHorarios(() => null))
                {
                    var handle = form.Handle;
                    var intento = Guid.NewGuid();
                    var repo = new Mock<IAgScheduleRepository>(MockBehavior.Strict);
                    SetPrivateField(form, "_scheduleRepository", repo.Object);
                    SetPrivateField(form, "_intentoOptimizacion", intento);
                    EstablecerEstado(form, "Validando");
                    GetPrivateField<Button>(form, "btnPromoverProduccion").Enabled = true;

                    EstablecerEstado(form, estados[i], "CODIGO_CONTROLADO");

                    GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(textos[i]);
                    GetPrivateField<Label>(form, "lblStatusAG").ForeColor.Should().Be(colores[i]);
                    GetPrivateField<bool>(form, "_terminalVisualOptimizacion").Should().BeTrue();
                    GetPrivateField<string>(form, "_faseFinOptimizacion").Should().Be("Validando");
                    GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().Be("CODIGO_CONTROLADO");
                    GetPrivateField<Button>(form, "btnPromoverProduccion").Enabled.Should().BeFalse();
                    ((bool)InvokePrivate(form, "PuedeRecibirProgresoOptimizacion", intento)).Should().BeFalse();
                    repo.VerifyNoOtherCalls(); // El estado visual no escribe ni autoriza cobertura.
                }
            }
        }

        [STATestMethod]
        public void HelpersPreparacion_SinIntento_NoAlteranDiagnostico_ConIntentoConservanPrimerMotivo()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            {
                var estado = GetPrivateField<Label>(form, "lblStatusAG").Text;
                InvokePrivate(form, "MarcarErrorPreparacion", "LECTURA");
                InvokePrivate(form, "MarcarDecisionPreparacion");
                InvokePrivate(form, "ReanudarValidacionPreparacion");
                GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be(estado);

                SetPrivateField(form, "_intentoOptimizacion", Guid.NewGuid());
                InvokePrivate(form, "MarcarDecisionPreparacion");
                GetPrivateField<bool>(form, "_terminalVisualOptimizacion").Should().BeFalse();
                InvokePrivate(form, "ReanudarValidacionPreparacion");
                InvokePrivate(form, "MarcarErrorPreparacion", "P06_FiltroTercerosVacio");
                InvokePrivate(form, "MarcarErrorPreparacion", "P08_ValidarContexto");
                GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().Be("P06_FiltroTercerosVacio");
                GetPrivateField<string>(form, "_faseFinOptimizacion").Should().Be("Validando");
                GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be("Estado: Error en ejecución.");
            }
        }

        [STATestMethod]
        public void ErrorPreparacion_ConTokenCancelado_NoLoConvierteEnDeficitNiErrorGenerico()
        {
            using (var form = new AgFormGeneradorHorarios(() => null))
            using (var cts = new CancellationTokenSource())
            {
                SetPrivateField(form, "_intentoOptimizacion", Guid.NewGuid());
                SetPrivateField(form, "_ctsEjecucion", cts);
                EstablecerEstado(form, "Validando");
                cts.Cancel();

                InvokePrivate(form, "MarcarErrorPreparacion", "LECTURA");

                GetPrivateField<Label>(form, "lblStatusAG").Text.Should().Be("Estado: Cancelado por el usuario.");
                GetPrivateField<string>(form, "_motivoFinOptimizacion").Should().Be("CANCELACION");
                GetPrivateField<bool>(form, "_terminalVisualOptimizacion").Should().BeTrue();
            }
        }

        [STATestMethod]
        public void Constructor_LectorCanonicoInyectado_NoLeeNiEscribeDuranteComposicion()
        {
            var lector = new Mock<IAgPlanAprobadoRepository>(MockBehavior.Strict);
            using (var form = new AgFormGeneradorHorarios(
                () => { throw new InvalidOperationException("La composición no debe abrir conexiones"); }, lector.Object))
            {
                GetPrivateField<IAgPlanAprobadoRepository>(form, "_planAprobadoRepository").Should().BeSameAs(lector.Object);
                lector.VerifyNoOtherCalls();
                GetPrivateField<AgContexto>(form, "_contextoActual").Should().BeNull();
            }
        }

        private static void EstablecerEstado(AgFormGeneradorHorarios form, string estado, string motivo = null)
        {
            var tipo = typeof(AgFormGeneradorHorarios).GetNestedType("EstadoVisualOptimizacion", BindingFlags.NonPublic);
            InvokePrivate(form, "EstablecerEstadoOptimizacion", Enum.Parse(tipo, estado), motivo, null);
        }

        private static object InvokePrivate(object target, string method, params object[] args)
        {
            return target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(target, args);
        }

        private static void SetPrivateField(object target, string fieldName, object value)
        {
            var field = target.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (field == null)
            {
                throw new ArgumentException("Campo '" + fieldName + "' no encontrado en " + target.GetType().Name);
            }
            field.SetValue(target, value);
        }

        private static T GetPrivateField<T>(object target, string fieldName)
        {
            var field = target.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (field == null)
            {
                throw new ArgumentException("Campo '" + fieldName + "' no encontrado en " + target.GetType().Name);
            }
            return (T)field.GetValue(target);
        }
    }
}
