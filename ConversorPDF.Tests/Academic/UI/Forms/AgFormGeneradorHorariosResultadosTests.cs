using System;
using System.Collections.Generic;
using System.Drawing;
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
