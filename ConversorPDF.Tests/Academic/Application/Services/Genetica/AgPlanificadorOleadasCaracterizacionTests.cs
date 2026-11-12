using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Pruebas de caracterización de AgPlanificadorOleadas antes y después de la extracción
    /// del evaluador y selector de partición (Fase 1581).
    /// Aseguran que listas ordenadas completas e invariantes de negocio se preservan idénticamente.
    /// </summary>
    public class AgPlanificadorOleadasCaracterizacionTests
    {
        private readonly AgPlanificadorOleadas _sut = new AgPlanificadorOleadas();

        [Fact]
        public void Construir_SinPatrones_MarVie4Asignaturas_DebePreservarSalidaExacta()
        {
            // Arrange
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 11, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 64.0, 1),
                (102, 64.0, 2),
                (103, 64.0, 3),
                (104, 64.0, 4)
            };

            // Act
            var res = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            res.IsSuccess.Should().BeTrue();
            var val = res.Value;
            val.Infactibles.Should().BeEmpty();
            val.DetalleInfactibles.Should().BeEmpty();
            val.VentanaSugerida.Should().BeEmpty();
            val.PatronesRenormalizados.Should().Be(0);

            val.Planes.Should().HaveCount(4);
            val.Planes.Select(p => p.IdAsignatura).Should().Equal(101, 102, 103, 104);
            val.Planes.Select(p => p.Carril).Should().Equal(0, 1, 0, 1);
            val.Planes.Select(p => p.Oleada).Should().Equal(0, 0, 1, 1);

            val.Particion.Should().HaveCount(2);
            val.Particion[0].Should().Equal(2, 3);
            val.Particion[1].Should().Equal(4, 5);
        }

        [Fact]
        public void Construir_PatronesYaAlojables_VieSab_DebePreservarSalidaExacta()
        {
            // Arrange (VIE-SAB con 2 materias sembradas en día 5 = viernes, ya alojable en carril canónico 0)
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 6, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1),
                (2, 16.0, 2)
            };
            var patrones = new Dictionary<int, string> { { 1, "5" }, { 2, "5" } };

            // Act
            var res = _sut.Construir(asignaturas, "VIE-SAB", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.Should().HaveCount(2);
            res.Value.Planes.Select(p => p.Carril).Should().Equal(0, 0);
            res.Value.Planes.Select(p => p.Oleada).Should().Equal(0, 1);
        }

        [Fact]
        public void Construir_PatronFueraDeVentana_VieSabConJueves_DebeMarcarInfactibleYSugerirVentana()
        {
            // Arrange (patrón '4,5' en ventana VIE-SAB '5,6')
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 6, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 16.0, 1)
            };
            var patrones = new Dictionary<int, string> { { 1, "4,5" } };

            // Act
            var res = _sut.Construir(asignaturas, "VIE-SAB", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().Equal(1);
            res.Value.DetalleInfactibles.Should().ContainSingle()
                .Which.Should().Contain("Asig 1").And.Contain("R8_DiaFueraDePatron");
            res.Value.VentanaSugerida.Should().Be("JUE-SAB");
        }

        [Fact]
        public void Construir_CapacidadGlobalInsuficiente_DebeFallarConI5SinCapacidad()
        {
            // Arrange
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 8, 15);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 64.0, 1)
            };

            // Act
            var res = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SIN_CAPACIDAD");
        }

        [Fact]
        public void Construir_SinSemanasHabiles_DebeFallarConI5SinSemanasHabiles()
        {
            // Arrange (rango donde todos los días de la ventana son festivos o no hay días)
            var fInicio = new DateTime(2026, 3, 9);
            var fFin = new DateTime(2026, 3, 9);
            Func<DateTime, bool> esFestivo = d => true;
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 16.0, 1)
            };

            // Act
            var res = _sut.Construir(asignaturas, "LUN", fInicio, fFin, 4.0, esFestivo: esFestivo);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SIN_SEMANAS_HABILES");
        }

        [Fact]
        public void Construir_FestivosEnCarril_DebePreservarFechasYExtension()
        {
            // Arrange
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 5, 31);
            var asigs = new List<(int, double, int)> { (1, 16.0, 1) };
            Func<DateTime, bool> esFestivo = d => d.Date == new DateTime(2026, 3, 9);

            // Act
            var res = _sut.Construir(asigs, "LUN", fInicio, fFin, 4.0, esFestivo: esFestivo);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Planes.Should().HaveCount(1);
            var plan = res.Value.Planes[0];
            plan.FechaInicio.Should().Be(new DateTime(2026, 3, 2));
            plan.FechaFin.Should().Be(new DateTime(2026, 4, 5));
            plan.SemanasOleada.Should().Be(4);
        }

        [Fact]
        public void Construir_R8PatronCruzaCarril_MarVieConPatron2Y6_DebeReportarDetalleExacto()
        {
            // Arrange
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 6, 30);
            var asigs = new List<(int, double, int)> { (1, 16.0, 1) };
            var patrones = new Dictionary<int, string> { { 1, "2,6" } }; // 6 fuera de MAR-VIE

            // Act
            var res = _sut.Construir(asigs, "MAR-VIE", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().Equal(1);
            res.Value.DetalleInfactibles.Should().ContainSingle()
                .Which.Should().Contain("Asig 1")
                .And.Contain("patrón '2,6'")
                .And.Contain("R8_DiaFueraDePatron");
        }
    }
}
