using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgModoDatosTests
    {
        [Fact]
        public void Determinista_ShouldPointToProductionTablesAndViews()
        {
            // Arrange
            var modo = AgModoDatos.Determinista;

            // Assert
            modo.EsModoAg.Should().BeFalse();
            modo.TablaVinculaciones.Should().Be("vinculaciones");
            modo.TablaSesiones.Should().Be("sesiones");
            modo.TablaSalvedades.Should().Be("salvedades_conflictos");
            modo.VistaVinculacionesDetalle.Should().Be("vw_VinculacionesDetalle");
            modo.VistaSesionesDetalle.Should().Be("vw_SesionesDetalle");
        }

        [Fact]
        public void Ag_ShouldPointToStagingTablesAndViews()
        {
            // Arrange
            var modo = AgModoDatos.Ag;

            // Assert
            modo.EsModoAg.Should().BeTrue();
            modo.TablaVinculaciones.Should().Be("ag_vinculaciones");
            modo.TablaSesiones.Should().Be("ag_sesiones");
            modo.TablaSalvedades.Should().Be("ag_salvedades_conflictos");
            modo.VistaVinculacionesDetalle.Should().Be("vw_AgVinculacionesDetalle");
            modo.VistaSesionesDetalle.Should().Be("vw_AgSesionesDetalle");
        }
    }
}
