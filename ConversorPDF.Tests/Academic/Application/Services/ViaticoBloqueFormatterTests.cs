using System;
using System.Collections.Generic;
using System.Globalization;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class ViaticoBloqueFormatterTests
    {
        public ViaticoBloqueFormatterTests()
        {
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("es-CO");
            CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo("es-CO");
        }

        [Fact]
        public void ObtenerTextoCierre_WhenBloqueNull_ShouldReturnSF()
        {
            // Act
            string cierre = ViaticoBloqueFormatter.ObtenerTextoCierre(null);

            // Assert
            cierre.Should().Be("S/F");
        }

        [Fact]
        public void ObtenerTextoCierre_WhenSinSesionesEnFechaMaxima_ShouldReturnSF()
        {
            // Arrange
            var bloque = new ViaticoBloque();

            // Act
            string cierre = ViaticoBloqueFormatter.ObtenerTextoCierre(bloque);

            // Assert
            cierre.Should().Be("S/F");
        }

        [Fact]
        public void ObtenerTextoCierre_WhenUnSoloDia_ShouldReturnDiaYHorario()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                DiaSemana = "Viernes",
                HorarioSesion = "18:00 A 22:00",
                Horas = 4
            });

            // Act
            string cierre = ViaticoBloqueFormatter.ObtenerTextoCierre(bloque);

            // Assert
            cierre.Should().Be("VIERNES 18:00 A 22:00");
        }

        [Fact]
        public void ObtenerTextoCierre_WhenDosSesionesMismoDia_ShouldReturnSesionConHoraInicioMasTardia()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            // Sesión de la mañana
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 17),
                DiaSemana = "Sábado",
                HorarioSesion = "08:00 A 12:00",
                Horas = 4
            });
            // Sesión de la tarde (cierre)
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17),
                DiaSemana = "Sábado",
                HorarioSesion = "14:00 A 18:00",
                Horas = 4
            });

            // Act
            string cierre = ViaticoBloqueFormatter.ObtenerTextoCierre(bloque);

            // Assert
            cierre.Should().Be("SÁBADO 14:00 A 18:00");
        }

        [Fact]
        public void ObtenerTextoCierre_WhenHorario12Horas_ShouldParseInicioCorrectly()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                DiaSemana = "Viernes",
                HorarioSesion = "6:00 PM A 10:00 PM",
                Horas = 4
            });

            // Act
            string cierre = ViaticoBloqueFormatter.ObtenerTextoCierre(bloque);

            // Assert
            cierre.Should().Be("VIERNES 6:00 PM A 10:00 PM");
        }

        [Fact]
        public void FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludePernoctaUltimoDiaSuffix()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                DiaSemana = "Viernes",
                HorarioSesion = "18:00 A 22:00",
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17),
                DiaSemana = "Sábado",
                HorarioSesion = "08:00 A 12:00 M",
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });

            // Act
            string resultado = ViaticoBloqueFormatter.FormatearNodoNivel2(
                bloque,
                "16/10/2026",
                "17/10/2026",
                "Terrestre",
                "POPAYÁN - EL ÁGUILA - POPAYÁN",
                120000m);

            // Assert
            resultado.Should().Be("📅 16/10/2026 - 17/10/2026 | 1,5 d | Terrestre | POPAYÁN - EL ÁGUILA - POPAYÁN | $120.000 | SÁBADO 08:00 A 12:00 M");
            resultado.Should().NotContain("Pernocta Ultimo Dia");
        }

        [Fact]
        public void FormatearNodoNivel2_WhenConPernocta_ShouldIncludePernoctaUltimoDiaSuffix()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                DiaSemana = "Viernes",
                HorarioSesion = "18:00 A 22:00",
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17),
                DiaSemana = "Sábado",
                HorarioSesion = "08:00 A 12:00 M",
                Horas = 4,
                IsPernoctaUltimoDia = 1
            });

            // Act
            string fechaFin = bloque.FechaFinComision.Value.ToString("dd/MM/yyyy");
            string resultado = ViaticoBloqueFormatter.FormatearNodoNivel2(
                bloque,
                "16/10/2026",
                fechaFin,
                "Terrestre",
                "POPAYÁN - EL ÁGUILA - POPAYÁN",
                120000m);

            // Assert
            resultado.Should().Be("📅 16/10/2026 - 18/10/2026 | 2,5 d | Terrestre | POPAYÁN - EL ÁGUILA - POPAYÁN | $120.000 | SÁBADO 08:00 A 12:00 M | Pernocta Ultimo Dia");
        }

        [Fact]
        public void FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludeCierrePrefixNorCodigoRuta()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                DiaSemana = "Viernes",
                HorarioSesion = "18:00 A 22:00",
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });

            // Act
            string resultado = ViaticoBloqueFormatter.FormatearNodoNivel2(
                bloque,
                "16/10/2026",
                "16/10/2026",
                "Terrestre",
                "CALI - BUENAVENTURA - CALI",
                80000m);

            // Assert
            resultado.Should().NotContain("Cierre:");
            resultado.Should().NotContain("SIN-COD");
            resultado.Should().Contain("| 0,5 d |");
        }

        [Fact]
        public void FormatearNodoNivel2_WhenBloqueNull_ShouldReturnEmpty()
        {
            // Act
            string resultado = ViaticoBloqueFormatter.FormatearNodoNivel2(
                null,
                "16/10/2026",
                "17/10/2026",
                "Terrestre",
                "POPAYÁN - EL ÁGUILA - POPAYÁN",
                120000m);

            // Assert
            resultado.Should().BeEmpty();
        }

        [Fact]
        public void ObtenerRangoComision_WhenConPernoctaUltimoDia_ShouldReturnFinExtendido()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17),
                Horas = 4,
                IsPernoctaUltimoDia = 1
            });

            // Act
            var rango = ViaticoBloqueFormatter.ObtenerRangoComision(bloque);

            // Assert
            rango.Should().NotBeNull();
            rango.Value.Inicio.Should().Be(new DateTime(2026, 10, 16));
            rango.Value.Fin.Should().Be(new DateTime(2026, 10, 18));
        }

        [Fact]
        public void ObtenerRangoComision_WhenSinPernocta_ShouldReturnFechaMaxima()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16),
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });
            bloque.Agregar(new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17),
                Horas = 4,
                IsPernoctaUltimoDia = 0
            });

            // Act
            var rango = ViaticoBloqueFormatter.ObtenerRangoComision(bloque);

            // Assert
            rango.Should().NotBeNull();
            rango.Value.Inicio.Should().Be(new DateTime(2026, 10, 16));
            rango.Value.Fin.Should().Be(new DateTime(2026, 10, 17));
        }

        [Fact]
        public void ObtenerRangoComision_WhenBloqueVacio_ShouldReturnNull()
        {
            // Arrange
            var bloque = new ViaticoBloque();

            // Act
            var rango = ViaticoBloqueFormatter.ObtenerRangoComision(bloque);

            // Assert
            rango.Should().BeNull();
        }
    }
}
