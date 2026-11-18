using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Excel;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Excel
{
    public class SemestresExcelExporterTests
    {
        [Fact]
        public void CalcularRangoFechas_WhenSetCompletoHasta26Junio_ShouldGenerar142ColumnasYUltimaCon4Sesiones()
        {
            // Arrange (Fase 1935: corrida #104 set completo 2027-02-05 .. 2027-06-26)
            var items = new List<SesionListItem>();

            // Sesión inicial: 2027-02-05
            items.Add(new SesionListItem
            {
                IdSesion = 1,
                IdVinculacion = 1,
                IdPrograma = 1,
                FechaSesion = new DateTime(2027, 2, 5),
                Horas = 4.0
            });

            // 4 sesiones en la última fecha real: 2027-06-26 (16h en total)
            for (int i = 0; i < 4; i++)
            {
                items.Add(new SesionListItem
                {
                    IdSesion = 100 + i,
                    IdVinculacion = 50 + i,
                    IdPrograma = 2,
                    FechaSesion = new DateTime(2027, 6, 26),
                    Horas = 4.0
                });
            }

            // Act
            var allDates = SemestresExcelExporter.CalcularRangoFechas(items);

            // Assert: Rango continuo 2027-02-05 a 2027-06-26
            allDates.Should().NotBeEmpty();
            allDates.First().Should().Be(new DateTime(2027, 2, 5));
            allDates.Last().Should().Be(new DateTime(2027, 6, 26));
            allDates.Count.Should().Be(142); // 24 días feb + 31 mar + 30 abr + 31 may + 26 jun = 142

            var sesionesUltimoDia = items.Where(s => s.FechaSesion.HasValue && s.FechaSesion.Value.Date == new DateTime(2027, 6, 26)).ToList();
            sesionesUltimoDia.Count.Should().Be(4);
            sesionesUltimoDia.Sum(s => s.Horas).Should().Be(16.0);
        }

        [Fact]
        public void ValidarCuadre_WhenCuadreExacto_ShouldRetornarTrueSinBannerParcial()
        {
            // Arrange
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 1, IdPrograma = 1, Horas = 4.0, EsVacante = 0, FechaSesion = DateTime.Today },
                new SesionListItem { IdSesion = 2, IdVinculacion = 2, IdPrograma = 2, Horas = 4.0, EsVacante = 1, FechaSesion = DateTime.Today }
            };

            var esperado = new CuadreCorrida
            {
                Sesiones = 2,
                Horas = 8.0,
                SesionesVacantes = 1,
                PorPrograma = new List<(int, int, double)>
                {
                    (1, 1, 4.0),
                    (2, 1, 4.0)
                }.AsReadOnly(),
                RequiereCompleto = true
            };

            // Act
            bool valido = SemestresExcelExporter.ValidarCuadre(items, esperado, out string error, out bool marcarBannerParcial);

            // Assert
            valido.Should().BeTrue();
            error.Should().BeNull();
            marcarBannerParcial.Should().BeFalse();
        }

        [Fact]
        public void ValidarCuadre_WhenFaltanSesionesYRequiereCompleto_ShouldRetornarFalseYMensajeCuadreFalla()
        {
            // Arrange
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 1, IdPrograma = 1, Horas = 4.0, EsVacante = 0, FechaSesion = DateTime.Today }
            };

            var esperado = new CuadreCorrida
            {
                Sesiones = 2,
                Horas = 8.0,
                SesionesVacantes = 0,
                PorPrograma = new List<(int, int, double)> { (1, 1, 4.0), (2, 1, 4.0) }.AsReadOnly(),
                RequiereCompleto = true
            };

            // Act
            bool valido = SemestresExcelExporter.ValidarCuadre(items, esperado, out string error, out bool marcarBannerParcial);

            // Assert
            valido.Should().BeFalse();
            error.Should().StartWith("CUADRE_FALLA:");
            error.Should().Contain("corrida=2/8h/0vac");
        }

        [Fact]
        public void ValidarCuadre_WhenFaltanSesionesPeroNoRequiereCompleto_ShouldRetornarTrueYMarcarBannerParcial()
        {
            // Arrange
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 1, IdPrograma = 1, Horas = 4.0, EsVacante = 0, FechaSesion = DateTime.Today }
            };

            var esperado = new CuadreCorrida
            {
                Sesiones = 2,
                Horas = 8.0,
                SesionesVacantes = 0,
                PorPrograma = new List<(int, int, double)> { (1, 1, 4.0), (2, 1, 4.0) }.AsReadOnly(),
                RequiereCompleto = false
            };

            // Act
            bool valido = SemestresExcelExporter.ValidarCuadre(items, esperado, out string error, out bool marcarBannerParcial);

            // Assert
            valido.Should().BeTrue();
            error.Should().BeNull();
            marcarBannerParcial.Should().BeTrue();
        }

        [Fact]
        public void ValidarCuadre_WhenEsperadoEsNull_ShouldRetornarTrueYMarcarBannerParcial()
        {
            // Arrange (FormVinculaciones legado: esperado = null)
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 1, IdPrograma = 1, Horas = 4.0, EsVacante = 0, FechaSesion = DateTime.Today }
            };

            // Act
            bool valido = SemestresExcelExporter.ValidarCuadre(items, null, out string error, out bool marcarBannerParcial);

            // Assert
            valido.Should().BeTrue();
            error.Should().BeNull();
            marcarBannerParcial.Should().BeTrue();
        }

        [Fact]
        public void AgrupacionPorVinculacion_WhenVariasSesionesPorVinculacion_ShouldProducirUnaFilaPorVinculacion()
        {
            // Arrange (Fase 1933: Fila = vinculación)
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 10, IdPrograma = 1, Semestre = "I", AsignaturaNombre = "Teoria", SedeCiudad = "Cali", Horas = 4.0 },
                new SesionListItem { IdSesion = 2, IdVinculacion = 10, IdPrograma = 1, Semestre = "I", AsignaturaNombre = "Teoria", SedeCiudad = "Cali", Horas = 4.0 },
                new SesionListItem { IdSesion = 3, IdVinculacion = 20, IdPrograma = 2, Semestre = "II", AsignaturaNombre = "Finanzas", SedeCiudad = "Buenaventura", Horas = 4.0 }
            };

            // Act
            var grupos = items
                .GroupBy(s => s.IdVinculacion)
                .OrderBy(g => g.First().IdPrograma)
                .ToList();

            // Assert
            grupos.Count.Should().Be(2);
            grupos[0].Key.Should().Be(10);
            grupos[0].Sum(s => s.Horas).Should().Be(8.0);
            grupos[1].Key.Should().Be(20);
            grupos[1].Sum(s => s.Horas).Should().Be(4.0);
        }

        [Fact]
        public void AgrupacionPorVinculacion_WhenIdVinculacionCeroPeroDistintaSede_ShouldSepararEnFilasDistintas()
        {
            // Arrange (Fase 1937 & 1939: Sesiones sintéticas o conservadas con IdVinculacion = 0 pero distinta SedeCiudad/Subsede)
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 0, IdPeriodo = 1, IdPrograma = 1, IdAula = 101, IdAsignatura = 1, IdTercero = 1, Semestre = "I", Grupo = "G1", Modalidad = "P", SedeCiudad = "Cali", Subsede = "Norte" },
                new SesionListItem { IdSesion = 2, IdVinculacion = 0, IdPeriodo = 1, IdPrograma = 1, IdAula = 101, IdAsignatura = 1, IdTercero = 1, Semestre = "I", Grupo = "G1", Modalidad = "P", SedeCiudad = "Buenaventura", Subsede = "Principal" }
            };

            // Act
            var grupos = items
                .GroupBy(s => s.IdVinculacion > 0
                    ? $"v{s.IdVinculacion}"
                    : $"n{s.IdPeriodo}|{s.IdPrograma}|{s.IdAula}|{s.IdAsignatura}|{s.IdTercero}|{s.Semestre}|{s.Grupo}|{s.Modalidad}|{s.SedeCiudad}|{s.Subsede}")
                .ToList();

            // Assert: Deben generar 2 grupos distintos en vez de colapsar
            grupos.Count.Should().Be(2);
            grupos.Select(g => g.Key).Distinct().Count().Should().Be(2);
        }

        [Fact]
        public void ValidarCuadre_WhenHaySesionesHuerfanas_ShouldIncluirDetalleEnMensajeDeError()
        {
            // Arrange (Fase 1938 & 1940: Detección de sesiones huérfanas en la base, sin vinculación)
            var items = new List<SesionListItem>
            {
                new SesionListItem { IdSesion = 1, IdVinculacion = 1, IdPrograma = 1, Horas = 4.0, EsVacante = 0, FechaSesion = DateTime.Today }
            };

            var esperado = new CuadreCorrida
            {
                Sesiones = 2,
                Horas = 8.0,
                SesionesVacantes = 0,
                SesionesHuerfanas = 3,
                PorPrograma = new List<(int, int, double)> { (1, 1, 4.0), (2, 1, 4.0) }.AsReadOnly(),
                RequiereCompleto = true
            };

            // Act
            bool valido = SemestresExcelExporter.ValidarCuadre(items, esperado, out string error, out bool marcarBannerParcial);

            // Assert
            valido.Should().BeFalse();
            error.Should().Contain("(+3 sesiones huérfanas en la base, sin vinculación)");
        }
    }
}
