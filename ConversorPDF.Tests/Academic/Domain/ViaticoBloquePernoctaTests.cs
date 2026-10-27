using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class ViaticoBloquePernoctaTests
    {
        [Fact]
        public void DuracionTotal_WhenSinFlag_ShouldReturn1Punto5AndFinIgualFechaMaxima()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            var sesion1 = new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16), // Viernes
                Horas = 4,
                IsPernoctaUltimoDia = 0
            };
            var sesion2 = new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17), // Sábado
                Horas = 4,
                IsPernoctaUltimoDia = 0
            };

            bloque.Agregar(sesion1);
            bloque.Agregar(sesion2);

            // Assert
            bloque.TienePernoctaUltimoDia.Should().BeFalse();
            bloque.FechaMaxima.Should().Be(new DateTime(2026, 10, 17));
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 17));
            bloque.DuracionTotal.Should().Be(1.5);
        }

        [Fact]
        public void DuracionTotal_WhenConFlagPernocta_ShouldReturn2Punto5AndFinSiguienteDia()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            var sesion1 = new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 16), // Viernes
                Horas = 4,
                IsPernoctaUltimoDia = 0
            };
            var sesion2 = new SesionListItem
            {
                IdSesion = 2,
                FechaSesion = new DateTime(2026, 10, 17), // Sábado
                Horas = 4,
                IsPernoctaUltimoDia = 1 // Flag activo en vinculación del bloque
            };

            bloque.Agregar(sesion1);
            bloque.Agregar(sesion2);

            // Assert
            bloque.TienePernoctaUltimoDia.Should().BeTrue();
            bloque.FechaMaxima.Should().Be(new DateTime(2026, 10, 17)); // FechaMaxima intacta
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 18)); // +1 día
            bloque.DuracionTotal.Should().Be(2.5); // (18 - 16).Days + 0.5 = 2.5
        }

        [Fact]
        public void DuracionTotal_WhenBloqueDe1DiaConPernocta_ShouldReturn1Punto5()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            var sesion = new SesionListItem
            {
                IdSesion = 1,
                FechaSesion = new DateTime(2026, 10, 17),
                Horas = 4,
                IsPernoctaUltimoDia = 1
            };

            bloque.Agregar(sesion);

            // Assert
            bloque.TienePernoctaUltimoDia.Should().BeTrue();
            bloque.FechaMaxima.Should().Be(new DateTime(2026, 10, 17));
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 18));
            bloque.DuracionTotal.Should().Be(1.5); // (18 - 17).Days + 0.5 = 1.5
        }

        [Fact]
        public void AgruparConsecutivos_WhenDosBloquesRecurrentesConPernocta_ShouldBothHavePlusOne()
        {
            // Arrange: 2 semanas distintas (fin de semana 1 y fin de semana 2)
            var sesiones = new List<SesionListItem>
            {
                // Semana 1: 16 y 17 de Octubre
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 10, 16), Horas = 4, IsPernoctaUltimoDia = 1 },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 10, 17), Horas = 4, IsPernoctaUltimoDia = 1 },
                // Semana 2: 23 y 24 de Octubre
                new SesionListItem { IdSesion = 3, FechaSesion = new DateTime(2026, 10, 23), Horas = 4, IsPernoctaUltimoDia = 1 },
                new SesionListItem { IdSesion = 4, FechaSesion = new DateTime(2026, 10, 24), Horas = 4, IsPernoctaUltimoDia = 1 }
            };

            // Act
            var bloques = ViaticoBloque.AgruparConsecutivos(sesiones);

            // Assert
            bloques.Should().HaveCount(2);

            // Bloque 1
            bloques[0].TienePernoctaUltimoDia.Should().BeTrue();
            bloques[0].FechaMaxima.Should().Be(new DateTime(2026, 10, 17));
            bloques[0].FechaFinComision.Should().Be(new DateTime(2026, 10, 18));
            bloques[0].DuracionTotal.Should().Be(2.5);

            // Bloque 2
            bloques[1].TienePernoctaUltimoDia.Should().BeTrue();
            bloques[1].FechaMaxima.Should().Be(new DateTime(2026, 10, 24));
            bloques[1].FechaFinComision.Should().Be(new DateTime(2026, 10, 25));
            bloques[1].DuracionTotal.Should().Be(2.5);
        }

        [Fact]
        public void TienePernoctaUltimoDia_WhenBloqueAjenoSinFlag_ShouldBeFalseAndDuracionSinCambio()
        {
            // Arrange: Bloque A con pernocta, Bloque B sin pernocta
            var sesiones = new List<SesionListItem>
            {
                // Bloque con flag
                new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 10, 16), Horas = 4, IsPernoctaUltimoDia = 1 },
                new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 10, 17), Horas = 4, IsPernoctaUltimoDia = 1 },
                // Bloque sin flag
                new SesionListItem { IdSesion = 3, FechaSesion = new DateTime(2026, 10, 23), Horas = 4, IsPernoctaUltimoDia = 0 },
                new SesionListItem { IdSesion = 4, FechaSesion = new DateTime(2026, 10, 24), Horas = 4, IsPernoctaUltimoDia = 0 }
            };

            // Act
            var bloques = ViaticoBloque.AgruparConsecutivos(sesiones);

            // Assert
            bloques.Should().HaveCount(2);
            bloques[0].DuracionTotal.Should().Be(2.5);
            bloques[1].TienePernoctaUltimoDia.Should().BeFalse();
            bloques[1].DuracionTotal.Should().Be(1.5);
            bloques[1].FechaFinComision.Should().Be(new DateTime(2026, 10, 24));
        }

        [Fact]
        public void FechaMaxima_WhenConFlagPernocta_ShouldRemainIntactForConsecutividad()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 10, 16), Horas = 4, IsPernoctaUltimoDia = 1 });
            bloque.Agregar(new SesionListItem { IdSesion = 2, FechaSesion = new DateTime(2026, 10, 17), Horas = 4, IsPernoctaUltimoDia = 1 });

            // Assert: FechaMaxima no fue alterada por el cálculo de FechaFinComision
            bloque.FechaMaxima.Should().Be(new DateTime(2026, 10, 17));
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 18));
        }

        [Fact]
        public void DuracionTotal_WhenSesionPernoctaMutatesInPlace_ShouldRecalculateWithoutCallingAgregar()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            var sesion = new SesionListItem { IdSesion = 1, FechaSesion = new DateTime(2026, 10, 16), Horas = 4, IsPernoctaUltimoDia = 0 };
            bloque.Agregar(sesion);

            // Act 1: Initial state without pernocta
            double inicial = bloque.DuracionTotal;
            inicial.Should().Be(0.5);

            // Act 2: Mutate session flag in place
            sesion.IsPernoctaUltimoDia = 1;
            double conPernocta = bloque.DuracionTotal;

            // Assert
            conPernocta.Should().Be(1.5);
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 17));
        }

        [Fact]
        public void IdsVinculacion_WhenSesionesConMultiplesVinculaciones_ShouldReturnDistinctIds()
        {
            // Arrange
            var bloque = new ViaticoBloque();
            bloque.Agregar(new SesionListItem { IdSesion = 1, IdVinculacion = 101, FechaSesion = new DateTime(2026, 10, 16), Horas = 4 });
            bloque.Agregar(new SesionListItem { IdSesion = 2, IdVinculacion = 101, FechaSesion = new DateTime(2026, 10, 16), Horas = 4 });
            bloque.Agregar(new SesionListItem { IdSesion = 3, IdVinculacion = 102, FechaSesion = new DateTime(2026, 10, 17), Horas = 4 });
            bloque.Agregar(new SesionListItem { IdSesion = 4, IdVinculacion = 0, FechaSesion = new DateTime(2026, 10, 17), Horas = 2 }); // Inválido

            // Act
            var ids = bloque.IdsVinculacion;

            // Assert
            ids.Should().HaveCount(2);
            ids.Should().Contain(new[] { 101, 102 });
        }

        [Fact]
        public void AlternarPernocta_WhenMutandoSesionesDelBloque_ShouldToggleEntre1Punto5Y2Punto5()
        {
            // Arrange: Bloque de 2 días consecutivos (16 y 17 de Octubre)
            var bloque = new ViaticoBloque();
            var s1 = new SesionListItem { IdSesion = 1, IdVinculacion = 50, FechaSesion = new DateTime(2026, 10, 16), Horas = 4, IsPernoctaUltimoDia = 0 };
            var s2 = new SesionListItem { IdSesion = 2, IdVinculacion = 50, FechaSesion = new DateTime(2026, 10, 17), Horas = 4, IsPernoctaUltimoDia = 0 };
            bloque.Agregar(s1);
            bloque.Agregar(s2);

            // Estado inicial: sin pernocta
            bloque.TienePernoctaUltimoDia.Should().BeFalse();
            bloque.DuracionTotal.Should().Be(1.5);
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 17));

            // Paso 1: Activar pernocta (simulando alternancia on)
            s2.IsPernoctaUltimoDia = 1;
            bloque.TienePernoctaUltimoDia.Should().BeTrue();
            bloque.DuracionTotal.Should().Be(2.5);
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 18));

            // Paso 2: Desactivar pernocta (simulando alternancia off)
            s2.IsPernoctaUltimoDia = 0;
            bloque.TienePernoctaUltimoDia.Should().BeFalse();
            bloque.DuracionTotal.Should().Be(1.5);
            bloque.FechaFinComision.Should().Be(new DateTime(2026, 10, 17));
        }
    }
}
