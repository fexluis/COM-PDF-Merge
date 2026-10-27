using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class DiasHorasSemanaFormatterTests
    {
        [Fact]
        public void DesdeSesionListItems_CarrilPregrado_EmiteDiaHorasYHorario()
        {
            // Arrange (Fase 1136 / Fase 1138): Dos sesiones en días distintos de 4h con su horario
            var items = new List<SesionListItem>
            {
                new SesionListItem
                {
                    DiaSemana = "MARTES",
                    Horas = 4.0,
                    HorarioSesion = "08:00-12:00",
                    FechaSesion = new DateTime(2026, 3, 3)
                },
                new SesionListItem
                {
                    DiaSemana = "MIERCOLES",
                    Horas = 4.0,
                    HorarioSesion = "08:00-12:00",
                    FechaSesion = new DateTime(2026, 3, 4)
                }
            };

            // Act
            string resultado = DiasHorasSemanaFormatter.DesdeSesionListItems(items);

            // Assert
            resultado.Should().Be("MARTES|4|08:00-12:00, MIERCOLES|4|08:00-12:00");
        }

        [Fact]
        public void DesdeSesionListItems_Duplicados_NoRepitePar()
        {
            // Arrange (Fase 1136 / Fase 1138): Múltiples sesiones semanales repetidas para el mismo día y horario
            var items = new List<SesionListItem>
            {
                new SesionListItem { DiaSemana = "JUEVES", Horas = 4.0, HorarioSesion = "18:00-22:00", FechaSesion = new DateTime(2026, 3, 5) },
                new SesionListItem { DiaSemana = "JUEVES", Horas = 4.0, HorarioSesion = "18:00-22:00", FechaSesion = new DateTime(2026, 3, 12) },
                new SesionListItem { DiaSemana = "JUEVES", Horas = 4.0, HorarioSesion = "18:00-22:00", FechaSesion = new DateTime(2026, 3, 19) }
            };

            // Act
            string resultado = DiasHorasSemanaFormatter.DesdeSesionListItems(items);

            // Assert
            resultado.Should().Be("JUEVES|4|18:00-22:00");
        }

        [Fact]
        public void DesdeSesionListItems_SinDia_DevuelveVacio()
        {
            // Arrange (Fase 1136 / Fase 1138): Sesión sin DiaSemana y sin FechaSesion
            var items = new List<SesionListItem>
            {
                new SesionListItem { DiaSemana = null, FechaSesion = null, Horas = 4.0, HorarioSesion = "08:00-12:00" }
            };

            // Act
            string resultado = DiasHorasSemanaFormatter.DesdeSesionListItems(items);

            // Assert
            resultado.Should().BeEmpty();
        }

        [Fact]
        public void DesdeSesionListItems_ColeccionNula_DevuelveVacio()
        {
            // Act
            string resultado = DiasHorasSemanaFormatter.DesdeSesionListItems(null);

            // Assert
            resultado.Should().BeEmpty();
        }

        [Fact]
        public void DesdeSesiones_SesionesDominio_EmiteFormatoCanonico()
        {
            // Arrange (Fase 941 / Fase 991 / Fase 1138): Sesiones de dominio Sesion
            var sesiones = new List<Sesion>
            {
                new Sesion { FechaSesion = new DateTime(2026, 3, 2), Horas = 4.0, HorarioSesion = "08:00-12:00" },
                new Sesion { FechaSesion = new DateTime(2026, 3, 4), Horas = 4.0, HorarioSesion = "08:00-12:00" }
            };

            // Act
            string resultado = DiasHorasSemanaFormatter.DesdeSesiones(sesiones);

            // Assert
            resultado.Should().Be("LUNES|4|08:00-12:00, MIERCOLES|4|08:00-12:00");
        }
    }
}
