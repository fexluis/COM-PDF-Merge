using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests
{
    public class AgValidadorVentanaTests
    {
        private List<AgBloqueHorario> CrearBloquesSemilla()
        {
            var bloques = new List<AgBloqueHorario>();
            int id = 1;
            // 7 días, 6 bloques/día
            for (int d = 1; d <= 7; d++)
            {
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "08:00", HoraFin = "12:00", DuracionHoras = 4.0, Franja = "MANANA", Activo = 1 });
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "13:00", HoraFin = "17:00", DuracionHoras = 4.0, Franja = "TARDE", Activo = 1 });
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "18:00", HoraFin = "22:00", DuracionHoras = 4.0, Franja = "NOCHE", Activo = 1 });
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "08:00", HoraFin = "17:00", DuracionHoras = 8.0, Franja = "MANANA", Activo = 1 });
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "14:00", HoraFin = "22:00", DuracionHoras = 8.0, Franja = "TARDE", Activo = 1 });
                bloques.Add(new AgBloqueHorario { Id = id++, DiaSemanaNumero = d, HoraInicio = "14:00", HoraFin = "22:00", DuracionHoras = 8.0, Franja = "NOCHE", Activo = 1 });
            }
            return bloques;
        }

        [Fact]
        public void Validar_WhenDiasPatronContenidosYBloquesExisten_ShouldReturnSuccess()
        {
            // Arrange
            var bloques = CrearBloquesSemilla();

            // Act: MAR-VIE (2,3,4,5) contiene a los días martes y jueves (2,4)
            var res = AgValidadorVentana.Validar("2,4", "MAR-VIE", "TARDE", 8, bloques);

            // Assert
            res.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void Validar_WhenDiasPatronNoContenidosEnVentana_ShouldReturnFailure()
        {
            // Arrange
            var bloques = CrearBloquesSemilla();

            // Act: LUN-JUE (1,2,3,4) no contiene a viernes (5)
            var res = AgValidadorVentana.Validar("1,2,5", "LUN-JUE", "MANANA", 4, bloques);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no están contenidos en la ventana seleccionada");
        }

        [Fact]
        public void Validar_WhenBloqueNoExisteParaJornadaYHoras_ShouldReturnFailure()
        {
            // Arrange
            var bloques = CrearBloquesSemilla();
            // Desactivar bloques de tarde 8h
            foreach (var b in bloques)
            {
                if (b.Franja == "TARDE" && b.DuracionHoras == 8.0)
                {
                    b.Activo = 0;
                }
            }

            // Act
            var res = AgValidadorVentana.Validar("MAR-VIE", "MAR-VIE", "TARDE", 8, bloques);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("No existe ningún bloque horario activo");
        }
    }
}
