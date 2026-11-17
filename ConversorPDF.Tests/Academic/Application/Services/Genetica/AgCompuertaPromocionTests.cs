using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgCompuertaPromocionTests
    {
        [Fact]
        public void EsPromovible_WhenIndividuoTieneVacantesYNoSePermiten_ShouldRetornarFalse()
        {
            // Arrange
            var corrida = new AgCorrida { Id = 1, Vacantes = 0 };
            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen>
                {
                    new AgGen { EsVacante = 1, IdTercero = 0 },
                    new AgGen { EsVacante = 1, IdTercero = 0 }
                }
            };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());

            // Act
            bool resultado = AgCompuertaPromocion.EsPromovible(corrida, ind, sesiones, contexto: null, permitirVacantes: false);

            // Assert
            Assert.False(resultado);
        }

        [Fact]
        public void EsPromovible_WhenIndividuoTieneVacantesYSePermiten_ShouldValidarSinRechazarPorVacante()
        {
            // Arrange
            var corrida = new AgCorrida { Id = 1, Vacantes = 2 };
            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen>
                {
                    new AgGen { EsVacante = 1, IdTercero = 0 },
                    new AgGen { EsVacante = 1, IdTercero = 0 }
                }
            };
            // Sin genes válidos completos ni sesiones, falla en ValidarCobertura ("No hay genes en el individuo" o sesión), no en ValidarPromocion ("vacantes pendientes")
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());

            // Act: con permitirVacantes: true, la compuerta no rechaza por vacantes
            var valResult = AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto: null, permitirVacantes: true);

            // Assert: no debe fallar por vacantes pendientes ni por genes vacantes (Fase 1876)
            Assert.True(valResult.IsFailure);
            Assert.DoesNotContain("vacante", valResult.Error, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void EsPromovible_WhenCorridaTieneParametrosJsonConVacantesConsentidas_ShouldPermitirSinFlagExplicito()
        {
            // Arrange (Fase 1874)
            var corrida = new AgCorrida
            {
                Id = 1,
                Vacantes = 2,
                ParametrosJson = "{\"vacantesConsentidas\":[101,102]}"
            };
            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen>
                {
                    new AgGen { EsVacante = 1, IdTercero = 0 },
                    new AgGen { EsVacante = 1, IdTercero = 0 }
                }
            };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());

            // Act: pasar permitirVacantes: false explícito, pero la corrida tiene vacantes consentidas persistidas
            var valResult = AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto: null, permitirVacantes: false);

            // Assert: se rehidrata el consentimiento desde ParametrosJson y la validación no se detiene por vacantes
            Assert.True(valResult.IsFailure);
            Assert.DoesNotContain("vacante", valResult.Error, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void EsPromovible_WhenSolucionValidaConVacanteConsentida_ShouldRetornarTrue()
        {
            // Arrange (Fase 1876): Gen + Sesión en staging en vacante consentida, consistente en fechas, horas y bloques
            var fecha = new DateTime(2027, 2, 5);
            var bloque = new AgBloqueHorario
            {
                Id = 1,
                Activo = 1,
                DiaSemanaNumero = fecha.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)fecha.DayOfWeek,
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4.0,
                Franja = "MANANA"
            };

            var gen = new AgGen
            {
                Id = 10,
                IdAula = 1,
                IdAsignatura = 101,
                IdTercero = 0,
                EsVacante = 1,
                IdBloqueHorario = 1,
                HorasGen = 4.0,
                HorasSemanales = 4.0,
                FechasGen = new List<DateTime> { fecha }
            };

            var sesion = new AgSesion
            {
                Id = 1,
                IdGen = 10,
                IdAula = 1,
                IdTercero = 0,
                EsVacante = 1,
                FechaSesion = fecha,
                HorarioSesion = "08:00 A 12:00",
                Horas = 4.0
            };

            var contexto = new AgContexto
            {
                Ventana = VentanaSemestre.Crear(1, "2027-1", fecha, fecha, AgOrigenVentana.Default).Value,
                AsignaturasVacantesConsentidas = new HashSet<int> { 101 }
            };
            contexto.Bloques.Add(bloque);
            contexto.ConstruirIndices();

            var corrida = new AgCorrida
            {
                Id = 1,
                Vacantes = 1,
                ParametrosJson = "{\"vacantesConsentidas\":[101]}"
            };

            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen> { gen }
            };

            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion> { sesion });

            // Act: evaluar con permitirVacantes: false (debe deducirlo de ParametrosJson)
            bool esPromovible = AgCompuertaPromocion.EsPromovible(corrida, ind, sesiones, contexto, permitirVacantes: false);

            // Assert
            Assert.True(esPromovible);
        }

        [Fact]
        public void EsPromovible_WhenCorridaTieneVacantesConsentidasTipadas_NoRequiereParametrosJson()
        {
            // Arrange (Fase 1885/1886): Corrida con VacantesConsentidas tipadas en memoria/BD y ParametrosJson vacío
            var fecha = new DateTime(2027, 2, 5);
            var bloque = new AgBloqueHorario
            {
                Id = 1,
                Activo = 1,
                DiaSemanaNumero = fecha.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)fecha.DayOfWeek,
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4.0,
                Franja = "MANANA"
            };

            var gen = new AgGen
            {
                Id = 10,
                IdAula = 1,
                IdAsignatura = 101,
                IdTercero = 0,
                EsVacante = 1,
                IdBloqueHorario = 1,
                HorasGen = 4.0,
                HorasSemanales = 4.0,
                FechasGen = new List<DateTime> { fecha }
            };

            var sesion = new AgSesion
            {
                Id = 1,
                IdGen = 10,
                IdAula = 1,
                IdTercero = 0,
                EsVacante = 1,
                FechaSesion = fecha,
                HorarioSesion = "08:00 A 12:00",
                Horas = 4.0
            };

            var contexto = new AgContexto
            {
                Ventana = VentanaSemestre.Crear(1, "2027-1", fecha, fecha, AgOrigenVentana.Default).Value
            };
            contexto.Bloques.Add(bloque);
            contexto.ConstruirIndices();

            var corrida = new AgCorrida
            {
                Id = 1,
                Vacantes = 1,
                ParametrosJson = "{}", // Sin vacantes en el JSON
                VacantesConsentidas = new HashSet<int> { 101 } // Pobladas en la propiedad tipada
            };

            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen> { gen }
            };

            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion> { sesion });

            // Act: evaluar con permitirVacantes: false (debe deducirlo de la propiedad tipada sin parsear JSON)
            bool esPromovible = AgCompuertaPromocion.EsPromovible(corrida, ind, sesiones, contexto, permitirVacantes: false);

            // Assert
            Assert.True(esPromovible);
            Assert.Contains(101, contexto.AsignaturasVacantesConsentidas);
        }
    }
}
