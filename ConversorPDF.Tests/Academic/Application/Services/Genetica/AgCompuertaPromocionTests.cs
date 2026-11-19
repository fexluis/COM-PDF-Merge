using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
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
                    new AgGen { IdAsignatura = 101, EsVacante = 1, IdTercero = 0 },
                    new AgGen { IdAsignatura = 102, EsVacante = 1, IdTercero = 0 }
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

        [Fact]
        public void ValidarPromocion_WhenNoConsentidasYAutorizadas_ShouldLogAndPass()
        {
            // Arrange (Fase 2032 & 2035)
            var corrida = new AgCorrida { Id = 1, Vacantes = 1 };
            var gen = new AgGen { Id = 1, IdAsignatura = 10, EsVacante = 1, IdTercero = 0 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen> { gen } };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            var autorizadas = new HashSet<int> { 10 };

            // Act
            var res = AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto, permitirVacantes: true, autorizacionVacantes: autorizadas);

            // Assert: la compuerta de vacantes pasa (no falla por no consentidas ni falta de autorización)
            if (res.IsFailure)
            {
                res.Error.Should().NotContain("La autorización no cubre las asignaturas");
                res.Error.Should().NotContain("vacantes no consentidas");
            }
        }

        [Fact]
        public void ValidarPromocion_WhenNoConsentidasYAutorizacionIncompleta_ShouldFail()
        {
            // Arrange (Fase 2032 & 2035): Asignatura 10 no consentida y no cubierta por autorización
            var corrida = new AgCorrida { Id = 1, Vacantes = 1 };
            var gen = new AgGen { Id = 1, IdAsignatura = 10, EsVacante = 1, IdTercero = 0 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen> { gen } };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            var autorizadas = new HashSet<int> { 99 }; // Cubre 99, pero falta 10

            // Act
            var res = AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto, permitirVacantes: true, autorizacionVacantes: autorizadas);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("La autorización no cubre las asignaturas: 10");
        }

        [Fact]
        public void ValidarPromocion_WhenConsentimientoPorClaveSede_ShouldNotFailByIndividuoVacantes()
        {
            // Arrange (Fase 2033 & 2035): Consentimiento granular por clave "10:1"
            var corrida = new AgCorrida { Id = 1, Vacantes = 1 };
            var gen = new AgGen { Id = 1, IdAsignatura = 10, IdSede = 1, EsVacante = 1, IdTercero = 0 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen> { gen } };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            contexto.ClavesVacantesConsentidas.Add("10:1");

            // Act: permitirVacantes = false explícito, pero la clave está consentida en el contexto
            var res = AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto, permitirVacantes: false);

            // Assert: no debe fallar por "sin consentimiento ni autorización"
            if (res.IsFailure)
            {
                res.Error.Should().NotContain("sin consentimiento ni autorización");
                res.Error.Should().NotContain("vacantes no consentidas");
            }
        }

        [Fact]
        public void EsPromovible_WhenVacantesEnMultiplesAsignaturasYAutorizacionParcial_ShouldReturnFalse()
        {
            // Arrange (Fase 2037): Individuo con vacantes en asig 1 y 36; autorización solo cubre 1
            var corrida = new AgCorrida { Id = 1, Vacantes = 2 };
            var ind = new AgIndividuo
            {
                Id = 1,
                HardViolations = 0,
                Genes = new List<AgGen>
                {
                    new AgGen { Id = 1, IdAsignatura = 1, EsVacante = 1, IdTercero = 0 },
                    new AgGen { Id = 2, IdAsignatura = 36, EsVacante = 1, IdTercero = 0 }
                }
            };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            var autorizadas = new HashSet<int> { 1 }; // Solo 1 autorizada, falta 36

            // Act
            bool esPromovible = AgCompuertaPromocion.EsPromovible(corrida, ind, sesiones, contexto, permitirVacantes: true, autorizacionVacantes: autorizadas);

            // Assert
            esPromovible.Should().BeFalse();
        }

        [Theory]
        [InlineData("asig=10,20,30", new[] { 10, 20, 30 })]
        [InlineData("asignaturas: 5, 15", new[] { 5, 15 })]
        [InlineData("materias: 7, 14", new[] { 7, 14 })]
        [InlineData("Promoción parcial con vacantes autorizadas (asignaturas: 1, 5, 36).", new[] { 1, 5, 36 })]
        [InlineData("autorizo 5 grupos para el ciclo 2026 con 3 docentes", new int[0])]
        [InlineData("", new int[0])]
        [InlineData(null, new int[0])]
        public void ParseAutorizacionVacantes_ShouldExtractExpectedIds(string justificacion, int[] esperados)
        {
            // Act
            var res = AgPromotorVinculaciones.ParseAutorizacionVacantes(justificacion);

            // Assert
            res.Asigs.Should().BeEquivalentTo(esperados);
        }

        [Fact]
        public void ParseAutorizacionVacantes_WithGranularKeys_ShouldExtractBothAsigsAndClaves()
        {
            // Arrange (Fase 2045): justificación con mezcla de tokens ID y claves asig:sede
            string justificacion = "Promoción autorizada (asignaturas: 1:1, 36:2, 10).";

            // Act
            var (asigs, claves) = AgPromotorVinculaciones.ParseAutorizacionVacantes(justificacion);

            // Assert
            asigs.Should().BeEquivalentTo(new[] { 10 });
            claves.Should().BeEquivalentTo(new[] { "1:1", "36:2" });
        }

        [Fact]
        public void ValidarPromocion_WhenAutorizacionPorClaveGranular_ShouldPassCoveredGene()
        {
            // Arrange (Fase 2045): Gen con Asig 1 en Sede 2 cubierto por clave "1:2"
            var corrida = new AgCorrida { Id = 1, Vacantes = 1 };
            var gen = new AgGen { Id = 1, IdAsignatura = 1, IdSede = 2, EsVacante = 1, IdTercero = 0 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen> { gen } };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            var autorizadasClaves = new HashSet<string> { "1:2" };

            // Act
            var res = AgCompuertaPromocion.ValidarPromocion(
                corrida, ind, sesiones, contexto, permitirVacantes: true,
                autorizacionVacantes: new HashSet<int>(), autorizacionVacantesClaves: autorizadasClaves);

            // Assert
            if (res.IsFailure)
            {
                res.Error.Should().NotContain("La autorización no cubre las asignaturas");
                res.Error.Should().NotContain("vacantes no consentidas");
            }
        }

        [Fact]
        public void ParseAutorizacionVacantes_WhenTokenHasZeroSede_ShouldExtractAsAsig()
        {
            // Arrange (Fase 2047): Token "10:0" (sede no resuelta) no debe descartarse sino mapearse a ID de asignatura
            string justificacion = "Promoción autorizada (asignaturas: 10:0, 20:1).";

            // Act
            var (asigs, claves) = AgPromotorVinculaciones.ParseAutorizacionVacantes(justificacion);

            // Assert
            asigs.Should().Contain(10);
            claves.Should().Contain("20:1");
        }

        [Fact]
        public void ValidarPromocion_WhenVacanteConSedeCeroYAutorizada_ShouldPass()
        {
            // Arrange (Fase 2047): Gen vacante con IdSede = 0 autorizado vía ID de asignatura
            var corrida = new AgCorrida { Id = 1, Vacantes = 1 };
            var gen = new AgGen { Id = 1, IdAsignatura = 10, IdSede = 0, EsVacante = 1, IdTercero = 0 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen> { gen } };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            var autorizadas = new HashSet<int> { 10 };

            // Act
            var res = AgCompuertaPromocion.ValidarPromocion(
                corrida, ind, sesiones, contexto, permitirVacantes: true, autorizacionVacantes: autorizadas);

            // Assert
            if (res.IsFailure)
            {
                res.Error.Should().NotContain("La autorización no cubre las asignaturas");
                res.Error.Should().NotContain("vacantes no consentidas");
            }
        }

        [Fact]
        public void ValidarPromocion_ShouldNotMutarContextoClavesVacantesConsentidas()
        {
            // Arrange (Fase 2049): ValidarPromocion debe ser de solo lectura y no mutar contexto
            var corrida = new AgCorrida
            {
                Id = 1,
                ParametrosJson = "{\"vacantesConsentidas\":[\"10:1\"]}"
            };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0, Genes = new List<AgGen>() };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion>());
            var contexto = new AgContexto();
            contexto.ClavesVacantesConsentidas.Clear();

            // Act
            AgCompuertaPromocion.ValidarPromocion(corrida, ind, sesiones, contexto, permitirVacantes: false);

            // Assert: contexto permanece inmutable
            contexto.ClavesVacantesConsentidas.Should().BeEmpty();
        }

        [Fact]
        public void EsPromovibleSolicitable_WhenViable_ShouldReturnTrue()
        {
            // Arrange (Fase 2044): Corrida e individuo válidos con staging no vacío
            var corrida = new AgCorrida { Id = 1 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 0 };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion> { new AgSesion { Id = 1 } });

            // Act
            bool solicitable = AgCompuertaPromocion.EsPromovibleSolicitable(corrida, ind, sesiones);

            // Assert
            solicitable.Should().BeTrue();
        }

        [Fact]
        public void EsPromovibleSolicitable_WhenHardViolationsPresent_ShouldReturnFalse()
        {
            // Arrange (Fase 2044): Violaciones HARD bloquean solicitabilidad sin evaluar compuerta
            var corrida = new AgCorrida { Id = 1 };
            var ind = new AgIndividuo { Id = 1, HardViolations = 2 };
            var sesiones = Result.Success<IReadOnlyCollection<AgSesion>>(new List<AgSesion> { new AgSesion { Id = 1 } });

            // Act
            bool solicitable = AgCompuertaPromocion.EsPromovibleSolicitable(corrida, ind, sesiones);

            // Assert
            solicitable.Should().BeFalse();
        }
    }
}
