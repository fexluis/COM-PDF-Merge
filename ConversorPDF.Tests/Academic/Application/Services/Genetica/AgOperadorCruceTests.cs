using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgOperadorCruceTests
    {
        [Fact]
        public void Cruzar_WhenCalled_ShouldPreserveIdAulaAndSwapBlockAndDates()
        {
            // Arrange
            var random = new Random(42);
            var asignador = new AgAsignadorTercero(random);
            var sut = new AgOperadorCruce(random, asignador);
            var contexto = new AgContexto();

            var p1 = new AgIndividuo();
            var p2 = new AgIndividuo();

            for (int i = 1; i <= 4; i++)
            {
                p1.Genes.Add(new AgGen
                {
                    Id = i,
                    IdAsignatura = i,
                    IdAula = 100, // Aula fija de la corrida
                    IdTercero = 10 + i,
                    IdBloqueHorario = 1,
                    FechaInicioEstimada = new DateTime(2026, 2, 1),
                    FechaFinEstimada = new DateTime(2026, 3, 1),
                    HorasSemanales = 4,
                    EsFijo = 0
                });

                p2.Genes.Add(new AgGen
                {
                    Id = i,
                    IdAsignatura = i,
                    IdAula = 100, // Aula fija de la corrida
                    IdTercero = 20 + i,
                    IdBloqueHorario = 2,
                    FechaInicioEstimada = new DateTime(2026, 3, 15),
                    FechaFinEstimada = new DateTime(2026, 4, 15),
                    HorasSemanales = 6,
                    EsFijo = 0
                });
            }

            // Act
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            h1.Genes.Should().HaveCount(4);
            h2.Genes.Should().HaveCount(4);

            // Verify IdAula is invariant across all genes
            h1.Genes.Should().OnlyContain(g => g.IdAula == 100);
            h2.Genes.Should().OnlyContain(g => g.IdAula == 100);
        }

        [Fact]
        public void Cruzar_WhenHijoModified_ShouldNotMutateParents()
        {
            // Arrange
            var random = new Random(42);
            var asignador = new AgAsignadorTercero(random);
            var sut = new AgOperadorCruce(random, asignador);
            var contexto = new AgContexto();

            var p1 = new AgIndividuo();
            var p2 = new AgIndividuo();

            p1.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdTercero = 100, IdBloqueHorario = 1 });
            p1.Genes.Add(new AgGen { Id = 2, IdAsignatura = 20, IdTercero = 200, IdBloqueHorario = 2 });

            p2.Genes.Add(new AgGen { Id = 3, IdAsignatura = 10, IdTercero = 300, IdBloqueHorario = 3 });
            p2.Genes.Add(new AgGen { Id = 4, IdAsignatura = 20, IdTercero = 400, IdBloqueHorario = 4 });

            // Act
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Mutate child genes
            h1.Genes[0].IdTercero = 999;
            h2.Genes[0].IdTercero = 888;

            // Assert
            p1.Genes[0].IdTercero.Should().Be(100);
            p2.Genes[0].IdTercero.Should().Be(300);
        }

        [Fact]
        public void Cruzar_WhenExecutedWithContext_ShouldPreserveGroupHoursInvariants()
        {
            // Arrange
            var mockRandom = new Mock<Random>();
            mockRandom.Setup(r => r.NextDouble()).Returns(0.0); // fuerza cruce
            mockRandom.Setup(r => r.Next(It.IsAny<int>(), It.IsAny<int>())).Returns(0);

            var sut = new AgOperadorCruce(mockRandom.Object);
            var contexto = new AgContexto();

            var p1 = new AgIndividuo();
            p1.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 0 });
            p1.Genes.Add(new AgGen { Id = 2, IdAsignatura = 11, IdAula = 1, Grupo = "A", ClaveGrupo = "11|1|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 1 });
            p1.Genes.Add(new AgGen { Id = 3, IdAsignatura = 11, IdAula = 1, Grupo = "A", ClaveGrupo = "11|1|A", TotalHorasAsignatura = 48.0, HorasGen = 0.0, IndiceDiaPatron = 2 });

            var p2 = new AgIndividuo();
            p2.Genes.Add(new AgGen { Id = 4, IdAsignatura = 11, IdAula = 1, Grupo = "A", ClaveGrupo = "11|1|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 0 });
            p2.Genes.Add(new AgGen { Id = 5, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", TotalHorasAsignatura = 48.0, HorasGen = 24.0, IndiceDiaPatron = 1 });
            p2.Genes.Add(new AgGen { Id = 6, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", TotalHorasAsignatura = 48.0, HorasGen = 0.0, IndiceDiaPatron = 2 });

            // Act
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            h1.Genes.GroupBy(g => g.ClaveGrupo).Should().OnlyContain(grp => Math.Abs(grp.Sum(g => g.HorasGen) - grp.Max(g => g.TotalHorasAsignatura)) < 0.001);
            h2.Genes.GroupBy(g => g.ClaveGrupo).Should().OnlyContain(grp => Math.Abs(grp.Sum(g => g.HorasGen) - grp.Max(g => g.TotalHorasAsignatura)) < 0.001);

            // Idempotencia: segunda invocación de Sellar no altera las horas ya equilibradas
            var horasAntes = h1.Genes.Select(g => g.HorasGen).ToList();
            AgCoherenciaGrupo.Sellar(h1, contexto);
            var horasDespues = h1.Genes.Select(g => g.HorasGen).ToList();
            horasDespues.Should().Equal(horasAntes);
        }

        [Fact]
        public void Cruzar_DescendienteConOleada_NoHeredaFechasDelProgenitor()
        {
            // Arrange (Fase 1072): El descendiente no debe heredar las fechas del progenitor si ya tiene fechas de su oleada
            var mockRandom = new Mock<Random>();
            mockRandom.Setup(r => r.NextDouble()).Returns(0.0); // fuerza cruce
            mockRandom.Setup(r => r.Next(It.IsAny<int>(), It.IsAny<int>())).Returns(0);

            var sut = new AgOperadorCruce(mockRandom.Object);
            var contexto = new AgContexto();

            var p1 = new AgIndividuo();
            var fechaIniP1 = new DateTime(2026, 2, 1);
            var fechaFinP1 = new DateTime(2026, 3, 31);
            p1.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 1,
                Grupo = "A",
                ClaveGrupo = "10|1|A",
                FechaInicioEstimada = fechaIniP1,
                FechaFinEstimada = fechaFinP1
            });

            var p2 = new AgIndividuo();
            var fechaIniP2 = new DateTime(2026, 4, 1);
            var fechaFinP2 = new DateTime(2026, 5, 31);
            p2.Genes.Add(new AgGen
            {
                Id = 2,
                IdAsignatura = 10,
                IdAula = 1,
                Grupo = "A",
                ClaveGrupo = "10|1|A",
                FechaInicioEstimada = fechaIniP2,
                FechaFinEstimada = fechaFinP2
            });

            // Act: al cruzar, h1 (clonado de p1) no debe ser sobreescrito con las fechas de p2
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            h1.Genes[0].FechaInicioEstimada.Should().Be(fechaIniP1);
            h1.Genes[0].FechaFinEstimada.Should().Be(fechaFinP1);
            h2.Genes[0].FechaInicioEstimada.Should().Be(fechaIniP2);
            h2.Genes[0].FechaFinEstimada.Should().Be(fechaFinP2);
        }

        [Fact]
        public void Cruzar_SinBloqueDeUnidad_GenIncompletoSinBloque()
        {
            // Arrange (Fase 1086): Si no hay bloques de la duración esperada (unidadCruce=4),
            // el operador de cruce nunca debe asignar un bloque de 8h como último recurso;
            // debe dejar IdBloqueHorario = 0 y Estado = Incompleto.
            var mockRandom = new Mock<Random>();
            mockRandom.Setup(r => r.NextDouble()).Returns(0.0); // fuerza cruce
            mockRandom.Setup(r => r.Next(It.IsAny<int>(), It.IsAny<int>())).Returns(0);

            var sut = new AgOperadorCruce(mockRandom.Object);
            var contexto = new AgContexto();
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.AulasPorId[1] = new AgAula { Id = 1, IdPrograma = 1, HorasDia = 4, DiasVentana = "LUN-JUE" };

            // En el contexto solo existe un bloque de 8h (ajeno al nivel de pregrado)
            contexto.Bloques.Add(new AgBloqueHorario { Id = 99, DiaSemanaNumero = 1, DuracionHoras = 8.0, Activo = 1 });

            var p1 = new AgIndividuo();
            p1.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", DiasPatron = "1,2", HorasSemanales = 4 });
            var p2 = new AgIndividuo();
            p2.Genes.Add(new AgGen { Id = 2, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", DiasPatron = "1,2", HorasSemanales = 4 });

            // Act
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            h1.Genes[0].IdBloqueHorario.Should().Be(0);
            h1.Genes[0].Estado.Should().Be(AgEstadoGen.Incompleto);
            h1.Genes[0].HorasSemanales.Should().Be(4.0);
        }

        [Fact]
        public void Cruzar_AulaTarde_AsignaBloqueDeFranjaTarde()
        {
            // Arrange (Fase 1131 / Fase 1134): Aula con Jornada = "TARDE".
            // Existen dos bloques de 4h en el mismo día (Id 10 MANANA, Id 20 TARDE).
            // El cruce debe asignar el bloque correspondiente a la jornada TARDE (Id 20).
            var mockRandom = new Mock<Random>();
            mockRandom.Setup(r => r.NextDouble()).Returns(0.0); // fuerza cruce
            mockRandom.Setup(r => r.Next(It.IsAny<int>(), It.IsAny<int>())).Returns(0);

            var sut = new AgOperadorCruce(mockRandom.Object);
            var contexto = new AgContexto();
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.AulasPorId[1] = new AgAula { Id = 1, IdPrograma = 1, HorasDia = 4, DiasVentana = "LUN,MIE", Jornada = "TARDE" };

            contexto.Bloques.Add(new AgBloqueHorario { Id = 10, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "MANANA", Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 20, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "TARDE", Activo = 1 });

            var p1 = new AgIndividuo();
            p1.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", DiasPatron = "1,3", HorasSemanales = 4, IndiceDiaPatron = 0 });
            var p2 = new AgIndividuo();
            p2.Genes.Add(new AgGen { Id = 2, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", DiasPatron = "1,3", HorasSemanales = 4, IndiceDiaPatron = 0 });

            // Act
            var (h1, h2) = sut.Cruzar(p1, p2, tasaCruce: 1.0, generacion: 1, contexto: contexto);

            // Assert
            h1.Genes[0].IdBloqueHorario.Should().Be(20);
        }
    }
}
