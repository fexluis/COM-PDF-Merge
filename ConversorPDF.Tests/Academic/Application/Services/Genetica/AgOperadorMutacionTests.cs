using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgOperadorMutacionTests
    {
        [Fact]
        public void Mutar_WhenGenIsFijado_ShouldNotMutateGen()
        {
            // Arrange
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var genFijo = new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAula = 5,
                IdBloqueHorario = 2,
                Estado = AgEstadoGen.Fijado
            };
            individuo.Genes.Add(genFijo);

            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 99, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 99, Activo = 1, DuracionHoras = 4, DiaSemanaNumero = 1 });

            // Act: tasa de mutación = 1.0 (siempre mutaría si no estuviera fijado)
            sut.Mutar(individuo, 1.0, contexto);

            // Assert
            genFijo.IdTercero.Should().Be(10);
            genFijo.IdAula.Should().Be(5);
            genFijo.IdBloqueHorario.Should().Be(2);
        }

        [Fact]
        public void Mutar_WhenGenIsValid_ShouldMutateWhenRateIsOne()
        {
            // Arrange
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 1,
                IdTercero = 10,
                IdAula = 5,
                IdBloqueHorario = 2,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 99, Estado = AgEstadoAula.Disponible });
            var bloqueNuevo = new AgBloqueHorario { Id = 88, Activo = 1, DuracionHoras = 4, DiaSemanaNumero = 3 };
            contexto.Bloques.Add(bloqueNuevo);

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert
            gen.IdAula.Should().Be(99);
            gen.IdBloqueHorario.Should().Be(88);
            gen.DiaSemana.Should().Be(3);
        }

        [Fact]
        public void Mutar_WhenGenIsVacante_ShouldPreserveIdTerceroVacante()
        {
            // Arrange (Fase 463 / Fase 470)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var genVacante = new AgGen
            {
                Id = 1,
                IdAsignatura = 1,
                IdTercero = 192,
                IdAula = 5,
                IdBloqueHorario = 2,
                EsVacante = 1,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(genVacante);

            var contexto = new AgContexto();
            contexto.IdTerceroVacante = 192;
            contexto.PoolPorAsignatura[1] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 50, Activo = 1 }
            };

            // Act: mutación forzada al 100%
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: no muta a docente 50; preserva centinela 192
            genVacante.IdTercero.Should().Be(192);
        }

        [Fact]
        public void Mutar_WhenMutatingAula_ShouldOnlySelectCandidatesWithMatchingSemestreAndGrupo()
        {
            // Arrange (Fase 486)
            var sut = new AgOperadorMutacion(new Random(42));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 1,
                IdTercero = 10,
                IdAula = 101,
                Grupo = "A",
                Semestre = "I",
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Aulas.Add(new AgAula { Id = 101, Semestre = "I", Grupo = "A", Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 102, Semestre = "I", Grupo = "A", Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 201, Semestre = "II", Grupo = "A", Estado = AgEstadoAula.Disponible });

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: solo debe mutar a aulas de Semestre I, Grupo A (101 o 102), nunca a 201
            gen.IdAula.Should().BeOneOf(101, 102);
            gen.IdAula.Should().NotBe(201);
            gen.Semestre.Should().Be("I");
        }

        [Fact]
        public void Mutar_WhenAulaHasArabicSemestre_ShouldNormalizeToRoman()
        {
            // Arrange (Fase 493)
            var sut = new AgOperadorMutacion(new Random(42));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 1,
                IdTercero = 10,
                IdAula = 101,
                Grupo = "A",
                Semestre = "I",
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            // Aula candidata registrada con formato arábigo "1"
            contexto.Aulas.Add(new AgAula { Id = 102, Semestre = "1", Grupo = "A", Estado = AgEstadoAula.Disponible });

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: gen.Semestre debe normalizarse a "I", nunca quedar "1"
            gen.IdAula.Should().Be(102);
            gen.Semestre.Should().Be("I");
        }

        [Fact]
        public void Mutar_WhenGenesInSameClaveGrupoHaveMixedTeachers_ShouldAlignInconditionally()
        {
            // Arrange (Fase 519)
            var sut = new AgOperadorMutacion(new Random(42));
            var individuo = new AgIndividuo();

            // Dos genes que pertenecen a la misma cohorte/asignatura multi-día pero tienen docentes dispares
            var gen1 = new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", IdTercero = 50, EsVacante = 0 };
            var gen2 = new AgGen { Id = 2, IdAsignatura = 10, IdAula = 1, Grupo = "A", ClaveGrupo = "10|1|A", IdTercero = 192, EsVacante = 1 };
            individuo.Genes.Add(gen1);
            individuo.Genes.Add(gen2);

            var contexto = new AgContexto();
            contexto.IdTerceroVacante = 192;

            // Act: mutación con tasa 0.0 (no debe seleccionar nuevo docente, pero sí alinear al canon del grupo)
            sut.Mutar(individuo, 0.0, contexto);

            // Assert: ambos genes deben compartir exactamente el mismo docente y estado de vacante
            gen1.IdTercero.Should().Be(50);
            gen2.IdTercero.Should().Be(50);
            gen2.EsVacante.Should().Be(0);
        }

        [Fact]
        public void Mutacion_PregradoConBloque8h_NoElevaDuracion()
        {
            // Arrange (Fase 1114.b / Fase 1114.d): Gen de aula de PREGRADO (4h) que llega con HorasSemanales = 8h.
            // La mutación de bloque debe buscar bloques según la duración del nivel (4h), nunca elevar a 8h.
            var sut = new AgOperadorMutacion(new Random(42));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 101,
                Grupo = "A",
                Semestre = "I",
                HorasSemanales = 8.0,
                IdBloqueHorario = 80,
                DiaSemana = 1,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.AulasPorId[101] = new AgAula { Id = 101, IdPrograma = 1, HorasDia = 4, Estado = AgEstadoAula.Disponible };
            contexto.Bloques.Add(new AgBloqueHorario { Id = 40, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 80, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 8.0 });

            // Act: forzar mutación
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: se debe asignar el bloque de 4h y no de 8h
            gen.IdBloqueHorario.Should().Be(40);
        }

        [Fact]
        public void Mutacion_AulaTarde_NoMigraAManana()
        {
            // Arrange (Fase 1132 / Fase 1134): Aula con Jornada = "TARDE".
            // Bloques disponibles: Id 10 (MANANA) e Id 20 (TARDE) ambos de 4h.
            // La mutación debe filtrar por franjaObjetivo ("TARDE") y nunca asignar Id 10.
            var sut = new AgOperadorMutacion(new Random(42));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 101,
                Grupo = "A",
                Semestre = "I",
                HorasSemanales = 4.0,
                IdBloqueHorario = 20,
                DiaSemana = 1,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.AulasPorId[101] = new AgAula { Id = 101, IdPrograma = 1, HorasDia = 4, Jornada = "TARDE", Estado = AgEstadoAula.Disponible };
            contexto.Bloques.Add(new AgBloqueHorario { Id = 10, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "MANANA" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 20, Activo = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0, Franja = "TARDE" });

            // Act: forzar mutación
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: debe mantener o reasignar a bloque de franja TARDE (Id 20)
            gen.IdBloqueHorario.Should().Be(20);
        }
    }
}
