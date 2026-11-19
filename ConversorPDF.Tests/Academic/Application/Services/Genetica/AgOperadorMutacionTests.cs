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
        public void Mutar_WhenGenIsVacante_ShouldNormalizeIdTerceroToZero()
        {
            // Arrange (Fase 463 / Fase 470 / Fase 2014: normalización a IdTercero = 0)
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

            // Assert: no muta a docente 50 (sin habilitación); normaliza centinela a IdTercero = 0
            genVacante.IdTercero.Should().Be(0);
            genVacante.EsVacante.Should().Be(1);
        }

        [Fact]
        public void Mutar_WhenDocenteActualLibreYSinAlterno_ShouldPreserveDocenteActual()
        {
            // Arrange (Fase 2016: si el docente actual está disponible y no hay alterno, NO se libera a vacante)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdTercero = 25,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10),
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", Activo = 1 };
            contexto.RegistrarDocenteHabilitado(25, 1, 10, "I");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 25, Activo = 1 } // Solo está el actual, sin alterno
            };

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: no degenera a vacante; conserva al docente actual disponible
            gen.EsVacante.Should().Be(0);
            gen.IdTercero.Should().Be(25);
            gen.MotivoVacante.Should().BeEmpty();
        }

        [Fact]
        public void Mutar_WhenDocenteActualEnColisionYSinAlterno_ShouldMutateToVacanteExplicita()
        {
            // Arrange (Fase 2012 / Fase 2016: docente actual en colisión y sin alterno disponible => vacante explícita)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();

            // Gen 1 ocupa al docente 25 en bloque 1
            var gen1 = new AgGen
            {
                Id = 1,
                IdAsignatura = 5,
                IdTercero = 25,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10),
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };
            // Gen 2 en otra aula/asignatura colisiona con el docente 25 en el mismo bloque y fechas
            var gen2 = new AgGen
            {
                Id = 2,
                IdAsignatura = 10,
                IdTercero = 25,
                IdAula = 2,
                Grupo = "B",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10),
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen1);
            individuo.Genes.Add(gen2);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", Activo = 1 };
            contexto.RegistrarDocenteHabilitado(25, 1, 5, "I");
            contexto.RegistrarDocenteHabilitado(25, 1, 10, "I");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 25, Activo = 1 } // Sin alterno
            };

            // Act: mutación sobre gen2
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: al estar docente 25 colisionado y no haber alterno, gen2 pasa a vacante explícita
            gen2.EsVacante.Should().Be(1);
            gen2.IdTercero.Should().Be(0);
            gen2.MotivoVacante.Should().Be(AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo());
        }

        [Fact]
        public void Mutar_WhenGrupoTieneMotivoNoAsignable_ShouldIgnoreGrupo()
        {
            // Arrange (Fase 2017: guarda contra reversión de SIN_PLAN_OLEADA / SIN_BLOQUE_NIVEL)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdTercero = 0,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 0,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                Estado = AgEstadoGen.Incompleto
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 77, Activo = 1 }
            };

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: se ignora el grupo y se preserva intacto el motivo estructural
            gen.EsVacante.Should().Be(1);
            gen.IdTercero.Should().Be(0);
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_PLAN_OLEADA.Codigo());
        }

        [Fact]
        public void Mutar_WhenMotivoNoAsignableEnMinusculas_ShouldIgnoreGrupo()
        {
            // Arrange (Fase 2025: comparación case-insensitive para motivos no asignables)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdTercero = 0,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 0,
                EsVacante = 1,
                MotivoVacante = "sin_bloque_nivel",
                Estado = AgEstadoGen.Incompleto
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 77, Activo = 1 }
            };

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: se ignora el grupo y se preserva el motivo
            gen.EsVacante.Should().Be(1);
            gen.IdTercero.Should().Be(0);
            gen.MotivoVacante.Should().Be("sin_bloque_nivel");
        }

        [Fact]
        public void Mutar_WhenBloqueInactivoNoVerificable_ShouldNotMutateToVacante()
        {
            // Arrange (Fase 2025: grupo con bloque inactivo o no verificable no debe liberarse a vacante)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdTercero = 25,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 99, // Bloque inactivo o inexistente en contexto
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 25, Activo = 1 }
            };

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert: no se vacía ni enmascara R14; mantiene su asignación
            gen.EsVacante.Should().Be(0);
            gen.IdTercero.Should().Be(25);
        }

        [Fact]
        public void Mutar_WhenGrupoVacanteYSurgeAlterno_ShouldRecuperarDocente()
        {
            // Arrange (Fase 2015: movimiento inverso para recuperar docente en grupo vacante)
            var sut = new AgOperadorMutacion(new Random(1));
            var individuo = new AgIndividuo();
            var gen = new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdTercero = 0,
                IdAula = 1,
                Grupo = "A",
                Semestre = "I",
                IdPrograma = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10),
                EsVacante = 1,
                MotivoVacante = "SIN_DISPONIBILIDAD",
                Estado = AgEstadoGen.Valido
            };
            individuo.Genes.Add(gen);

            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", Activo = 1 };
            contexto.RegistrarDocenteHabilitado(77, 1, 10, "I");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 77, Activo = 1 }
            };

            // Act
            sut.Mutar(individuo, 1.0, contexto);

            // Assert
            gen.EsVacante.Should().Be(0);
            gen.IdTercero.Should().Be(77);
            gen.MotivoVacante.Should().BeEmpty();
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
