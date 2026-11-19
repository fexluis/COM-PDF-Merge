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
    public class AgConsentimientoVacantesTests
    {
        [Fact]
        public void Consentimiento_WhenProgramasHomonimos_ShouldOnlyIncludeDemandedSubjects()
        {
            // Arrange (Fase 2052): Materias 111-135 pertenecen a Programa 2 (catálogo) pero las aulas de la corrida son Programa 1 (37-42)
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            // Catálogo con materias de ambos programas
            var asig37 = new Asignatura { Id = 37, Nombre = "Materia Prog 1", IdPrograma = 1, Semestre = "I" };
            var asig111 = new Asignatura { Id = 111, Nombre = "Materia Prog 2", IdPrograma = 2, Semestre = "I" };
            ctx.Asignaturas.Add(asig37);
            ctx.Asignaturas.Add(asig111);
            ctx.AsignaturasPorId[37] = asig37;
            ctx.AsignaturasPorId[111] = asig111;

            // Demanda real de la corrida: únicamente aula con asignatura 37
            var aula1 = new AgAula { Id = 10, IdPrograma = 1, Semestre = "I", Grupo = "G1", IdSede = 5 };
            ctx.Aulas.Add(aula1);
            ctx.AulasPorId[10] = aula1;
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 10, IdAsignatura = 37 });

            // Ninguna tiene pool habilitado
            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert
            asigs.Should().Contain(37);
            asigs.Should().NotContain(111);
            claves.Should().Contain("37:5");
            claves.Should().NotContain("111:5");
            detalle.Should().ContainSingle(d => d.IdAsignatura == 37 && d.IdSede == 5);
        }

        [Fact]
        public void Consentimiento_WhenPoolHabilitado_ShouldNotConsent()
        {
            // Arrange (Fase 2052): Materia con docente activo y habilitado para el programa y semestre del aula
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig = new Asignatura { Id = 40, Nombre = "Matemáticas", IdPrograma = 1, Semestre = "II" };
            ctx.Asignaturas.Add(asig);
            ctx.AsignaturasPorId[40] = asig;

            var aula = new AgAula { Id = 20, IdPrograma = 1, Semestre = "II", Grupo = "A", IdSede = 3 };
            ctx.Aulas.Add(aula);
            ctx.AulasPorId[20] = aula;
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 20, IdAsignatura = 40 });

            // Registrar docente habilitado en pool
            ctx.PoolPorAsignatura[40] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { Id = 100, IdTercero = 999, IdAsignatura = 40, IdPrograma = 1, Semestre = "II", Activo = 1 }
            };
            ctx.RegistrarDocenteHabilitado(999, 1, 40, "II");

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert
            asigs.Should().BeEmpty();
            claves.Should().BeEmpty();
            detalle.Should().BeEmpty();
        }

        [Fact]
        public void Consentimiento_WhenPoolExistePeroSinDocenteHabilitado_ShouldMarkNoCumpleRequisitos()
        {
            // Arrange (Fase 2053): Docente existe en catálogo pero no para este semestre/programa
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig = new Asignatura { Id = 53, Nombre = "Derecho", IdPrograma = 1, Semestre = "III" };
            ctx.Asignaturas.Add(asig);
            ctx.AsignaturasPorId[53] = asig;

            var aula = new AgAula { Id = 30, IdPrograma = 1, Semestre = "III", Grupo = "U", IdSede = 2 };
            ctx.Aulas.Add(aula);
            ctx.AulasPorId[30] = aula;
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 30, IdAsignatura = 53 });

            // Existe en pool pero para otro programa o semestre
            ctx.PoolPorAsignatura[53] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { Id = 200, IdTercero = 888, IdAsignatura = 53, IdPrograma = 2, Semestre = "I", Activo = 1 }
            };
            ctx.RegistrarDocenteHabilitado(888, 2, 53, "I");

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert
            detalle.Should().ContainSingle();
            detalle[0].IdAsignatura.Should().Be(53);
            detalle[0].Motivo.Should().Be(AgMotivoVacante.NO_CUMPLE_REQUISITOS.Codigo());
        }

        [Fact]
        public void Consentimiento_WhenUnaCohorteSinDocenteYOtraConDocente_ShouldOnlyConsentClave()
        {
            // Arrange (Fase 2059): Asignatura 37 demandada en Aula 1 (Sede 5, sin docente) y Aula 2 (Sede 1, con docente)
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig37 = new Asignatura { Id = 37, Nombre = "Economía", IdPrograma = 1, Semestre = "I" };
            ctx.Asignaturas.Add(asig37);
            ctx.AsignaturasPorId[37] = asig37;

            var aulaSede5 = new AgAula { Id = 10, IdPrograma = 1, Semestre = "I", Grupo = "G1", IdSede = 5 };
            var aulaSede1 = new AgAula { Id = 20, IdPrograma = 1, Semestre = "I", Grupo = "G2", IdSede = 1 };
            ctx.Aulas.Add(aulaSede5);
            ctx.Aulas.Add(aulaSede1);
            ctx.AulasPorId[10] = aulaSede5;
            ctx.AulasPorId[20] = aulaSede1;

            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 10, IdAsignatura = 37 });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 20, IdAsignatura = 37 });

            // Docente habilitado para semestre I en Sede 1
            ctx.PoolPorAsignatura[37] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { Id = 1, IdTercero = 99, IdAsignatura = 37, IdPrograma = 1, Semestre = "I", Activo = 1 }
            };
            ctx.RegistrarDocenteHabilitado(99, 1, 37, "I");

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert: No debe emitir wildcard en 'asigs' porque una cohorte sí tiene docente
            asigs.Should().NotContain(37);
            // Pero sí debe estar en 'claves' si carece de cobertura (en este caso ambas comparten idProg/semestre, si ambas tienen hab, detalle es vacío)
            detalle.Should().BeEmpty();
        }

        [Fact]
        public void Consentimiento_WhenMultiSemestre_ShouldMatchFabricaMotivo()
        {
            // Arrange (Fase 2058): Materia con docente para semestre "II". Aula demanda semestre "II" y otra demanda "III"
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig = new Asignatura { Id = 42, Nombre = "Gestión", IdPrograma = 1, Semestre = "MultiSemestre" };
            ctx.Asignaturas.Add(asig);
            ctx.AsignaturasPorId[42] = asig;

            var aulaII = new AgAula { Id = 1, IdPrograma = 1, Semestre = "II", Grupo = "A", IdSede = 1 };
            var aulaIII = new AgAula { Id = 2, IdPrograma = 1, Semestre = "III", Grupo = "B", IdSede = 2 };
            ctx.Aulas.Add(aulaII);
            ctx.Aulas.Add(aulaIII);
            ctx.AulasPorId[1] = aulaII;
            ctx.AulasPorId[2] = aulaIII;

            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 42 });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 42 });

            ctx.PoolPorAsignatura[42] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { Id = 1, IdTercero = 88, IdAsignatura = 42, IdPrograma = 1, Semestre = "II", Activo = 1 }
            };
            ctx.RegistrarDocenteHabilitado(88, 1, 42, "II");

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert: Aula II tiene docente habilitado; Aula III no tiene docente habilitado
            asigs.Should().NotContain(42); // No wildcard porque Aula II sí tiene
            claves.Should().Contain("42:2"); // Sede 2 de Aula III
            detalle.Should().ContainSingle(d => d.IdAula == 2 && d.Semestre == "III");
            detalle[0].Motivo.Should().Be(AgFabricaIndividuos.ClasificarMotivo(ctx, asig, "III", 1).Codigo());
        }

        [Fact]
        public void Consentimiento_WhenAulaNoGenerable_ShouldNotConsent()
        {
            // Arrange (Fase 2063): Aula no generable (por ejemplo, estado Bloqueada o no permitida)
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig = new Asignatura { Id = 99, Nombre = "Electiva", IdPrograma = 1, Semestre = "I" };
            ctx.Asignaturas.Add(asig);
            ctx.AsignaturasPorId[99] = asig;

            // Aula inactiva / no disponible con manifiesto
            var aula = new AgAula { Id = 50, IdPrograma = 1, Semestre = "I", Grupo = "X", IdSede = 1, Estado = AgEstadoAula.Inactiva };
            ctx.Aulas.Add(aula);
            ctx.AulasPorId[50] = aula;
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 50, IdAsignatura = 99 });
            ctx.ManifiestoCobertura = new AgManifiestoCobertura(DateTime.Today, DateTime.Today.AddDays(90), new[]
            {
                new AgRelacionCobertura(new AgClaveRelacionCobertura(1, "SUB", 1, 1, 1, 50, 99, "I", "X", "Presencial"), 4, 4, (int)AgEstadoAula.Inactiva)
            });

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert: No debe consentir vacantes de aulas no generables
            asigs.Should().BeEmpty();
            claves.Should().BeEmpty();
            detalle.Should().BeEmpty();
        }

        [Fact]
        public void Consentimiento_WhenManifiestoNulo_ShouldIncludeAllAulas()
        {
            // Arrange (Fase 2067): Sin manifiesto (ej. diagnóstico inicial sin cobertura calculada), incluye todas las aulas demandadas
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;
            ctx.ManifiestoCobertura = null;

            var asig = new Asignatura { Id = 37, Nombre = "Estadística", IdPrograma = 1, Semestre = "I" };
            ctx.Asignaturas.Add(asig);
            ctx.AsignaturasPorId[37] = asig;

            var aula = new AgAula { Id = 10, IdPrograma = 1, Semestre = "I", Grupo = "G1", IdSede = 5, Estado = AgEstadoAula.Inactiva };
            ctx.Aulas.Add(aula);
            ctx.AulasPorId[10] = aula;
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 10, IdAsignatura = 37 });

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert: Con manifiesto nulo no se evalúa PuedeGenerarAula; se incluye el aula demandada
            asigs.Should().Contain(37);
            claves.Should().Contain("37:5");
            detalle.Should().ContainSingle(d => d.IdAula == 10 && d.IdAsignatura == 37);
        }

        [Fact]
        public void Consentimiento_WhenManifiestoConstruido_ShouldExcludeAulasNoGenerables()
        {
            // Arrange (Fase 2067): Con manifiesto construido, las aulas no generables son excluidas
            var ctx = new AgContexto();
            ctx.IdPrograma = 1;

            var asig1 = new Asignatura { Id = 37, Nombre = "Materia Generable", IdPrograma = 1, Semestre = "I" };
            var asig2 = new Asignatura { Id = 38, Nombre = "Materia No Generable", IdPrograma = 1, Semestre = "I" };
            ctx.Asignaturas.Add(asig1);
            ctx.Asignaturas.Add(asig2);
            ctx.AsignaturasPorId[37] = asig1;
            ctx.AsignaturasPorId[38] = asig2;

            var aulaGenerable = new AgAula { Id = 10, IdPrograma = 1, Semestre = "I", Grupo = "G1", IdSede = 1, Estado = AgEstadoAula.Disponible };
            var aulaNoGenerable = new AgAula { Id = 20, IdPrograma = 1, Semestre = "I", Grupo = "G2", IdSede = 2, Estado = AgEstadoAula.Disponible };
            ctx.Aulas.Add(aulaGenerable);
            ctx.Aulas.Add(aulaNoGenerable);
            ctx.AulasPorId[10] = aulaGenerable;
            ctx.AulasPorId[20] = aulaNoGenerable;

            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 10, IdAsignatura = 37 });
            ctx.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 20, IdAsignatura = 38 });

            // Manifiesto sólo incluye aula 10
            ctx.ManifiestoCobertura = new AgManifiestoCobertura(DateTime.Today, DateTime.Today.AddDays(90), new[]
            {
                new AgRelacionCobertura(new AgClaveRelacionCobertura(1, "SUB", 1, 1, 1, 10, 37, "I", "G1", "Presencial"), 4, 4, (int)AgEstadoAula.Disponible)
            });

            // Act
            var detalle = AgConsentimientoVacantes.Calcular(ctx, out var asigs, out var claves);

            // Assert: Solo aula 10 / asig 37 debe consentirse; aula 20 excluida por PuedeGenerarAula
            asigs.Should().Contain(37);
            asigs.Should().NotContain(38);
            claves.Should().Contain("37:1");
            claves.Should().NotContain("38:2");
            detalle.Should().ContainSingle(d => d.IdAula == 10 && d.IdAsignatura == 37);
        }

        [Fact]
        public void EsVigente_WhenMismoUniverso_ShouldBeTrue()
        {
            // Arrange (Fase 2071): Mismo universo de asignaturas y claves compuestas
            var ctx = new AgContexto
            {
                AsignaturasVacantesConsentidas = new HashSet<int> { 37, 53 },
                ClavesVacantesConsentidas = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "37:1", "53:2" }
            };

            var asigsActuales = new HashSet<int> { 53, 37 };
            var clavesActuales = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "53:2", "37:1" };

            // Act
            bool vigente = AgConsentimientoVacantes.EsVigente(ctx, asigsActuales, clavesActuales);

            // Assert
            vigente.Should().BeTrue();
        }

        [Fact]
        public void EsVigente_WhenFaltaUnaClave_ShouldBeFalse()
        {
            // Arrange (Fase 2071): Cambia una clave compuesta (por ejemplo, nueva sede o cohorte añadida)
            var ctx = new AgContexto
            {
                AsignaturasVacantesConsentidas = new HashSet<int> { 37 },
                ClavesVacantesConsentidas = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "37:1" }
            };

            var asigsActuales = new HashSet<int> { 37 };
            var clavesActuales = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "37:1", "37:2" };

            // Act
            bool vigente = AgConsentimientoVacantes.EsVigente(ctx, asigsActuales, clavesActuales);

            // Assert
            vigente.Should().BeFalse();
        }

        [Fact]
        public void EsVigente_WhenAsignaturasNulas_ShouldBeFalse()
        {
            // Arrange (Fase 2071): Contexto sin consentimiento previo registrado (null)
            var ctx = new AgContexto
            {
                AsignaturasVacantesConsentidas = null,
                ClavesVacantesConsentidas = null
            };

            var asigsActuales = new HashSet<int> { 37 };
            var clavesActuales = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "37:1" };

            // Act
            bool vigente = AgConsentimientoVacantes.EsVigente(ctx, asigsActuales, clavesActuales);

            // Assert
            vigente.Should().BeFalse();
        }
    }
}
