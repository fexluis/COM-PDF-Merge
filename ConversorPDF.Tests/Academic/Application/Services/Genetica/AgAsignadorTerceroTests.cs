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
    public class AgAsignadorTerceroTests
    {
        private readonly AgAsignadorTercero _sut = new AgAsignadorTercero(new Random(42));

        [Fact]
        public void AsignarDocente_WhenNoCandidatesInPool_ShouldMarkAsVacanteSinCandidatos()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 1 };
            var contexto = new AgContexto();

            // Act
            bool asignado = _sut.AsignarDocente(gen, contexto);

            // Assert
            asignado.Should().BeFalse();
            gen.EsVacante.Should().Be(1);
            gen.IdTercero.Should().Be(0);
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_CANDIDATOS.Codigo());
        }

        [Fact]
        public void AsignarDocente_WhenCandidateAvailableInPool_ShouldAssignTeacher()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 5 };
            var contexto = new AgContexto();
            contexto.PoolPorAsignatura[5] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 25, Activo = 1, Prioridad = 1 }
            };

            // Act
            bool asignado = _sut.AsignarDocente(gen, contexto);

            // Assert
            asignado.Should().BeTrue();
            gen.IdTercero.Should().Be(25);
            gen.EsVacante.Should().Be(0);
            gen.MotivoVacante.Should().BeEmpty();
        }

        [Fact]
        public void AsignarDocente_WhenPoolExistsButAllExcludedByConflict_ShouldDiagnoseSinDisponibilidad()
        {
            // Arrange - Materia con pool habilitado, pero todos sus docentes excluidos por conflicto (Fase 1896-A RED)
            var gen = new AgGen { IdAsignatura = 7, IdPrograma = 1, Semestre = "I" };
            var ctx = new AgContexto();
            ctx.PoolPorAsignatura[7] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 11, Activo = 1 },
                new TerceroAsignatura { IdTercero = 12, Activo = 1 }
            };
            ctx.RegistrarDocenteHabilitado(11, 1, 7, "I");
            ctx.RegistrarDocenteHabilitado(12, 1, 7, "I");
            var excluidos = new HashSet<int> { 11, 12 };

            // Act
            bool ok = _sut.AsignarDocente(gen, ctx, excluidos);

            // Assert
            Assert.False(ok);
            Assert.Equal(1, gen.EsVacante);
            Assert.Equal(AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo(), gen.MotivoVacante);
        }

        [Fact]
        public void DiagnosticarMotivo_WhenNoPoolExists_ShouldReturnSinCandidatos()
        {
            // Arrange
            var gen = new AgGen { IdAsignatura = 99 };
            var contexto = new AgContexto();

            // Act
            var motivo = _sut.DiagnosticarMotivo(gen, contexto);

            // Assert
            motivo.Should().Be(AgMotivoVacante.SIN_CANDIDATOS.Codigo());
        }

        [Fact]
        public void DocenteDisponible_WhenDocenteOccupiedExternallyInProjectedDateOfGroup_ShouldReturnFalse()
        {
            // Arrange
            var contexto = CrearContextoBase();
            int docente = 101;
            contexto.RegistrarDocenteHabilitado(docente, 20, 10, "1");

            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = docente,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 9),
                FechaFinEstimada = new DateTime(2026, 2, 9),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) },
                IdTercero = docente,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            var individuo = new AgIndividuo
            {
                Genes = new List<AgGen> { gen1, gen2 }
            };

            // Sin ocupación externa: disponible
            _sut.DocenteDisponible(gen1, individuo, contexto).Should().BeTrue();

            // Con ocupación externa en la fecha proyectada del segundo gen del grupo (2026-02-09):
            contexto.CargarOcupacionExterna(new[]
            {
                (docente, new DateTime(2026, 2, 9), new FranjaHoraria(new TimeSpan(8, 0, 0), new TimeSpan(12, 0, 0)))
            });

            // Act
            bool disponible = _sut.DocenteDisponible(gen1, individuo, contexto);

            // Assert
            disponible.Should().BeFalse();
        }

        [Fact]
        public void BuscarAlterno_WhenAlternateCandidateAvailable_ShouldUpdateAllNonFixedGroupMembersCoherently()
        {
            // Arrange
            var contexto = CrearContextoBase();
            int docenteInicial = 101;
            int alterno = 202;

            contexto.RegistrarDocenteHabilitado(alterno, 20, 10, "1");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = alterno, Activo = 1, Prioridad = 1 }
            };

            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = docenteInicial,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 9),
                FechaFinEstimada = new DateTime(2026, 2, 9),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) },
                IdTercero = docenteInicial,
                EsVacante = 0,
                Estado = AgEstadoGen.Conflicto
            };

            var genFijado = new AgGen
            {
                Id = 3,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G_FIJADO",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 16),
                FechaFinEstimada = new DateTime(2026, 2, 16),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 16) },
                IdTercero = 777,
                EsVacante = 0,
                Estado = AgEstadoGen.Fijado
            };

            var genOtroGrupo = new AgGen
            {
                Id = 4,
                IdAula = 2,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G2",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = 888,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            var individuo = new AgIndividuo
            {
                Genes = new List<AgGen> { gen1, gen2, genFijado, genOtroGrupo }
            };

            // Act
            int resultado = _sut.BuscarAlterno(gen1, contexto, individuo);

            // Assert
            resultado.Should().Be(alterno);
            gen1.IdTercero.Should().Be(alterno);
            gen1.EsVacante.Should().Be(0);
            gen1.MotivoVacante.Should().BeEmpty();

            gen2.IdTercero.Should().Be(alterno);
            gen2.EsVacante.Should().Be(0);
            gen2.MotivoVacante.Should().BeEmpty();

            // El gen fijado no se debe alterar
            genFijado.IdTercero.Should().Be(777);
            genFijado.Estado.Should().Be(AgEstadoGen.Fijado);

            // El gen de otro grupo no se debe alterar
            genOtroGrupo.IdTercero.Should().Be(888);
        }

        [Fact]
        public void BuscarAlterno_WhenGroupHasFixedMemberWithDifferentTeacher_ShouldNotAssignDifferentCandidate()
        {
            // Arrange
            var contexto = CrearContextoBase();
            int docenteInicial = 101;
            int candAlterno = 202;
            int docenteFijado = 777;

            contexto.RegistrarDocenteHabilitado(candAlterno, 20, 10, "1");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = candAlterno, Activo = 1, Prioridad = 1 }
            };

            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = docenteInicial,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            // genFijado pertenece al MISMO grupo G1 pero está fijado a docenteFijado (777)
            var genFijado = new AgGen
            {
                Id = 2,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 9),
                FechaFinEstimada = new DateTime(2026, 2, 9),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) },
                IdTercero = docenteFijado,
                EsVacante = 0,
                Estado = AgEstadoGen.Fijado
            };

            var individuo = new AgIndividuo
            {
                Genes = new List<AgGen> { gen1, genFijado }
            };

            // Act: intentamos buscar alterno para gen1; candAlterno (202) es incompatible con genFijado (777)
            int resultado = _sut.BuscarAlterno(gen1, contexto, individuo);

            // Assert
            resultado.Should().Be(0);
            gen1.IdTercero.Should().Be(docenteInicial);
            genFijado.IdTercero.Should().Be(docenteFijado);
        }

        [Fact]
        public void BuscarAlterno_WhenCandidateCollidesOnPosteriorOrIntermediateDate_ShouldNotAssign()
        {
            // Arrange
            var contexto = CrearContextoBase();
            int docenteInicial = 101;
            int candAlterno = 202;

            contexto.RegistrarDocenteHabilitado(candAlterno, 20, 10, "1");
            contexto.RegistrarDocenteHabilitado(candAlterno, 20, 99, "1");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = candAlterno, Activo = 1, Prioridad = 1 }
            };

            // gen1 del grupo G1 está en fecha 2026-02-02
            var gen1 = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = docenteInicial,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            // gen2 del grupo G1 está en fecha posterior 2026-02-09
            var gen2 = new AgGen
            {
                Id = 2,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 9),
                FechaFinEstimada = new DateTime(2026, 2, 9),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) },
                IdTercero = docenteInicial,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            // genOtroGrupo es de OTRA cohorte (Grupo "G_OTRO", Asignatura 99), pero ya tiene asignado a candAlterno
            // exactamente en la fecha posterior 2026-02-09 en el mismo bloque horario 1 (08:00 - 12:00)
            var genOtroGrupo = new AgGen
            {
                Id = 3,
                IdAula = 2,
                IdAsignatura = 99,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G_OTRO",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 9),
                FechaFinEstimada = new DateTime(2026, 2, 9),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 9) },
                IdTercero = candAlterno,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            };

            var individuo = new AgIndividuo
            {
                Genes = new List<AgGen> { gen1, gen2, genOtroGrupo }
            };

            // Act
            int alterno = _sut.BuscarAlterno(gen1, contexto, individuo);

            // Assert: No debe asignar a candAlterno porque colisiona en la fecha 2026-02-09 con genOtroGrupo
            alterno.Should().Be(0);
            gen1.IdTercero.Should().Be(docenteInicial);
            gen2.IdTercero.Should().Be(docenteInicial);
        }

        [Fact]
        public void BuscarYAsignarAlternoGrupo_ShouldDelegateToBuscarAlternoAndReturnCandidate()
        {
            // Arrange
            var contexto = CrearContextoBase();
            int alterno = 303;

            contexto.RegistrarDocenteHabilitado(alterno, 20, 10, "1");
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = alterno, Activo = 1, Prioridad = 1 }
            };

            var gen = new AgGen
            {
                Id = 1,
                IdAula = 1,
                IdAsignatura = 10,
                IdPrograma = 20,
                IdSede = 30,
                Grupo = "G1",
                Semestre = "1",
                IdBloqueHorario = 1,
                HorasGen = 4,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 2, 2),
                FechasGen = new List<DateTime> { new DateTime(2026, 2, 2) },
                IdTercero = 0,
                EsVacante = 1,
                Estado = AgEstadoGen.Valido
            };

            var individuo = new AgIndividuo
            {
                Genes = new List<AgGen> { gen }
            };

            // Act
            int resultado = _sut.BuscarYAsignarAlternoGrupo(gen, contexto, individuo);

            // Assert
            resultado.Should().Be(alterno);
            gen.IdTercero.Should().Be(alterno);
            gen.EsVacante.Should().Be(0);
        }

        private static AgContexto CrearContextoBase()
        {
            var contexto = new AgContexto
            {
                IdCorrida = 1,
                IdPrograma = 20,
                IdTerceroVacante = 999999
            };
            contexto.Ventana = VentanaSemestre.Crear(
                1,
                "2026-1",
                new DateTime(2026, 2, 1),
                new DateTime(2026, 6, 30),
                AgOrigenVentana.Presencial).Value;

            var bloque = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 1, // Lunes
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4,
                Franja = "MANANA",
                Activo = 1
            };
            contexto.Bloques.Add(bloque);
            contexto.AulasPorId[1] = new AgAula { Id = 1, ModalidadVinculacion = "Presencial" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, ModalidadVinculacion = "Presencial" };
            contexto.ConstruirIndices();
            return contexto;
        }
    }
}
