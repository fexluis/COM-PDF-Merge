using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgValidadorReglasTests
    {
        private readonly AgValidadorReglas _sut = new AgValidadorReglas();

        private static AgBloqueHorario CrearBloque(int id, int diaSemana = 2, string horaInicio = "08:00", string horaFin = "12:00", string franja = "MANANA")
        {
            return new AgBloqueHorario
            {
                Id = id,
                DiaSemanaNumero = diaSemana,
                HoraInicio = horaInicio,
                HoraFin = horaFin,
                DuracionHoras = 4.0,
                Franja = franja,
                Activo = 1
            };
        }

        [Fact]
        public void Evaluar_WhenEmptyGenesAndRequiredSubjectsExist_ShouldRecordHardViolations()
        {
            // Arrange
            var individuo = new AgIndividuo();
            var contexto = new AgContexto
            {
                Asignaturas = new List<Asignatura>
                {
                    new Asignatura { Id = 1, Nombre = "Cálculo I" },
                    new Asignatura { Id = 2, Nombre = "Física I" }
                }
            };

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.HardViolations.Should().Be(2);
            individuo.Estado.Should().Be(AgEstadoIndividuo.Infactible);
            individuo.EsFactible.Should().BeFalse();
        }

        [Fact]
        public void Evaluar_WhenDocenteHasSimultaneousClassesR1_ShouldRecordR1HardViolation()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.AulasPorId[2] = new AgAula { Id = 2, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.RegistrarDocenteHabilitado(10, 1, 2, "I");
            contexto.BloquesPorId[5] = CrearBloque(5, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1,
                IdBloqueHorario = 5,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 10,
                IdAsignatura = 2,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 2,
                IdBloqueHorario = 5, // Mismo bloque horario y fechas superpuestas
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R1_SuperposicionDocente);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenAulaHasSimultaneousClassesR2_ShouldRecordR2HardViolation()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.RegistrarDocenteHabilitado(20, 1, 2, "I");
            contexto.BloquesPorId[3] = CrearBloque(3, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1,
                IdBloqueHorario = 3,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 2,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1, // Misma aula
                IdBloqueHorario = 3, // Mismo bloque
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenDatesOutsideSemesterWindowW1_ShouldRecordW1HardViolation()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 1, 15), // Antes del semestre
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.W1_VentanaSemestre);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenGenIsVacante_ShouldNotRecordR5HardViolation_AndApplySoftPenalty()
        {
            // Arrange (Fase 320: R5 relajada para vacantes)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 2);
            contexto.Pesos.PesoVacante = 50.0;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 0,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_CANDIDATOS.Codigo(),
                IdAsignatura = 10,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R5_HabilitacionDocentePool && v.EsDura);
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R5_HabilitacionDocentePool && !v.EsDura);
            individuo.HardViolations.Should().Be(0);
            individuo.Vacantes.Should().Be(1);
            individuo.SoftViolations.Should().Be(50.0);
        }

        [Fact]
        public void Evaluar_WhenGenVacanteSinCentinela_ShouldRecordR9HardViolation()
        {
            // Arrange (Fase 464 / Fase 470)
            var contexto = new AgContexto();
            contexto.IdTerceroVacante = 192;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 25, // No es el centinela 192
                EsVacante = 1,
                IdAsignatura = 10,
                IdPrograma = 1,
                Semestre = "I",
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R9_VacanteSinCentinela && v.EsDura);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenDiasDuplicadosEnGrupo_ShouldRecordR8HardViolation()
        {
            // Arrange (Fase 456 / Fase 470)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Grupo = "A" };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 5, horaInicio: "08:00", horaFin: "12:00", franja: "MANANA");
            contexto.BloquesPorId[2] = CrearBloque(2, diaSemana: 5, horaInicio: "08:00", horaFin: "12:00", franja: "MANANA");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                ClaveGrupo = "1|1|A",
                DiaSemana = 5, // Viernes
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                ClaveGrupo = "1|1|A",
                DiaSemana = 5, // Viernes repetido
                IdAula = 1,
                IdBloqueHorario = 2,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R8_DiasDuplicadosGrupo && v.EsDura);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenAulaFueraDeCohorte_ShouldRecordR4HardViolation()
        {
            // Arrange (Fase 487)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Codigo = "AULA-II-A", Semestre = "II", Grupo = "A", Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R4_CompatibilidadAula && v.EsDura && v.Mensaje.Contains("fuera de cohorte"));
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Evaluar_WhenAulaSinSemestreOGrupo_ShouldTreatAsWildcardAndNotTriggerR4()
        {
            // Arrange (Fase 489)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Codigo = "AULA-GRAL", Semestre = null, Grupo = null, Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.BloquesPorId[1] = CrearBloque(1, diaSemana: 2);

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1,
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.HardViolations.Should().Be(0);
        }

        [Fact]
        public void Evaluar_WhenGenHasHorasButIdBloqueZero_ShouldRecordR8HardViolation()
        {
            // Arrange: Caso H6 (gen con horas pero IdBloqueHorario <= 0)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                HorasGen = 4.0,
                IdBloqueHorario = 0,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R14_BloqueHorarioNoAsignado && v.EsDura && v.Mensaje.Contains("bloque horario no asignado"));
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Validador_DosAulasMismaCohorteMismoCarril_SinR13()
        {
            // Arrange (Fase 1075): Dos aulas distintas (p.ej. Presencial Aula 1 y Plataforma Aula 2) de la misma cohorte
            // programadas en las mismas fechas y mismo carril no deben considerarse oleadas solapadas entre sí.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.RegistrarDocenteHabilitado(20, 2, 1, "I");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1, // Aula 1
                DiasPatron = "2,3",
                IdBloqueHorario = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 2,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 2, // Aula 2
                DiasPatron = "2,3",
                IdBloqueHorario = 2,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R13_OleadasDispersas);
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R6_SecuenciaAsignaturas);
        }

        [Fact]
        public void Validador_Gen4hConBloque8h_EmiteR11()
        {
            // Arrange (Fase 1076): gen con HorasSemanales=4 pero IdBloqueHorario de 8h debe ser detectado por R11
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.AulasPorId[1] = new AgAula { Id = 1, HorasDia = 4 };

            var bloque8h = new AgBloqueHorario { Id = 99, DiaSemanaNumero = 2, DuracionHoras = 8.0, HoraInicio = "08:00", HoraFin = "17:00", Franja = "MANANA", Activo = 1 };
            contexto.BloquesPorId[99] = bloque8h;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1,
                IdBloqueHorario = 99,
                HorasSemanales = 4.0, // Forzado a 4, pero el bloque real es de 8h
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R11_DuracionSesionInvalida && v.EsDura && v.Mensaje.Contains("8h"));
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Fabrica_PlanInfactible_EmiteR7DuraYAislaAsignatura()
        {
            // Arrange (Fase 1077 / Fase 1085): Gen vacante por SIN_PLAN_OLEADA (plan oleadas infactible) genera R7 dura
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 101,
                IdAula = 1,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R7_CoberturaHorariaIncompleta && v.EsDura && v.Mensaje.Contains("I5"));
        }

        [Fact]
        public void Validador_VacanteSinPlanOleada_EmiteR7()
        {
            // Arrange (Fase 1085): Gen vacante por SIN_PLAN_OLEADA emite R7 dura
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 101,
                IdAula = 1,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R7_CoberturaHorariaIncompleta && v.EsDura);
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Validador_VacanteSinBloqueNivel_NoEmiteR7()
        {
            // Arrange (Fase 1085): Gen vacante por SIN_DISPONIBILIDAD (sin bloque de 4h) NO genera R7 dura (evita falso positivo de plan infactible)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 101,
                IdAula = 1,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R7_CoberturaHorariaIncompleta);
        }

        [Fact]
        public void Validador_CohorteConDosAulasDistintaVentana_UsaMaxCarriles()
        {
            // Arrange (Fase 1087): Cohorte con Aula 1 (LUN-JUE, 2 carriles) y Aula 2 (VIE-SAB, 1 carril).
            // maxCarriles de la cohorte debe ser max(2, 1) = 2.
            // Dos asignaturas concurrentes en la misma semana ISO no deben disparar R5.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.RegistrarDocenteHabilitado(20, 2, 1, "I");

            contexto.AulasPorId[1] = new AgAula { Id = 1, DiasVentana = "LUN-JUE" }; // 2 carriles
            contexto.AulasPorId[2] = new AgAula { Id = 2, DiasVentana = "VIE-SAB" }; // 1 carril

            var individuo = new AgIndividuo();
            // Asignatura 1 en Aula 2 (si tomara solo Aula 2 daría maxCarriles=1 y fallaría por tener 2 materias)
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 2,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });
            // Asignatura 2 en Aula 1
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 2,
                IdPrograma = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R5_MaxAsignaturasSemana);
        }

        [Fact]
        public void Evaluar_OleadaBordeSemanaISO_NoEmiteR5Max()
        {
            // Asignatura 1 termina domingo 15/03/2026, Asignatura 2 empieza lunes 16/03/2026.
            // Aunque compartan o linden en la misma semana ISO en los extremos, por días reales nunca son concurrentes.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, DiasVentana = "VIE-SAB" }; // 1 carril
            contexto.RegistrarDocenteHabilitado(10, 1, 148, "IX");
            contexto.RegistrarDocenteHabilitado(20, 1, 150, "IX");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 148,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 3, 15)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 150,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 3, 16),
                FechaFinEstimada = new DateTime(2026, 5, 3)
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R5_MaxAsignaturasSemana);
        }

        [Fact]
        public void Evaluar_TresAsignaturasConcurrentesMismoDia_EmiteR5Max()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, DiasVentana = "VIE-SAB" }; // 1 carril
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "IX");
            contexto.RegistrarDocenteHabilitado(20, 1, 2, "IX");
            contexto.RegistrarDocenteHabilitado(30, 1, 3, "IX");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 4, 30)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 2,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 4, 30)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 3,
                IdTercero = 30,
                IdAsignatura = 3,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "IX",
                Grupo = "A",
                IdAula = 1,
                FechaInicioEstimada = new DateTime(2026, 2, 2),
                FechaFinEstimada = new DateTime(2026, 4, 30)
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R5_MaxAsignaturasSemana);
        }

        [Fact]
        public void Evaluar_AulaAprobada_NoEmiteR4PorEstado()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Aprobado, Semestre = "I", Grupo = "A", IdPrograma = 1 };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R4_CompatibilidadAula && v.Mensaje.Contains("no está disponible"));
        }

        [Fact]
        public void Evaluar_AsignaturasDeAulasNoCubiertas_NoEmiteR5Requerida()
        {
            var contexto = new AgContexto
            {
                Asignaturas = new List<Asignatura>
                {
                    new Asignatura { Id = 1, Nombre = "Cálculo X" },
                    new Asignatura { Id = 99, Nombre = "Cálculo I (Otra cohorte no cubierta)" }
                },
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 10, IdAsignatura = 1 },
                    new AgAulaAsignatura { IdAula = 20, IdAsignatura = 99 } // Aula 20 no está presente en los genes del cromosoma
                }
            };
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[10] = new AgAula { Id = 10, Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "X");

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 10
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R5_HabilitacionDocentePool
                && v.Mensaje.Contains("Cálculo I (Otra cohorte no cubierta)"));
        }

        [Fact]
        public void Validador_GenSinBloqueNivel_NoEmiteR14Generico()
        {
            // Arrange (Fase 1100): Gen con SIN_BLOQUE_NIVEL e IdBloqueHorario=0 no debe duplicar sanción con R14 genérico
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 1,
                HorasGen = 4.0,
                IdBloqueHorario = 0,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert: no debe emitir R14_BloqueHorarioNoAsignado
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R14_BloqueHorarioNoAsignado);
        }

        [Fact]
        public void Validador_GenSinBloqueNivel_EmiteR7Dura()
        {
            // Arrange (Fase 1103 / Fase 1107): Gen con SIN_BLOQUE_NIVEL debe emitir regla dura R7_CoberturaHorariaIncompleta
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 1,
                HorasGen = 4.0,
                IdBloqueHorario = 0,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert: debe emitir R7_CoberturaHorariaIncompleta como violación dura
            individuo.Violaciones.Should().Contain(v =>
                v.Regla == AgReglaViolacion.R7_CoberturaHorariaIncompleta &&
                v.EsDura &&
                v.Mensaje.Contains("sin bloque de duración del nivel"));
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Validador_GenSinPlanOleada_NoEmiteR14()
        {
            // Arrange (Fase 1106 / Fase 1107): Gen con SIN_PLAN_OLEADA e IdBloqueHorario=0 no debe emitir R14 (ya cubierto por R7)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 10,
                IdAula = 1,
                HorasGen = 4.0,
                IdBloqueHorario = 0,
                EsVacante = 1,
                MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                FechaInicioEstimada = new DateTime(2026, 2, 10),
                FechaFinEstimada = new DateTime(2026, 3, 10)
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert: no debe emitir R14_BloqueHorarioNoAsignado
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R14_BloqueHorarioNoAsignado);
            // Pero sí debe emitir R7
            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R7_CoberturaHorariaIncompleta && v.EsDura);
        }

        [Fact]
        public void Validador_GenDeFranjaFueraDeJornada_EmiteR11()
        {
            // Arrange (Fase 1133 / Fase 1134): Aula con Jornada = "TARDE" (franja esperada "TARDE").
            // Gen asignado con bloque de 4h pero con Franja = "MANANA".
            // R11 debe detectar la discrepancia de franja y emitir violación dura.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, HorasDia = 4, Jornada = "TARDE" };
            contexto.BloquesPorId[1] = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Franja = "MANANA" };

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdPrograma = 1,
                IdAula = 10,
                HorasGen = 4.0,
                HorasSemanales = 4.0,
                IdBloqueHorario = 1,
                DiaSemana = 2,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert: debe emitir R11 como violación dura
            individuo.Violaciones.Should().Contain(v =>
                v.Regla == AgReglaViolacion.R11_DuracionSesionInvalida &&
                v.EsDura &&
                v.Mensaje.Contains("jornada TARDE"));
            individuo.HardViolations.Should().BeGreaterThan(0);
        }

        [Fact]
        public void ValidadorReglas_Sabado4hDosFranjas_CeroViolacionesDuras()
        {
            // Arrange (Fase 1292.d): Grupo con 2 genes en sábado 4h (MANANA 08:00-12:00 y TARDE 13:00-17:00).
            // No debe violar R4 (tope 2 franjas en sábado) ni R14 (solapamiento de bloques).
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, Jornada = "DIURNA", HorasDia = 4 };
            contexto.RegistrarDocenteHabilitado(10, 1, 100, "");   // Fase 1303: comodín de semestre para habilitar docente 10 en R5

            var bManana = new AgBloqueHorario { Id = 1, DiaSemanaNumero = 6, DuracionHoras = 4.0, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA", Activo = 1 };
            var bTarde = new AgBloqueHorario { Id = 2, DiaSemanaNumero = 6, DuracionHoras = 4.0, HoraInicio = "13:00", HoraFin = "17:00", Franja = "TARDE", Activo = 1 };
            contexto.BloquesPorId[1] = bManana;
            contexto.BloquesPorId[2] = bTarde;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 100,
                IdPrograma = 1,
                IdAula = 10,
                ClaveGrupo = "100|10|A",
                HorasGen = 24.0,
                HorasSemanales = 4.0,
                IdBloqueHorario = 1,
                DiaSemana = 6,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 10,
                IdAsignatura = 100,
                IdPrograma = 1,
                IdAula = 10,
                ClaveGrupo = "100|10|A",
                HorasGen = 24.0,
                HorasSemanales = 4.0,
                IdBloqueHorario = 2,
                DiaSemana = 6,
                EsVacante = 0,
                Estado = AgEstadoGen.Valido
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert: R4 y R14 no deben emitir violaciones duras y el individuo no debe tener HardViolations (Fase 1297 / Fase 1299)
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R4_DosAsignaturasMismoDia && v.EsDura);
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R14_BloqueSolapadoPorIntervalo && v.EsDura);
            individuo.HardViolations.Should().Be(0);
        }

        [Fact]
        public void Evaluar_SinGenes_RegistraDetalleDeViolaciones()
        {
            // Fase 1525: antes sólo se guardaba el contador y ag_violacion quedaba vacío
            var contexto = new AgContexto
            {
                Asignaturas = new List<Asignatura>
                {
                    new Asignatura { Id = 1, Nombre = "MATEMATICAS" },
                    new Asignatura { Id = 2, Nombre = "ESTADISTICA" }
                }
            };

            var individuo = new AgIndividuo();

            _sut.Evaluar(individuo, contexto);

            individuo.HardViolations.Should().Be(2);
            individuo.Violaciones.Should().HaveCount(2, "el detalle debe quedar disponible para persistirse");
            individuo.Violaciones.Should().OnlyContain(v => v.EsDura);
        }

        [Fact]
        public void Evaluar_OleadasConsecutivasMismaAula_NoEmiteR2NiR3NiR14()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[12] = new AgAula { Id = 12, Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(10, 1, 154, "X");
            contexto.RegistrarDocenteHabilitado(20, 1, 156, "X");
            contexto.BloquesPorId[29] = new AgBloqueHorario { Id = 29, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", Franja = "M" };

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 5)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 4, 6),
                FechaFinEstimada = new DateTime(2027, 4, 19)
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R3_SuperposicionGrupo);
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R14_BloqueSolapadoPorIntervalo);
        }

        [Fact]
        public void Evaluar_OleadasConsecutivasMismoDocente_NoEmiteR1()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[1] = new AgAula { Id = 1, Estado = AgEstadoAula.Disponible };
            contexto.AulasPorId[2] = new AgAula { Id = 2, Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(10, 1, 1, "I");
            contexto.RegistrarDocenteHabilitado(10, 1, 2, "I");
            contexto.BloquesPorId[5] = new AgBloqueHorario { Id = 5, DiaSemanaNumero = 1, HoraInicio = "08:00", HoraFin = "12:00", Franja = "M" };

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 1,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 1,
                IdBloqueHorario = 5,
                DiaSemana = 1,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 3, 15)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 10,
                IdAsignatura = 2,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I",
                Grupo = "B",
                IdAula = 2,
                IdBloqueHorario = 5,
                DiaSemana = 1,
                FechaInicioEstimada = new DateTime(2027, 3, 16),
                FechaFinEstimada = new DateTime(2027, 4, 30)
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R1_SuperposicionDocente);
        }

        [Fact]
        public void Evaluar_VentanasSolapadasMismaAula_EmiteR2()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[12] = new AgAula { Id = 12, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 154, "X");
            contexto.RegistrarDocenteHabilitado(20, 1, 156, "X");
            contexto.BloquesPorId[29] = new AgBloqueHorario { Id = 29, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", Franja = "M", Activo = 1 };

            var fechaCruce = new DateTime(2027, 4, 9);
            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 10),
                FechasGen = new List<DateTime> { fechaCruce }
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "B",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 4, 5), // Solape 05/04 a 10/04
                FechaFinEstimada = new DateTime(2027, 4, 19),
                FechasGen = new List<DateTime> { fechaCruce }
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
        }

        [Fact]
        public void Evaluar_MismoBloqueSinFechas_EmiteR2Conservador()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[12] = new AgAula { Id = 12, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 154, "X");
            contexto.RegistrarDocenteHabilitado(20, 1, 156, "X");
            contexto.BloquesPorId[29] = new AgBloqueHorario { Id = 29, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", Franja = "M", Activo = 1 };

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = null,
                FechaFinEstimada = null
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "B",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = null,
                FechaFinEstimada = null
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
        }

        [Fact]
        public void Evaluar_MismoBloqueYFechas_EmiteR2SinDuplicarR14()
        {
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 1), new DateTime(2027, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[12] = new AgAula { Id = 12, Estado = AgEstadoAula.Disponible, Grupo = null };
            contexto.RegistrarDocenteHabilitado(10, 1, 154, "X");
            contexto.RegistrarDocenteHabilitado(20, 1, 156, "X");
            contexto.BloquesPorId[29] = new AgBloqueHorario { Id = 29, DiaSemanaNumero = 5, HoraInicio = "08:00", HoraFin = "12:00", Franja = "M", Activo = 1 };

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 10,
                IdAsignatura = 154,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "A",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 10)
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 20,
                IdAsignatura = 156,
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "X",
                Grupo = "B",
                IdAula = 12,
                IdBloqueHorario = 29,
                DiaSemana = 5,
                FechaInicioEstimada = new DateTime(2027, 2, 2),
                FechaFinEstimada = new DateTime(2027, 4, 10)
            });

            _sut.Evaluar(individuo, contexto);

            individuo.Violaciones.Should().Contain(v => v.Regla == AgReglaViolacion.R2_SuperposicionAula);
            individuo.Violaciones.Should().NotContain(v => v.Regla == AgReglaViolacion.R14_BloqueSolapadoPorIntervalo,
                "R14 no debe duplicar la colisión si tienen exactamente el mismo bloque horario");
        }

        [Fact]
        public void Evaluar_MateriasIntercaladasSinCruceReal_EstadisticaVsTeorias_CeroHard()
        {
            // Arrange (Fase 1809 / Libro 8 Corrida 94):
            // Estadística vs Teorías con semanas alternas/intercaladas en el mismo bloque y aula/cohorte.
            // Aunque sus ventanas [FechaInicio..FechaFin] se solapen temporalmente a nivel macro,
            // no comparten ninguna fecha de sesión real en AgCalendarioConflictos -> 0 HARD.
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.AulasPorId[10] = new AgAula { Id = 10, IdPrograma = 1, Semestre = "I", Grupo = "A", Estado = AgEstadoAula.Disponible };
            contexto.RegistrarDocenteHabilitado(101, 1, 1, "I"); // Docente Estadística
            contexto.RegistrarDocenteHabilitado(102, 1, 2, "I"); // Docente Teorías

            var bloqueSabado = new AgBloqueHorario
            {
                Id = 1,
                DiaSemanaNumero = 6, // Sábado
                HoraInicio = "08:00",
                HoraFin = "12:00",
                DuracionHoras = 4.0,
                Franja = "MANANA",
                Activo = 1
            };
            contexto.BloquesPorId[1] = bloqueSabado;

            var individuo = new AgIndividuo();
            individuo.Genes.Add(new AgGen
            {
                Id = 1,
                IdTercero = 101,
                IdAsignatura = 1, // Estadística
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 10,
                DiasPatron = "CARRIL_1",
                IdBloqueHorario = 1,
                DiaSemana = 6,
                HorasGen = 8.0,
                HorasSemanales = 4.0,
                FechaInicioEstimada = new DateTime(2026, 2, 7),
                FechaFinEstimada = new DateTime(2026, 2, 21),
                FechasGen = new List<DateTime>
                {
                    new DateTime(2026, 2, 7),  // Semana 1 (Sábado)
                    new DateTime(2026, 2, 21)  // Semana 3 (Sábado)
                }
            });
            individuo.Genes.Add(new AgGen
            {
                Id = 2,
                IdTercero = 102,
                IdAsignatura = 2, // Teorías
                IdPrograma = 1,
                IdSede = 1,
                Semestre = "I",
                Grupo = "A",
                IdAula = 10,
                DiasPatron = "CARRIL_2",
                IdBloqueHorario = 1,
                DiaSemana = 6,
                HorasGen = 8.0,
                HorasSemanales = 4.0,
                FechaInicioEstimada = new DateTime(2026, 2, 14),
                FechaFinEstimada = new DateTime(2026, 2, 28),
                FechasGen = new List<DateTime>
                {
                    new DateTime(2026, 2, 14), // Semana 2 (Sábado)
                    new DateTime(2026, 2, 28)  // Semana 4 (Sábado)
                }
            });

            // Act
            _sut.Evaluar(individuo, contexto);

            // Assert
            individuo.HardViolations.Should().Be(0);
            individuo.Violaciones.Should().NotContain(v => v.EsDura);
            individuo.Estado.Should().Be(AgEstadoIndividuo.Factible);
        }
    }
}
