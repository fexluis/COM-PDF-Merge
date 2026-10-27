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
    public class AgFabricaIndividuosTests
    {
        private readonly AgFabricaIndividuos _sut = new AgFabricaIndividuos(new Random(42));

        [Fact]
        public void CrearAleatorio_WhenContextHasSubjects_ShouldGenerateOneGenPerSubject()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Subunidad = "FACULTAD_ING";
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Algoritmos", Horas = 48, HorasSemanales = 4 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "Bases de Datos", Horas = 64, HorasSemanales = 4 });

            contexto.Aulas.Add(new AgAula { Id = 10, Estado = AgEstadoAula.Disponible, DiasVentana = "MAR-VIE" });
            var bloque = new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 };
            contexto.Bloques.Add(bloque);
            contexto.Bloques.Add(new AgBloqueHorario { Id = 6, DiaSemanaNumero = 3, DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 7, DiaSemanaNumero = 4, DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 8, DiaSemanaNumero = 5, DuracionHoras = 4, Activo = 1 });

            contexto.Docentes.Add(new Tercero { Id = 100, Nombre = "Docente Turing", Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().HaveCount(4); // v4 carriles disjuntos: 2 ranuras/carril x 2 asignaturas = 4 genes
            individuo.Genes.Select(g => g.IdAsignatura).Distinct().Should().Contain(new[] { 1, 2 });
            individuo.Genes[0].IdAsignatura.Should().Be(1);
            individuo.Genes[0].IdBloqueHorario.Should().Be(5);
            individuo.Genes[0].DiaSemana.Should().Be(2);
            individuo.Genes[0].HorasSemanales.Should().Be(4);
        }

        [Fact]
        public void CrearAleatorio_WhenIdAulaCorridaIsSet_ShouldAssignSpecificAulaToAllGenes()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.IdAulaCorrida = 99;
            contexto.Subunidad = "FACULTAD_ING";
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Algoritmos", Horas = 48, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 10, Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 99, Estado = AgEstadoAula.Disponible });
            var bloque = new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 };
            contexto.Bloques.Add(bloque);
            contexto.Docentes.Add(new Tercero { Id = 100, Nombre = "Docente Turing", Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().HaveCount(2);
            individuo.Genes.Should().OnlyContain(g => g.IdAula == 99);
        }

        [Fact]
        public void CrearAleatorio_WhenPatronMultiDia_ShouldPreserveInvarianteI1SumaHorasGen()
        {
            // Arrange (Fase 416 / Fase 422: Validar que Σ HorasGen == TotalHorasAsignatura)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Subunidad = "FACULTAD_ING";
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-2", new DateTime(2026, 8, 11), new DateTime(2026, 9, 14), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Redes", Horas = 64, IntHoraria = 64 });
            contexto.Aulas.Add(new AgAula { Id = 5, Estado = AgEstadoAula.Disponible });

            // 2 bloques horarios: Martes y Miércoles
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, DuracionHoras = 4, Activo = 1 });

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura
            {
                IdAula = 5,
                IdAsignatura = 10,
                HorasObjetivo = 64,
                DiasPatron = "2,3"
            });
            contexto.ConstruirIndices();

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().NotBeEmpty();
            double sumaHorasGen = 0;
            foreach (var g in individuo.Genes)
            {
                if (g.IdAsignatura == 10)
                {
                    sumaHorasGen += g.HorasGen;
                }
            }
            sumaHorasGen.Should().Be(64.0);
        }

        [Fact]
        public void CrearAleatorio_WhenSinPoolDocente_ShouldAssignIdTerceroVacante()
        {
            // Arrange (Fase 462 / Fase 470)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.IdTerceroVacante = 192;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Algoritmos", Horas = 48, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 10, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().NotBeEmpty();
            individuo.Genes.Should().OnlyContain(g => g.EsVacante == 1 && g.IdTercero == 192);
        }

        [Fact]
        public void CrearAleatorio_WhenAulaGrupoIsNull_ShouldDeriveGrupoFromAulaRef()
        {
            // Arrange (Fase 475)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Algoritmos", Horas = 48, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 10, Grupo = null, AulaRef = "3|1|1|I|A|PRESENCIAL", Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 20, Grupo = null, AulaRef = "3|1|1|I|B|PRESENCIAL", Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            var grupos = individuo.Genes.Select(g => g.Grupo).Distinct().ToList();
            grupos.Should().Contain("A");
            grupos.Should().Contain("B");
        }

        [Fact]
        public void CrearAleatorio_WhenMultiSemestre_ShouldAssignSubjectsMatchingAulaSemester()
        {
            // Arrange (Fase 478)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Cálculo I", Semestre = "I", Horas = 48, HorasSemanales = 4 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "Cálculo II", Semestre = "II", Horas = 48, HorasSemanales = 4 });

            contexto.Aulas.Add(new AgAula { Id = 10, Semestre = "I", Grupo = "A", AulaRef = "3|1|1|I|A|PRESENCIAL", Estado = AgEstadoAula.Disponible });
            contexto.Aulas.Add(new AgAula { Id = 20, Semestre = "II", Grupo = "A", AulaRef = "3|1|1|II|A|VIRTUAL", Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            var genesAula10 = individuo.Genes.Where(g => g.IdAula == 10).ToList();
            genesAula10.Should().NotBeEmpty();
            genesAula10.Should().OnlyContain(g => g.IdAsignatura == 1);

            var genesAula20 = individuo.Genes.Where(g => g.IdAula == 20).ToList();
            genesAula20.Should().NotBeEmpty();
            genesAula20.Should().OnlyContain(g => g.IdAsignatura == 2);
        }

        [Fact]
        public void CrearAleatorio_WhenSemestreDoesNotMatch_ShouldTrackDescartadasSemestreWithTraza()
        {
            // Arrange (Fase 481)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Cálculo I", Semestre = "I", Horas = 48, HorasSemanales = 4 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "Cálculo II", Semestre = "II", Horas = 48, HorasSemanales = 4 });

            contexto.Aulas.Add(new AgAula { Id = 10, Semestre = "I", Grupo = "A", AulaRef = "3|1|1|I|A|PRESENCIAL", Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });

            var logLineas = new List<string>();
            var traza = new AgDiagnosticoTraza("Test", s => logLineas.Add(s)) { Habilitado = true };

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0, traza);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().OnlyContain(g => g.IdAsignatura == 1);
            logLineas.Should().Contain(l => l.Contains("AsignaturasDescartadasPorSemestre: 1"));
        }

        [Fact]
        public void CrearAleatorio_WhenMultiPrograma_ShouldNotAssignSubjectsFromOtherPrograms()
        {
            // Arrange (Fase 664: Aislamiento estricto de asignaturas por programa)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            // Programa 1 (APT-53): CATASTRO en Semestre VII
            contexto.Asignaturas.Add(new Asignatura { Id = 37, Nombre = "CATASTRO", IdPrograma = 1, Semestre = "VII", Horas = 48, HorasSemanales = 4 });
            // Programa 2 (APT-52): FINANZAS PUBLICAS en Semestre VII
            contexto.Asignaturas.Add(new Asignatura { Id = 140, Nombre = "FINANZAS PUBLICAS", IdPrograma = 2, Semestre = "VII", Horas = 48, HorasSemanales = 4 });

            // Aula para Programa 1
            contexto.Aulas.Add(new AgAula { Id = 11, IdPrograma = 1, Semestre = "VII", Grupo = "B", AulaRef = "3|1|1|VII|B|PRESENCIAL", Estado = AgEstadoAula.Disponible });
            // Aula para Programa 2
            contexto.Aulas.Add(new AgAula { Id = 12, IdPrograma = 2, Semestre = "VII", Grupo = "A", AulaRef = "3|1|2|VII|A|PRESENCIAL", Estado = AgEstadoAula.Disponible });

            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            var genesAulaProg1 = individuo.Genes.Where(g => g.IdAula == 11).ToList();
            genesAulaProg1.Should().NotBeEmpty();
            genesAulaProg1.Should().OnlyContain(g => g.IdAsignatura == 37 && g.IdPrograma == 1);

            var genesAulaProg2 = individuo.Genes.Where(g => g.IdAula == 12).ToList();
            genesAulaProg2.Should().NotBeEmpty();
            genesAulaProg2.Should().OnlyContain(g => g.IdAsignatura == 140 && g.IdPrograma == 2);
        }

        [Fact]
        public void CrearAleatorio_PregradoSinBloque4h_GenSinBloqueYVacante()
        {
            // Arrange (Fase 1071): En pregrado sin bloque de 4h para el día requerido, no se debe asignar
            // un bloque arbitrario ni de 8h; el gen debe quedar en cuarentena con IdBloqueHorario = 0 y EsVacante = 1.
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "DERECHO", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "LUN-JUE", HorasDia = 4, Estado = AgEstadoAula.Disponible });

            // Único bloque existente en el contexto es de 8h (p.ej. de posgrado)
            contexto.Bloques.Add(new AgBloqueHorario { Id = 99, DiaSemanaNumero = 1, DuracionHoras = 8.0, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().NotBeEmpty();
            var gen = individuo.Genes.First();
            gen.IdBloqueHorario.Should().Be(0);
            gen.EsVacante.Should().Be(1);
            gen.Estado.Should().Be(AgEstadoGen.Incompleto);
            gen.HorasSemanales.Should().Be(4.0);
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo());
        }

        [Fact]
        public void CrearAleatorio_PregradoSinBloque4hDiaUnico_GenSinBloqueYVacante()
        {
            // Arrange (Fase 1083 / Fase 1096): En rama de día único (diasK.Count == 1), sin bloque de 4h, el gen debe quedar
            // en cuarentena con IdBloqueHorario = 0, EsVacante = 1 y MotivoVacante = SIN_BLOQUE_NIVEL.
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "DERECHO", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "LUN-JUE", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.AulaAsignaturasPorId[10] = new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, DiasPatron = "1" }; // Día único: Lunes

            // Único bloque existente en el contexto es de 8h
            contexto.Bloques.Add(new AgBloqueHorario { Id = 99, DiaSemanaNumero = 1, DuracionHoras = 8.0, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().NotBeEmpty();
            var gen = individuo.Genes.First();
            gen.IdBloqueHorario.Should().Be(0);
            gen.EsVacante.Should().Be(1);
            gen.Estado.Should().Be(AgEstadoGen.Incompleto);
            gen.HorasSemanales.Should().Be(4.0);
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo());
        }

        [Fact]
        public void Fabrica_DosIndividuosMismoContexto_CalculaPlanUnaVez()
        {
            // Arrange (Fase 1090 & Fase 1101): Memoización verificada mediante contador de construcciones
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 101, Nombre = "ASIG 1", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "MAR-VIE", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, DuracionHoras = 4.0, Activo = 1 });

            var planner = new AgPlanificadorOleadas();
            var sut = new AgFabricaIndividuos(null, null, planner);

            // Act
            sut.CrearAleatorio(contexto, 1, 0);
            sut.CrearAleatorio(contexto, 1, 1);

            // Assert: se calculó exactamente 1 vez el plan de oleadas
            planner.Construcciones.Should().Be(1);
        }

        [Fact]
        public void Fabrica_SinBloqueNivel_UsaSinBloqueNivel()
        {
            // Arrange (Fase 1096): Sin bloque de 4h para pregrado -> motivo SIN_BLOQUE_NIVEL
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "DERECHO", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "LUN-JUE", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            // Sin bloques de 4h
            contexto.Bloques.Add(new AgBloqueHorario { Id = 99, DiaSemanaNumero = 1, DuracionHoras = 8.0, Activo = 1 });

            // Act
            var individuo = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            var gen = individuo.Genes.First();
            gen.MotivoVacante.Should().Be(AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo());
        }

        [Fact]
        public void Fabrica_AulaSinVentana_GenSinBloqueYEstadoIncompleto()
        {
            // Arrange (Fase 1099): Aula sin DiasVentana -> cuarentena efectiva (IdBloqueHorario=0, Incompleto, SIN_PLAN_OLEADA)
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "ASIG 1", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1 });

            // Act
            var ind = _sut.CrearAleatorio(contexto, 1, 0);

            // Assert: gen no tiene bloque asignado, patrón vacío, estado incompleto y motivo SIN_PLAN_OLEADA
            ind.Genes.Should().NotBeEmpty();
            var g = ind.Genes.First();
            g.IdBloqueHorario.Should().Be(0);
            g.DiasPatron.Should().BeEmpty();
            g.Estado.Should().Be(AgEstadoGen.Incompleto);
            g.MotivoVacante.Should().Be(AgMotivoVacante.SIN_PLAN_OLEADA.Codigo());
        }

        [Fact]
        public void Fabrica_NuevaCorrida_MismoAula_RecalculaPlan()
        {
            // Arrange (Fase 1102): Dos corridas distintas con la misma aula deben recalcular el plan al cambiar IdCorrida
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 101, Nombre = "ASIG 1", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "MAR-VIE", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, DuracionHoras = 4.0, Activo = 1 });

            var planner = new AgPlanificadorOleadas();
            var sut = new AgFabricaIndividuos(null, null, planner);

            // Act
            sut.CrearAleatorio(contexto, 1, 0); // Corrida 1
            sut.CrearAleatorio(contexto, 2, 0); // Corrida 2 (invalida caché y recalcula)

            // Assert: se calculó 2 veces
            planner.Construcciones.Should().Be(2);
        }

        [Fact]
        public void Fabrica_PregradoConBloques4hY8h_SoloAsignaBloquesDe4h()
        {
            // Arrange (Fase 1114.d): Entorno de PREGRADO con bloques de 4h y 8h activos.
            // La fábrica debe asignar exclusivamente bloques de 4h (nunca de 8h).
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 101, Nombre = "ASIG 1", IdPrograma = 1, Semestre = "I", Horas = 32, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "MAR-VIE", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 10, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 20, DiaSemanaNumero = 3, DuracionHoras = 4.0, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 80, DiaSemanaNumero = 2, DuracionHoras = 8.0, Activo = 1 });

            var sut = new AgFabricaIndividuos();

            // Act
            var ind = sut.CrearAleatorio(contexto, 1, 0);

            // Assert: los genes asignados deben tener exclusivamente bloques de 4h
            ind.Genes.Should().NotBeEmpty();
            ind.Genes.Where(g => g.IdBloqueHorario > 0).Should().OnlyContain(g => g.IdBloqueHorario != 80);
            ind.Genes.Where(g => g.IdBloqueHorario > 0).Should().OnlyContain(g => g.HorasSemanales == 4.0);
        }

        [Fact]
        public void Fabrica_Sabado4h_AsignaFranjasMananaYTardeDistintas()
        {
            // Arrange (Fase 1292.a / Fase 1299): Aula 4h con ventana JUE-SAB donde sábado tiene 2 ranuras (MANANA y TARDE) en carril 1
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 101, Nombre = "ASIG JUE-VIE", IdPrograma = 1, Semestre = "I", Horas = 48, HorasSemanales = 4 });
            // Fase 1299: segunda asignatura para alcanzar el carril 1 ([6:MANANA, 6:TARDE])
            contexto.Asignaturas.Add(new Asignatura { Id = 102, Nombre = "ASIG SAB", IdPrograma = 1, Semestre = "I", Horas = 48, HorasSemanales = 4 });
            contexto.Aulas.Add(new AgAula { Id = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", DiasVentana = "JUE-SAB", HorasDia = 4, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 4, DuracionHoras = 4.0, HoraInicio = "18:00", HoraFin = "22:00", Franja = "NOCHE", Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 5, DuracionHoras = 4.0, HoraInicio = "18:00", HoraFin = "22:00", Franja = "NOCHE", Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, DuracionHoras = 4.0, HoraInicio = "08:00", HoraFin = "12:00", Franja = "MANANA", Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 6, DuracionHoras = 4.0, HoraInicio = "13:00", HoraFin = "17:00", Franja = "TARDE", Activo = 1 });

            var sut = new AgFabricaIndividuos();

            // Act
            var ind = sut.CrearAleatorio(contexto, 1, 0);

            // Assert: los genes del sábado deben materializarse con franjas distintas (MANANA y TARDE)
            var grupoSab = ind.Genes.Where(g => g.DiaSemana == 6).GroupBy(g => g.IdAsignatura).ToList();
            grupoSab.Should().ContainSingle();
            grupoSab[0].Should().HaveCount(2);
            var franjasSab = grupoSab[0]
                .Select(g => contexto.Bloques.First(b => b.Id == g.IdBloqueHorario).Franja)
                .ToList();
            franjasSab.Should().BeEquivalentTo(new[] { "MANANA", "TARDE" });
        }

        [Fact]
        public void Fabrica_Sabado8h_Posgrado_ConservaFranjaUnica()
        {
            // Arrange (Fase 1292.b): Posgrado 8h en sábado debe conservar franja única MANANA (08:00-17:00)
            var contexto = new AgContexto();
            contexto.IdPrograma = 2;
            contexto.Programas.Add(new Programa { Id = 2, Modalidad = "POSGRADO ESPECIALIZACION" });
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            contexto.Asignaturas.Add(new Asignatura { Id = 201, Nombre = "ASIG POS", IdPrograma = 2, Semestre = "I", Horas = 32, HorasSemanales = 8 });
            contexto.Aulas.Add(new AgAula { Id = 2, IdPrograma = 2, Semestre = "I", Grupo = "A", DiasVentana = "VIE-SAB", HorasDia = 8, Estado = AgEstadoAula.Disponible });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 5, DuracionHoras = 8.0, HoraInicio = "08:00", HoraFin = "17:00", Franja = "MANANA", Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 6, DiaSemanaNumero = 6, DuracionHoras = 8.0, HoraInicio = "08:00", HoraFin = "17:00", Franja = "MANANA", Activo = 1 });

            var sut = new AgFabricaIndividuos();

            // Act
            var ind = sut.CrearAleatorio(contexto, 1, 0);

            // Assert
            var genesSabado = ind.Genes.Where(g => g.DiaSemana == 6).ToList();
            genesSabado.Should().NotBeEmpty();
            genesSabado.Should().OnlyContain(g => g.IdBloqueHorario == 6);
        }
    }
}
