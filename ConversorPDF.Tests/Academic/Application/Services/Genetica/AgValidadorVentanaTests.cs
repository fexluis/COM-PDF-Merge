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
    public class AgValidadorVentanaTests
    {
        private readonly AgValidadorVentana _sut = new AgValidadorVentana();

        [Fact]
        public void ValidarFactibilidad_WhenContextIsNull_ShouldReturnFailure()
        {
            // Act
            var result = _sut.ValidarFactibilidad(null);

            // Assert
            result.IsFailure.Should().BeTrue();
        }

        [Fact]
        public void ValidarFactibilidad_WhenNoWindowDefined_ShouldHaveBlockingIssues()
        {
            // Arrange
            var contexto = new AgContexto();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("ventana temporal"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenNoAulasAvailable_ShouldHaveBlockingIssues()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1, DuracionHoras = 4 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("aulas"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenResourcesSufficient_ShouldBeFeasible()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Matemáticas", Horas = 32 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Bloqueos.Should().BeEmpty();
            result.Value.EsFactible.Should().BeTrue();
        }

        [Fact]
        public void ValidarFactibilidad_WhenHorasExceedOccurrenceCapacity_ShouldBlock()
        {
            // Arrange
            var contexto = new AgContexto();
            // Ventana corta de 2 semanas (3 lunes: 2, 9 y 16 de febrero = 12h de capacidad)
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 2, 16), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Matemáticas", Horas = 32 }); // Requiere 32h > 12h
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("Capacidad insuficiente"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenPatternDaysHaveZeroOrInsufficientOccurrences_ShouldBlockWithCapacidadReal()
        {
            // Arrange (Fase 586): ventana lunes-viernes (0 sábados), asignatura con patrón sábado ("6") y 48h
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 2, 6), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 6, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });

            var asig = new Asignatura { Id = 10, Nombre = "Química", Horas = 48 };
            contexto.Asignaturas.Add(asig);
            contexto.AulaAsignaturasPorId[10] = new AgAulaAsignatura
            {
                IdAsignatura = 10,
                DiasPatron = "6",
                HorasObjetivo = 48
            };

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("capacidad real"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenMultipleClassroomsForSubject_ShouldUseMaxCapacityAcrossClassrooms()
        {
            // Arrange (Fase 591): dos aulas para la misma asignatura (una sin sábados con DiasPatron="6" y otra viable con DiasPatron="2,3")
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 3, 31), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, DiasVentana = "VIE-SAB", Capacidad = 30 });
            contexto.Aulas.Add(new AgAula { Id = 2, Estado = AgEstadoAula.Disponible, DiasVentana = "MAR-VIE", Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });

            var asig = new Asignatura { Id = 10, Nombre = "Química", Horas = 48 };
            contexto.Asignaturas.Add(asig);

            var aaInviable = new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, DiasPatron = "6", HorasObjetivo = 48 };
            var aaViable = new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10, DiasPatron = "2,3", HorasObjetivo = 48 };

            contexto.AulaAsignaturas.Add(aaInviable);
            contexto.AulaAsignaturas.Add(aaViable);
            // El índice monolítico por IdAsignatura simula last-write-wins con el inviable
            contexto.AulaAsignaturasPorId[10] = aaInviable;

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: no debe bloquearse falsamente porque el aula 2 posee capacidad suficiente
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeTrue();
            result.Value.Bloqueos.Should().BeEmpty();
        }

        [Fact]
        public void ValidarFactibilidad_WhenSubjectHasNoAulaAsignaturaRow_ShouldNotBlockFalsely()
        {
            // Arrange (Fase 601): asignatura sin fila en AulaAsignaturas no debe obtener capacidadReal = 0
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Química", Horas = 48 }); // sin fila propia
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 99, DiasPatron = "6", HorasObjetivo = 16 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeTrue();
            result.Value.Bloqueos.Should().BeEmpty();
        }

        [Fact]
        public void ValidarFactibilidad_AulaSinDiasVentana_BloqueaI1()
        {
            // Arrange (Fase 1093): Aula con asignaturas pero sin dias_ventana debe generar bloqueo I1_VENTANA_VACIA
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-SIN-VENTANA", DiasVentana = "", Estado = AgEstadoAula.Disponible });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, HorasObjetivo = 32 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("I1_VENTANA_VACIA"));
        }

        [Fact]
        public void ValidarVentanaAdvertencia_VentanaImpar_ReportaCarrilesCanonicos()
        {
            // Arrange (Fase 1094 / Fase 1154): Ventana de 3 días ("MAR,MIE,JUE") computa AgCarriles.Cantidad = 2 carriles ([2,3] y [4]) sin descartar días
            var ventana = AgDiasVentana.Dias("MAR,MIE,JUE");

            // Act
            int carriles = ConversorPDF.XLL.Academic.Domain.Models.Genetica.AgCarriles.Cantidad(ventana);

            // Assert
            carriles.Should().Be(2);
        }

        [Fact]
        public void ValidarFactibilidad_HorasObjetivoCero_NoBloqueaYAdvierte()
        {
            // Arrange (Fase 1105 / Fase 1107): Fila de AulaAsignaturas con HorasObjetivo = 0 y sin IntHoraria
            // no debe asumir 64h ficticias que bloqueen la corrida con I5_INFACTIBLE, sino advertir I5_HORAS_DESCONOCIDAS.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-1", DiasVentana = "LUN,MIE", Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, HorasObjetivo = 0 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor X", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: No debe bloquear por I5_INFACTIBLE y debe incluir advertencia I5_HORAS_DESCONOCIDAS
            result.IsSuccess.Should().BeTrue();
            result.Value.Bloqueos.Should().NotContain(b => b.Contains("I5_INFACTIBLE"));
            result.Value.Advertencias.Should().Contain(a => a.Contains("I5_HORAS_DESCONOCIDAS"));
        }

        [Fact]
        public void ValidarFactibilidad_ConHorasDesconocidas_ReportaMateriasComputadas()
        {
            // Arrange (Fase 1110 / Fase 1112): Aula con 2 asignaturas asignadas: 1 computada (32h) y 1 con horas=0 (omitida).
            // Ventana corta de 1 semana (2 ranuras) => demanda de 8 sesiones excede la capacidad de 2 ranuras.
            // El mensaje de bloqueo I5 debe reportar "(1 materias computadas de 2)".
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 2, 6), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-TEST", DiasVentana = "LUN,MIE", Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, HorasObjetivo = 32 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 20, HorasObjetivo = 0 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: Debe emitir bloqueo I5 indicando "1 materias computadas de 2"
            result.IsSuccess.Should().BeTrue();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("1 materias computadas de 2"));
            result.Value.Advertencias.Should().Contain(a => a.Contains("I5_HORAS_DESCONOCIDAS"));
        }

        [Fact]
        public void ValidarFactibilidad_PregradoConHorasDia8_BloqueaD4()
        {
            // Arrange (Fase 1114.c / Fase 1114.d): Aula de PREGRADO configurada con HorasDia = 8.
            // ValidarFactibilidad debe emitir bloqueo explícito [D4_PREGRADO_8H].
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-PREGRADO-8H", IdPrograma = 1, DiasVentana = "LUN,MIE", HorasDia = 8, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, HorasObjetivo = 32 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: Debe emitir bloqueo D4_PREGRADO_8H
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("[D4_PREGRADO_8H]"));
        }

        [Fact]
        public void ValidarFactibilidad_AsignaturaExcedeSuOleada_BloqueaI5()
        {
            // Arrange (Fase 1122 / Fase 1145): Ventana de 12 semanas (84 días: 2026-02-02 a 2026-04-26).
            // Aula con 4 días ("LUN,MAR,MIE,JUE") ⇒ 2 carriles. Con 2 asignaturas ⇒ 1 oleada de 24 ranuras por carril.
            // Asignatura 1 requiere 128h (32 sesiones > 24 ranuras máximas de su carril en toda la ventana).
            // Demanda total: 32 + 4 = 36 <= 48 capacidad global del aula (no dispara [I5_INFACTIBLE] global).
            // La validación de sub-ventana por oleada debe detectar que la asignatura 1 excede su oleada y emitir [I5_INFACTIBLE_OLEADA].
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 4, 26), AgOrigenVentana.Presencial).Value;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-OLEADAS", IdPrograma = 1, DiasVentana = "LUN,MAR,MIE,JUE", HorasDia = 4, Estado = AgEstadoAula.Disponible, Capacidad = 30 });

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 128, OrdenCurricular = 1 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 2, HorasObjetivo = 16, OrdenCurricular = 2 });

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG 1", Horas = 128 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "ASIG 2", Horas = 16 });

            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 4, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: debe bloquear por capacidad real de carril u oleada
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => (b.Contains("[I5_INFACTIBLE_OLEADA]") || b.Contains("capacidad real")) && b.Contains("1"));
        }

        [Fact]
        public void ValidarFactibilidad_Pregrado8h_NoDuplicaCapacidadPorDia()
        {
            // Arrange (Fase 1123 / Fase 1126): Ventana corta de 1 semana (2026-02-02 a 2026-02-06, 1 lunes y 1 miércoles = 2 ocurrencias).
            // Asignatura de 16h (requiere 4 sesiones de 4h).
            // Con 1 sesión/día/carril (R4), la capacidad real es de 2 sesiones (8h) < 4 sesiones requeridas.
            // Si permitiera 2 sesiones/día por HorasDia=8, la capacidad subiría a 4 sesiones aprobando erróneamente la corrida.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 2, 6), AgOrigenVentana.Presencial).Value;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "POSGRADO ESPECIALIZACION" }); // Posgrado para que no bloquee D4_PREGRADO_8H
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-CAPACIDAD", IdPrograma = 1, DiasVentana = "LUN,MIE", HorasDia = 8, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, DiasPatron = "1,3", HorasObjetivo = 32 });
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "ASIG TEST", Horas = 32 });

            // Bloques de mañana y tarde (2 franjas de 4h)
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4, Franja = "MANANA" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4, Franja = "TARDE" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4, Franja = "MANANA" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4, Franja = "TARDE" });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: debe rechazar porque capacidadReal es 2 sesiones (1 sesión/día), no 4
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b => b.Contains("capacidad real"));
        }

        [Fact]
        public void ValidarFactibilidad_AulaConAsignaturaDeOtroSemestre_NoBloquea()
        {
            // Arrange (Fase 1129 / Fase 1130): Aula con Semestre = "I".
            // AulaAsignaturas contiene Asig 1 (Semestre "I", 16h) y Asig 2 (Semestre "II", 64h).
            // La sub-ventana de oleada solo debe evaluar Asig 1 (Semestre "I"); Asig 2 pertenece a otro semestre
            // y no entra al alcance del motor, evitando un falso bloqueo [I5_INFACTIBLE_OLEADA].
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Programas.Add(new Programa { Id = 1, Modalidad = "PREGRADO PRESENCIAL" });
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-SEM-I", IdPrograma = 1, Semestre = "I", DiasVentana = "LUN,MIE", HorasDia = 4, Estado = AgEstadoAula.Disponible, Capacidad = 30 });

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 16, OrdenCurricular = 1 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 2, HorasObjetivo = 64, OrdenCurricular = 2 });

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG SEM 1", IdPrograma = 1, Semestre = "I", Horas = 16 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "ASIG SEM 2", IdPrograma = 1, Semestre = "II", Horas = 64 });

            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: Asig 2 no debe evaluarse en la sub-ventana del aula semestre I
            result.IsSuccess.Should().BeTrue();
            result.Value.Bloqueos.Should().NotContain(b => b.Contains("[I5_INFACTIBLE_OLEADA]"));
        }

        [Fact]
        public void ValidarFactibilidad_CapacidadSeCalculaConDiasHabilesReales()
        {
            // Arrange (Fase 1142 / Fase 1144): MAR-VIE, 2 semanas, un martes festivo -> capacidad real 7 (no 8)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 8, 10), new DateTime(2026, 8, 23), AgOrigenVentana.Presencial).Value;
            contexto.Festivos.Add((new DateTime(2026, 8, 18), new DateTime(2026, 8, 18)));
            contexto.Aulas.Add(new AgAula { Id = 1, AulaRef = "AULA-1", IdPrograma = 1, Semestre = "I", DiasVentana = "MAR,MIE,JUE,VIE", HorasDia = 4, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 32, OrdenCurricular = 1 });
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG", IdPrograma = 1, Semestre = "I", Horas = 32 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 4, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 5, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.Value.Bloqueos.Should().Contain(b => b.Contains("capacidad real de 7 ranuras"));
        }

        [Fact]
        public void ValidarFactibilidad_BloqueoI5_IncluyeNombreDeSede()
        {
            // Arrange (Fase 1579): el bloqueo [I5_INFACTIBLE] debe identificar la sede además del AulaRef.
            // Mismo escenario que el test de festivos (7 ranuras reales < 8 sesiones requeridas) para forzar el bloqueo.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 8, 10), new DateTime(2026, 8, 23), AgOrigenVentana.Presencial).Value;
            contexto.Festivos.Add((new DateTime(2026, 8, 18), new DateTime(2026, 8, 18)));
            contexto.Sedes.Add(new Sede { Id = 13, Ciudad = "BOGOTA", Subsede = "SUBSEDE NORTE" });
            contexto.Aulas.Add(new AgAula { Id = 1, IdSede = 13, AulaRef = "1|13|1|I|A|PRESENCIAL", IdPrograma = 1, Semestre = "I", DiasVentana = "MAR,MIE,JUE,VIE", HorasDia = 4, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 32, OrdenCurricular = 1 });
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG", IdPrograma = 1, Semestre = "I", Horas = 32 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 4, Activo = 1, DuracionHoras = 4 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 5, Activo = 1, DuracionHoras = 4 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert: el mensaje incluye la sede en catálogo entre el código y el AulaRef
            result.Value.Bloqueos.Should().Contain(b =>
                b.Contains("[I5_INFACTIBLE]") &&
                b.Contains("Sede 'BOGOTA (SUBSEDE NORTE)'") &&
                b.Contains("Aula '1|13|1|I|A|PRESENCIAL'"));
        }

        [Fact]
        public void ValidarFactibilidad_MateriaSinPoolDocente_PermiteVacanteYGeneracionContinua()
        {
            // Arrange (Fase 1602 / Fase 1603): Materia sin docentes habilitados en PoolDocente genera advertencia
            // de vacante (AgMotivoVacante.SIN_CANDIDATOS), pero NO bloquea la factibilidad
            // mientras existan docentes en catálogo para otras asignaturas.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Aulas.Add(new AgAula { Id = 1, Estado = AgEstadoAula.Disponible, Capacidad = 30 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Activo = 1, DuracionHoras = 4 });

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Matemáticas Con Docente", Horas = 32 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "Física Sin Docente", Horas = 32 });

            contexto.Docentes.Add(new Tercero { Id = 101, Nombre = "Docente Matemáticas", Activo = 1 });

            // F1603: Registrar docente habilitado formalmente en el pool solo para la asignatura 1
            contexto.RegistrarDocenteHabilitado(101, 0, 1);
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.EsFactible.Should().BeTrue();
            result.Value.AsignaturasVacantes.Should().ContainKey(2);
            result.Value.AsignaturasVacantes[2].Should().Be(AgMotivoVacante.SIN_CANDIDATOS);
            result.Value.Advertencias.Should().Contain(a => a.Contains("Física Sin Docente") && a.Contains("no tiene ningún docente habilitado en el Pool"));
        }

        [Fact]
        public void ValidarFactibilidad_PatronSaturadoEnAula_ReportaBloqueoI5DiaSaturado()
        {
            // Arrange (Fase 1601 / Fase 1602 / Fase 1604): Aula con ventana amplia (1..5)
            // donde las asignaturas tienen patrón exclusivo "5" (viernes) saturando la capacidad del carril compatible.
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 3, 2), new DateTime(2026, 3, 22), AgOrigenVentana.Presencial).Value; // 3 semanas
            contexto.Sedes.Add(new Sede { Id = 5, Ciudad = "MEDELLIN", Subsede = "CENTRAL" });
            var aula = new AgAula
            {
                Id = 1,
                IdSede = 5,
                AulaRef = "MED-01",
                IdPrograma = 1,
                Semestre = "I",
                DiasVentana = "1,2,3,4,5",
                Jornada = "NOCHE",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30
            };
            contexto.Aulas.Add(aula);

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 12, OrdenCurricular = 1, DiasPatron = "5" });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 2, HorasObjetivo = 12, OrdenCurricular = 2, DiasPatron = "5" });

            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG 1", IdPrograma = 1, Semestre = "I", Horas = 12 });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "ASIG 2", IdPrograma = 1, Semestre = "I", Horas = 12 });

            for (int d = 1; d <= 5; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario { Id = d, DiaSemanaNumero = d, Franja = "NOCHE", DuracionHoras = 4, Activo = 1 });
            }

            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b =>
                b.Contains("[I5_DIA_SATURADO]") &&
                b.Contains("Sede 'MEDELLIN (CENTRAL)'") &&
                b.Contains("Aula 'MED-01'") &&
                b.Contains("saturación de franja") &&
                b.Contains("semanas"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenDemandaExcedeCapacidadEnVieDom_ShouldSugerirJueDomEnBloqueoCritico()
        {
            // Arrange (W = 4 semanas: 2026-08-07 al 2026-08-30, VIE-DOM capacidad = 4W = 16 ranuras; Demanda = 20 sesiones = 5W)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-2", new DateTime(2026, 8, 7), new DateTime(2026, 8, 30), AgOrigenVentana.Presencial).Value;

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-VIE-DOM",
                DiasVentana = "VIE-DOM",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 30
            };
            contexto.Aulas.Add(aula);

            // Bloques activos para JUE (4), VIE (5), SAB (6M, 6T), DOM (7)
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 5, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 6, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, Franja = "TARDE", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 7, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 5, DiaSemanaNumero = 4, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });

            // Asignatura con 80h = 20 sesiones de 4h (> 16 ranuras)
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "ASIG 1", Horas = 80 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 1, HorasObjetivo = 80, OrdenCurricular = 1 });
            contexto.Docentes.Add(new Tercero { Id = 1, Nombre = "Profesor", Activo = 1 });
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b =>
                b.Contains("[I5_INFACTIBLE]") &&
                b.Contains("Escale la ventana a 'JUE-DOM'"));
        }

        [Fact]
        public void ValidarFactibilidad_WhenEscalonNoGanaRanurasPorFaltaDeBloques_ShouldEmitirMensajeAccionable()
        {
            // Arrange (LUN-JUE sin bloque de viernes en la jornada)
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", new DateTime(2027, 2, 5), new DateTime(2027, 5, 30), AgOrigenVentana.Manual).Value;

            var aula = new AgAula
            {
                Id = 2,
                AulaRef = "AULA-LUN-JUE",
                DiasVentana = "LUN-JUE",
                Jornada = "MANANA",
                HorasDia = 4,
                Estado = AgEstadoAula.Disponible,
                Capacidad = 35
            };
            contexto.Aulas.Add(aula);

            // Bloques solo para LUN..JUE (1..4), sin bloque de viernes (5) para MANANA
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 4, DiaSemanaNumero = 4, Franja = "MANANA", DuracionHoras = 4, Activo = 1 });

            // Demanda que excede la capacidad de LUN-JUE
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "ASIG DEMANDA", Horas = 300 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10, HorasObjetivo = 300, OrdenCurricular = 1 });
            contexto.Docentes.Add(new Tercero { Id = 2, Nombre = "Docente", Activo = 1 });
            contexto.ConstruirIndices();

            // Act
            var result = _sut.ValidarFactibilidad(contexto);

            // Assert
            result.Value.EsFactible.Should().BeFalse();
            result.Value.Bloqueos.Should().Contain(b =>
                b.Contains("[I5_INFACTIBLE]") &&
                b.Contains("no añade días") &&
                b.Contains("no existe bloque activo"));
        }

        [Fact]
        public void CalcularDeficitEstructural_WhenAulasMultiplesYDocentesInsuficientes_ShouldComputeCorrectDeficit()
        {
            // Arrange (Fase 2026, 2036): 3 cohortes/aulas para Asig 10 (con 2 docentes en pool) y 1 Asig 20 consentida
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Matemáticas" });
            contexto.Asignaturas.Add(new Asignatura { Id = 20, Nombre = "Física (Consentida)" });

            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 1, Semestre = "1", Grupo = "B" };
            contexto.AulasPorId[3] = new AgAula { Id = 3, IdSede = 2, Semestre = "1", Grupo = "A" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 3, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 20 });

            contexto.BloquesPorId = new Dictionary<int, AgBloqueHorario> { { 1, new AgBloqueHorario { Id = 1, Activo = 1 } } };
            contexto.AsignaturasVacantesConsentidas = new HashSet<int> { 20 }; // Excluida del déficit

            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 1, Activo = 1 },
                new TerceroAsignatura { IdTercero = 2, Activo = 1 }
            };

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitEstructural(contexto);

            // Assert: 3 cohortes - 2 docentes = 1; Asig 20 ignorada
            deficit.Should().Be(1);

            // Contexto incompleto devuelve -1 (desactivado)
            AgValidadorVentana.CalcularDeficitEstructural(new AgContexto()).Should().Be(-1);
        }

        [Fact]
        public void CalcularDeficitEstructural_Corrida109_ShouldAllowAlMenosOchoVacantesForzadas()
        {
            // Arrange (Fase 2036 & 2040): Simulación de demanda estilo corrida 109
            // Asignatura 1: 8 cohortes con 3 docentes en pool => déficit = 5
            // Asignatura 2: 5 cohortes con 2 docentes en pool => déficit = 3
            // Total déficit = 8 vacantes forzadas permitidas (cap >= 8)
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 1, Nombre = "Fundamentos" });
            contexto.Asignaturas.Add(new Asignatura { Id = 2, Nombre = "Estadística" });

            for (int i = 1; i <= 8; i++)
            {
                contexto.AulasPorId[i] = new AgAula { Id = i, IdSede = i % 2 + 1, Semestre = "1", Grupo = $"G{i}" };
                contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = i, IdAsignatura = 1 });
            }

            for (int i = 9; i <= 13; i++)
            {
                contexto.AulasPorId[i] = new AgAula { Id = i, IdSede = 1, Semestre = "2", Grupo = $"G{i}" };
                contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = i, IdAsignatura = 2 });
            }

            contexto.BloquesPorId = new Dictionary<int, AgBloqueHorario>
            {
                { 1, new AgBloqueHorario { Id = 1, Activo = 1 } }
            };

            contexto.PoolPorAsignatura[1] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 101, Activo = 1 },
                new TerceroAsignatura { IdTercero = 102, Activo = 1 },
                new TerceroAsignatura { IdTercero = 103, Activo = 1 }
            };

            contexto.PoolPorAsignatura[2] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 201, Activo = 1 },
                new TerceroAsignatura { IdTercero = 202, Activo = 1 }
            };

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitEstructural(contexto);

            // Assert: (8 - 3) + (5 - 2) = 8
            deficit.Should().Be(8);
            deficit.Should().BeGreaterThanOrEqualTo(8);
        }

        [Fact]
        public void CalcularDeficitEstructural_WhenConsentidasPorClaveSede_ShouldExcludeOnlyThatCohorte()
        {
            // Arrange (Fase 2036 & 2040): Asig 10 con 2 cohortes en Sede 1 y Sede 2.
            // Clave "10:1" consentida => solo Sede 2 genera déficit (1 cohorte - 0 docentes = 1).
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Derecho" });

            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 2, Semestre = "1", Grupo = "A" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10 });

            contexto.BloquesPorId = new Dictionary<int, AgBloqueHorario> { { 1, new AgBloqueHorario { Id = 1, Activo = 1 } } };
            contexto.ClavesVacantesConsentidas.Add("10:1"); // Solo Sede 1 excluida

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitEstructural(contexto);

            // Assert: 1 cohorte en Sede 2 - 0 docentes = 1 (Sede 1 fue excluida por clave granular)
            deficit.Should().Be(1);
        }

        [Fact]
        public void CalcularDeficitEstructural_WhenAulasCompartenMismaCohorte_ShouldNotCountAulasSeparadas()
        {
            // Arrange (Fase 2042): 2 aulas físicas comparten la misma cohorte académica (Semestre 1, Grupo A, Sede 1)
            // Sin IdAula en la tupla, debe contar como 1 cohorte, no 2.
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Bioética" });

            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 1, Semestre = "1", Grupo = "A" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10 });

            contexto.BloquesPorId = new Dictionary<int, AgBloqueHorario> { { 1, new AgBloqueHorario { Id = 1, Activo = 1 } } };
            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 1, Activo = 1 }
            };

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitEstructural(contexto);

            // Assert: 1 cohorte efectiva - 1 docente = 0 (no sobreestima a 2 aulas - 1 = 1)
            deficit.Should().Be(0);
        }

        [Fact]
        public void CalcularDeficitEstructural_WhenInformeProvidesDeficit_ShouldReturnSumFromInforme()
        {
            // Arrange (Fase 2042): Informe de factibilidad suministra cálculo directo de déficit docente
            var contexto = new AgContexto();
            var informe = new InformeFactibilidad
            {
                DeficitDocentePorAsignatura = new Dictionary<int, int>
                {
                    { 10, 4 },
                    { 20, 2 }
                }
            };

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitEstructural(contexto, informe);

            // Assert
            deficit.Should().Be(6);
        }

        [Fact]
        public void ValidarFactibilidad_WhenDemandaExcedeDocentesEnMismaJornada_ShouldPopulateDeficitDocentePorAsignatura()
        {
            // Arrange (Fase 2046 / Fase 2090 / Fase 2096 / Fase 2101): Asignatura 10 con 3 cohortes académicas (A, B en Mañana, C en Tarde).
            // Con la fórmula unificada de concurrencia por jornada, maxConcurrentes = 2 (Mañana) - 1 docente en pool => déficit estructural = 1.
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 10, Nombre = "Economía", Horas = 32 });
            contexto.Ventana = new VentanaSemestre
            {
                FechaInicio = DateTime.Today,
                FechaFin = DateTime.Today.AddDays(90)
            };

            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A", Jornada = "MANANA", DiasVentana = "1,2,3,4,5" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 1, Semestre = "1", Grupo = "B", Jornada = "MANANA", DiasVentana = "1,2,3,4,5" };
            contexto.AulasPorId[3] = new AgAula { Id = 3, IdSede = 1, Semestre = "1", Grupo = "C", Jornada = "TARDE", DiasVentana = "1,2,3,4,5" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 3, IdAsignatura = 10 });

            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1, DuracionHoras = 4, Franja = "MANANA" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, Activo = 1, DuracionHoras = 4, Franja = "TARDE" });

            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 101, Activo = 1 }
            };

            // Act
            var res = _sut.ValidarFactibilidad(contexto);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.DeficitDocentePorAsignatura.Should().ContainKey(10);
            res.Value.DeficitDocentePorAsignatura[10].Should().Be(1);
        }

        [Fact]
        public void ValidarFactibilidad_WhenDosJornadasYUnDocente_ShouldHaveDeficitCero()
        {
            // Arrange (Fase 2096): 2 cohortes en jornadas distintas (Mañana y Tarde) cubiertas por 1 solo docente => maxConcurrentes = 1, déficit = 0
            var contexto = new AgContexto();
            contexto.Asignaturas.Add(new Asignatura { Id = 20, Nombre = "Derecho", Horas = 32 });
            contexto.Ventana = new VentanaSemestre
            {
                FechaInicio = DateTime.Today,
                FechaFin = DateTime.Today.AddDays(90)
            };

            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A", Jornada = "MANANA", DiasVentana = "1,2,3,4,5" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 1, Semestre = "1", Grupo = "A", Jornada = "TARDE", DiasVentana = "1,2,3,4,5" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 20 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 20 });

            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, Activo = 1, DuracionHoras = 4, Franja = "MANANA" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, Activo = 1, DuracionHoras = 4, Franja = "TARDE" });

            contexto.PoolPorAsignatura[20] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 201, Activo = 1 }
            };

            // Act
            var res = _sut.ValidarFactibilidad(contexto);

            // Assert: no debe haber déficit estructural para asignatura 20
            res.IsSuccess.Should().BeTrue();
            res.Value.DeficitDocentePorAsignatura.Should().NotContainKey(20);
        }

        [Fact]
        public void ValidarFactibilidad_WhenContextoIncompleto_ShouldNotMutaMaxVacantesForzadas()
        {
            // Arrange (Fase 2027: ValidarFactibilidad es de solo lectura y no muta MaxVacantesForzadas)
            var contexto = new AgContexto();
            contexto.MaxVacantesForzadas = -1;

            // Act
            _sut.ValidarFactibilidad(contexto);

            // Assert: se preserva -1 sin mutar a 0
            contexto.MaxVacantesForzadas.Should().Be(-1);
        }

        [Fact]
        public void CalcularDeficitPorAsignatura_WhenInvocadoDirectamente_ShouldComputeIdenticalDeficit()
        {
            // Arrange (Fase 2090, 2095): 3 cohortes con 1 docente activo => déficit = 2
            var contexto = new AgContexto();
            contexto.AulasPorId[1] = new AgAula { Id = 1, IdSede = 1, Semestre = "1", Grupo = "A" };
            contexto.AulasPorId[2] = new AgAula { Id = 2, IdSede = 1, Semestre = "1", Grupo = "B" };
            contexto.AulasPorId[3] = new AgAula { Id = 3, IdSede = 2, Semestre = "1", Grupo = "A" };

            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 2, IdAsignatura = 10 });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = 3, IdAsignatura = 10 });

            contexto.PoolPorAsignatura[10] = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 101, Activo = 1 }
            };

            // Act
            int deficit = AgValidadorVentana.CalcularDeficitPorAsignatura(contexto, 10, out int cohortes, out int docs);

            // Assert
            cohortes.Should().Be(3);
            docs.Should().Be(1);
            deficit.Should().Be(2);
        }

        [Fact]
        public void ValidarFactibilidad_WhenEscalonNoReducePendientesPorPatronesRigidos_ShouldEmitirAdvertenciaPatrones()
        {
            // Arrange (Fase 2110 / Fase 2111 / Fase 2114)
            var contexto = new AgContexto();
            var inicio = new DateTime(2027, 2, 5);
            var fin = new DateTime(2027, 5, 30);
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", inicio, fin, AgOrigenVentana.Presencial).Value;

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "3|1|1|I|A",
                DiasVentana = "MAR-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Capacidad = 30,
                IdPrograma = 1
            };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            // Bloques para MAR-VIE y LUN-VIE
            for (int d = 1; d <= 5; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario
                {
                    Id = d,
                    DiaSemanaNumero = d,
                    Franja = "MANANA",
                    DuracionHoras = 4,
                    Activo = 1
                });
            }

            // 6 materias: 12,12,8,8,12,12 = 64 sesiones > 61 ranuras
            int[] sesiones = { 12, 12, 8, 8, 12, 12 };
            string[] patrones = { "2,3", "2,3", "4,5", "4,5", "2,3", "4,5" };
            for (int i = 1; i <= 6; i++)
            {
                contexto.Asignaturas.Add(new Asignatura { Id = i, Horas = sesiones[i - 1] * 4 });
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = sesiones[i - 1] * 4,
                    DiasPatron = patrones[i - 1],
                    OrigenPatron = "LEGADO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }

            // 4 festivos en MAR-VIE y 2 en Lunes
            contexto.Festivos.Add((new DateTime(2027, 2, 9), new DateTime(2027, 2, 9)));
            contexto.Festivos.Add((new DateTime(2027, 3, 24), new DateTime(2027, 3, 24)));
            contexto.Festivos.Add((new DateTime(2027, 4, 1), new DateTime(2027, 4, 1)));
            contexto.Festivos.Add((new DateTime(2027, 4, 2), new DateTime(2027, 4, 2)));
            contexto.Festivos.Add((new DateTime(2027, 3, 22), new DateTime(2027, 3, 22)));
            contexto.Festivos.Add((new DateTime(2027, 5, 17), new DateTime(2027, 5, 17)));
            contexto.ConstruirIndices();

            // Act
            var res = _sut.ValidarFactibilidad(contexto);

            // Assert
            res.IsSuccess.Should().BeTrue();
            string bloqueos = string.Join(" | ", res.Value.BloqueosCriticos);
            bloqueos.Should().Contain("[I5_INFACTIBLE]");
            bloqueos.Should().Contain("no reduce las sesiones pendientes");
            bloqueos.Should().NotContain("Escale la ventana a 'LUN-VIE' (capacidad 75)");
        }

        [Fact]
        public void ValidarFactibilidad_WhenEscalonResuelveCompletamenteConAutomaticas_ShouldSugerirEscalonSimple()
        {
            // Arrange (Fase 2117 / Fase 2118)
            var contexto = new AgContexto();
            var inicio = new DateTime(2027, 2, 5);
            var fin = new DateTime(2027, 5, 30);
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", inicio, fin, AgOrigenVentana.Presencial).Value;

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "3|1|1|I|A",
                DiasVentana = "MAR-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Capacidad = 30,
                IdPrograma = 1
            };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            for (int d = 1; d <= 5; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario
                {
                    Id = d,
                    DiaSemanaNumero = d,
                    Franja = "MANANA",
                    DuracionHoras = 4,
                    Activo = 1
                });
            }

            int[] sesiones = { 12, 12, 8, 8, 12, 12 };
            for (int i = 1; i <= 6; i++)
            {
                contexto.Asignaturas.Add(new Asignatura { Id = i, Horas = sesiones[i - 1] * 4 });
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = sesiones[i - 1] * 4,
                    DiasPatron = string.Empty,
                    OrigenPatron = "AUTOMATICO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }

            contexto.Festivos.Add((new DateTime(2027, 2, 9), new DateTime(2027, 2, 9)));
            contexto.Festivos.Add((new DateTime(2027, 3, 24), new DateTime(2027, 3, 24)));
            contexto.Festivos.Add((new DateTime(2027, 4, 1), new DateTime(2027, 4, 1)));
            contexto.Festivos.Add((new DateTime(2027, 4, 2), new DateTime(2027, 4, 2)));
            contexto.Festivos.Add((new DateTime(2027, 3, 22), new DateTime(2027, 3, 22)));
            contexto.Festivos.Add((new DateTime(2027, 5, 17), new DateTime(2027, 5, 17)));
            contexto.ConstruirIndices();

            // Act
            var res = _sut.ValidarFactibilidad(contexto);

            // Assert
            res.IsSuccess.Should().BeTrue();
            string bloqueos = string.Join(" | ", res.Value.BloqueosCriticos);
            bloqueos.Should().Contain("[I5_INFACTIBLE]");
            bloqueos.Should().Contain("Escale la ventana a 'LUN-VIE' (capacidad 75) o amplíe el semestre.");
        }

        [Fact]
        public void ValidarFactibilidad_WhenEscalonReducePeroNoEliminaPendientes_ShouldDeclararResidual()
        {
            // Arrange (Fase 2123 / Fase 2125)
            var contexto = new AgContexto();
            var inicio = new DateTime(2027, 2, 5);
            var fin = new DateTime(2027, 5, 30);
            contexto.Ventana = VentanaSemestre.Crear(1, "2027-1", inicio, fin, AgOrigenVentana.Presencial).Value;

            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "3|1|1|I|A",
                DiasVentana = "MAR-VIE",
                Jornada = "MANANA",
                HorasDia = 4,
                Capacidad = 30,
                IdPrograma = 1
            };
            contexto.Aulas.Add(aula);
            contexto.AulasPorId[1] = aula;

            for (int d = 1; d <= 5; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario
                {
                    Id = d,
                    DiaSemanaNumero = d,
                    Franja = "MANANA",
                    DuracionHoras = 4,
                    Activo = 1
                });
            }

            // 4 materias rígidas (16,16 en 2,3 -> 32 sesiones > 31 capacidad carril) + (8,8 en 4,5 -> 16 sesiones)
            // + 1 AUTOMATICA (17 sesiones) => Demanda total 65 > 61 capacidad MAR-VIE (pendienteBase = 4)
            int[] sesiones = { 16, 16, 8, 8, 17 };
            string[] patrones = { "2,3", "2,3", "4,5", "4,5", "" };
            for (int i = 1; i <= 5; i++)
            {
                contexto.Asignaturas.Add(new Asignatura { Id = i, Horas = sesiones[i - 1] * 4 });
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = sesiones[i - 1] * 4,
                    DiasPatron = patrones[i - 1],
                    OrigenPatron = string.IsNullOrEmpty(patrones[i - 1]) ? "AUTOMATICO" : "LEGADO"
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }

            contexto.Festivos.Add((new DateTime(2027, 2, 9), new DateTime(2027, 2, 9)));
            contexto.Festivos.Add((new DateTime(2027, 3, 24), new DateTime(2027, 3, 24)));
            contexto.Festivos.Add((new DateTime(2027, 4, 1), new DateTime(2027, 4, 1)));
            contexto.Festivos.Add((new DateTime(2027, 4, 2), new DateTime(2027, 4, 2)));
            contexto.Festivos.Add((new DateTime(2027, 3, 22), new DateTime(2027, 3, 22)));
            contexto.Festivos.Add((new DateTime(2027, 5, 17), new DateTime(2027, 5, 17)));
            contexto.ConstruirIndices();

            // Act
            var res = _sut.ValidarFactibilidad(contexto);

            // Assert
            res.IsSuccess.Should().BeTrue();
            string bloqueos = string.Join(" | ", res.Value.BloqueosCriticos);
            bloqueos.Should().Contain("[I5_INFACTIBLE]");
            bloqueos.Should().Contain("reduce las sesiones pendientes de 4 a 2");
            bloqueos.Should().Contain("pero no las elimina");
            bloqueos.Should().NotContain("no reduce las sesiones pendientes");
        }

        [Theory]
        [InlineData(AgResultadoEscaladoUI.RechazadoPorUsuario, true)]
        [InlineData(AgResultadoEscaladoUI.SinCambios, false)]
        [InlineData(AgResultadoEscaladoUI.AplicadoExitosamente, false)]
        [InlineData(AgResultadoEscaladoUI.FalloAlAplicar, false)]
        public void PreVuelo_Override_ShouldDependerDelConsentimiento(AgResultadoEscaladoUI resultado, bool esperado)
            => AgPreVueloEscalado.ContinuarPeseAInfactibilidad(resultado).Should().Be(esperado);

        [Fact]
        public void ValidadorYEscalador_MismaParticion_ParaPatronesMarVie()
        {
            // Arrange: aula MAR-VIE con dos materias con patrones 2,3 y 4,5
            var contexto = new AgContexto();
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 6, 1), AgOrigenVentana.Presencial).Value;
            var aula = new AgAula
            {
                Id = 10,
                AulaRef = "A10",
                Subunidad = "",
                DiasVentana = "MAR-VIE",
                Jornada = "NOCHE",
                HorasDia = 4,
                Capacidad = 30,
                Estado = AgEstadoAula.Disponible
            };
            contexto.Aulas.Add(aula);

            for (int d = 2; d <= 5; d++)
            {
                contexto.Bloques.Add(new AgBloqueHorario
                {
                    Id = d,
                    DiaSemanaNumero = d,
                    Franja = "NOCHE",
                    DuracionHoras = 4,
                    Activo = 1
                });
            }

            contexto.Asignaturas.Add(new Asignatura { Id = 101, Nombre = "Materia A", Horas = 32 });
            contexto.Asignaturas.Add(new Asignatura { Id = 102, Nombre = "Materia B", Horas = 32 });

            contexto.AulaAsignaturaPorAulaAsig[(10, 101)] = new AgAulaAsignatura
            {
                IdAula = 10,
                IdAsignatura = 101,
                HorasObjetivo = 32,
                DiasPatron = "2,3",
                OrigenPatron = "MANUAL"
            };
            contexto.AulaAsignaturaPorAulaAsig[(10, 102)] = new AgAulaAsignatura
            {
                IdAula = 10,
                IdAsignatura = 102,
                HorasObjetivo = 32,
                DiasPatron = "4,5",
                OrigenPatron = "MANUAL"
            };
            contexto.AulaAsignaturas = new List<AgAulaAsignatura>(contexto.AulaAsignaturaPorAulaAsig.Values);
            contexto.ConstruirIndices();

            // Act
            var resValidador = _sut.ValidarFactibilidad(contexto);

            // Assert
            resValidador.IsSuccess.Should().BeTrue();
            resValidador.Value.BloqueosCriticos.Should().NotContain(b => b.Contains("I5_DIA_SATURADO"));
        }
    }
}

