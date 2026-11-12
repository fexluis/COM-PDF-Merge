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
    public class AgGateG1ParticionCarrilesTests
    {
        [Fact]
        public void GateG1_CasoHomogeneo67_ConAgFabricaIndividuosYValidador_NoDebeGenerarR8R5NiColisiones()
        {
            // Contexto completo representativo con aula JUE-DOM (4,5,6,7), 3 materias de 32h sembradas con 6,7
            var aula = new AgAula
            {
                Id = 10,
                IdPrograma = 1,
                AulaRef = "SEDE|PREG|1|I|G1|PRES",
                Semestre = "I",
                DiasVentana = "4,5,6,7",
                Jornada = "NOCHE",
                Capacidad = 40,
                Estado = AgEstadoAula.Disponible
            };

            var fInicio = new DateTime(2027, 2, 5);
            var fFin = new DateTime(2027, 5, 30);

            var contexto = new AgContexto
            {
                IdPrograma = 1,
                Subunidad = "FACULTAD_ADMON",
                Ventana = VentanaSemestre.Crear(1, "2027-1", fInicio, fFin, AgOrigenVentana.Presencial).Value
            };
            contexto.Aulas.Add(aula);

            var asig38 = new Asignatura { Id = 38, IdPrograma = 1, Semestre = "I", Nombre = "Materia 38", Horas = 32, HorasSemanales = 4 };
            var asig40 = new Asignatura { Id = 40, IdPrograma = 1, Semestre = "I", Nombre = "Materia 40", Horas = 32, HorasSemanales = 4 };
            var asig42 = new Asignatura { Id = 42, IdPrograma = 1, Semestre = "I", Nombre = "Materia 42", Horas = 32, HorasSemanales = 4 };
            contexto.Asignaturas.Add(asig38);
            contexto.Asignaturas.Add(asig40);
            contexto.Asignaturas.Add(asig42);

            // F1593 / F1596: Configurar relaciones AgAulaAsignatura explícitas para alimentar ObtenerDiasPatron
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = aula.Id, IdAsignatura = 38, HorasObjetivo = 32, OrdenCurricular = 1, DiasPatron = "6,7" });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = aula.Id, IdAsignatura = 40, HorasObjetivo = 32, OrdenCurricular = 2, DiasPatron = "6,7" });
            contexto.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = aula.Id, IdAsignatura = 42, HorasObjetivo = 32, OrdenCurricular = 3, DiasPatron = "6,7" });

            // F1596: Bloques válidos para días 4, 5 nocturnos y sábado/domingo diurnos (MANANA/TARDE)
            contexto.Bloques.Add(new AgBloqueHorario { Id = 41, DiaSemanaNumero = 4, DuracionHoras = 4, Activo = 1, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 51, DiaSemanaNumero = 5, DuracionHoras = 4, Activo = 1, Franja = "NOCHE", HoraInicio = "18:00", HoraFin = "22:00" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 61, DiaSemanaNumero = 6, DuracionHoras = 4, Activo = 1, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 62, DiaSemanaNumero = 6, DuracionHoras = 4, Activo = 1, Franja = "TARDE", HoraInicio = "13:00", HoraFin = "17:00" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 71, DiaSemanaNumero = 7, DuracionHoras = 4, Activo = 1, Franja = "MANANA", HoraInicio = "08:00", HoraFin = "12:00" });

            contexto.Docentes.Add(new Tercero { Id = 101, Nombre = "Docente 1", Activo = 1 });
            contexto.Docentes.Add(new Tercero { Id = 102, Nombre = "Docente 2", Activo = 1 });
            contexto.Docentes.Add(new Tercero { Id = 103, Nombre = "Docente 3", Activo = 1 });

            // F1597 / F1599: Registrar habilitaciones mediante API canónica RegistrarDocenteHabilitado
            var ids = new[] { 38, 40, 42 };
            for (int i = 0; i < ids.Length; i++)
            {
                contexto.RegistrarDocenteHabilitado(101 + i, 1, ids[i], "I");
            }

            // F1596: Reconstruir índices antes de crear individuo
            contexto.ConstruirIndices();

            for (int i = 0; i < ids.Length; i++)
            {
                contexto.EsDocenteHabilitado(101 + i, 1, ids[i], "I").Should().BeTrue();
            }

            contexto.AulaAsignaturaPorAulaAsig.Should().ContainKey((10, 38));
            contexto.AulaAsignaturaPorAulaAsig[(10, 38)].DiasPatron.Should().Be("6,7");
            contexto.BloquesPorId.Should().ContainKey(62);

            // Instanciar AgFabricaIndividuos real en memoria
            var fabrica = new AgFabricaIndividuos(new Random(42));
            var individuo = fabrica.CrearAleatorio(contexto, 1, 0);

            individuo.Should().NotBeNull();
            individuo.Genes.Should().NotBeEmpty();

            // F1593: Cobertura de las tres asignaturas
            individuo.Genes.Select(g => g.IdAsignatura).Distinct()
                .Should().BeEquivalentTo(new[] { 38, 40, 42 });

            // F1597: Verificar que los genes correspondan a asignaciones válidas
            individuo.Genes.Should().OnlyContain(g =>
                g.EsVacante == 0 &&
                g.IdTercero > 0 &&
                g.IdBloqueHorario > 0);

            // Validar fidelidad del patrón sembrado tanto en DiasPatron como en DiaSemana
            foreach (var gen in individuo.Genes)
            {
                var sembrado = AgCarriles.Ventana(contexto.AulaAsignaturaPorAulaAsig[(aula.Id, gen.IdAsignatura)].DiasPatron);

                AgCarriles.Ventana(gen.DiasPatron)
                    .Should().BeEquivalentTo(sembrado, "el gen no debe ampliar el patrón sembrado");

                sembrado.Should().Contain(gen.DiaSemana, "el día asignado a la sesión debe estar dentro del patrón sembrado");
            }

            // F1595: Validador con API real Evaluar(...) y comprobación de violaciones duras
            var validador = new AgValidadorReglas();
            validador.Evaluar(individuo, contexto);

            individuo.Violaciones
                .Where(v => v.EsDura)
                .Should().NotContain(v =>
                    v.Regla.ToString().StartsWith("R8_", StringComparison.Ordinal) ||
                    v.Regla.ToString().StartsWith("R5_", StringComparison.Ordinal) ||
                    v.Regla.ToString().StartsWith("R13_", StringComparison.Ordinal),
                    "la partición fiel no debe violar R8, R5 ni R13");
        }
    }
}
