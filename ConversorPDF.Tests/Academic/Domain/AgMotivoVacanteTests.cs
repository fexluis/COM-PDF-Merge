using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgMotivoVacanteTests
    {
        [Fact]
        public void ClasificarMotivo_WhenPoolIsNullAndNoDocentes_ShouldReturnPoolVacio()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Docentes = new List<Tercero>();
            contexto.PoolPorAsignatura = null;
            var asig = new Asignatura { Id = 10, Nombre = "Cálculo I" };

            // Act
            var motivo = AgFabricaIndividuos.ClasificarMotivo(contexto, asig);

            // Assert
            motivo.Should().Be(AgMotivoVacante.POOL_VACIO);
        }

        [Fact]
        public void ClasificarMotivo_WhenPoolIsNullButDocentesExist_ShouldReturnSinCandidatos()
        {
            // Arrange
            var contexto = new AgContexto();
            contexto.Docentes = new List<Tercero> { new Tercero { Id = 1, Nombre = "Docente A", Activo = 1 } };
            contexto.PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>();
            var asig = new Asignatura { Id = 10, Nombre = "Cálculo I" };

            // Act
            var motivo = AgFabricaIndividuos.ClasificarMotivo(contexto, asig);

            // Assert
            motivo.Should().Be(AgMotivoVacante.SIN_CANDIDATOS);
        }

        [Fact]
        public void ClasificarMotivo_ConPoolSoloCentinela_DevuelveSinCandidatos()
        {
            // Arrange
            var contexto = new AgContexto
            {
                IdTerceroVacante = 999
            };
            var pool = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 999, IdAsignatura = 10, Activo = 1 }
            };
            contexto.PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>
            {
                { 10, pool }
            };
            var asig = new Asignatura { Id = 10, Nombre = "Cálculo I" };

            // Act
            var motivo = AgFabricaIndividuos.ClasificarMotivo(contexto, asig);

            // Assert
            motivo.Should().Be(AgMotivoVacante.SIN_CANDIDATOS);
        }

        [Fact]
        public void ClasificarMotivo_WhenPoolHasRowsButAllInactive_ShouldReturnNoCumpleRequisitos()
        {
            // Arrange
            var contexto = new AgContexto();
            var pool = new List<TerceroAsignatura>
            {
                new TerceroAsignatura { IdTercero = 1, IdAsignatura = 10, Activo = 0 }
            };
            contexto.PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>
            {
                { 10, pool }
            };
            var asig = new Asignatura { Id = 10, Nombre = "Cálculo I" };

            // Act
            var motivo = AgFabricaIndividuos.ClasificarMotivo(contexto, asig);

            // Assert
            motivo.Should().Be(AgMotivoVacante.NO_CUMPLE_REQUISITOS);
        }

        [Fact]
        public void ClasificarMotivo_WhenPoolActiveTeachersExistAndHabilitados_ShouldReturnSinDisponibilidad()
        {
            // Arrange
            var contexto = new AgContexto();
            var ta = new TerceroAsignatura { IdTercero = 101, IdPrograma = 1, IdAsignatura = 20, Semestre = "II", Activo = 1 };
            ta.Id = TerceroAsignaturaKey.Calcular(ta.IdTercero, ta.IdPrograma, ta.IdAsignatura, ta.Semestre);
            contexto.PoolDocentes.Add(ta);
            contexto.ConstruirIndices();

            var asig = new Asignatura { Id = 20, IdPrograma = 1, Semestre = "II", Nombre = "Física II" };

            // Act
            var motivo = AgFabricaIndividuos.ClasificarMotivo(contexto, asig);

            // Assert
            motivo.Should().Be(AgMotivoVacante.SIN_DISPONIBILIDAD);
        }

        [Fact]
        public void AgFabricaIndividuos_WhenNoDocenteHabilitado_ShouldCreateVacantGenWithoutThrowing()
        {
            // Arrange
            var fabrica = new AgFabricaIndividuos(new Random(123));
            var contexto = new AgContexto();
            contexto.IdPrograma = 1;
            contexto.Subunidad = "FACULTAD_MED";
            contexto.Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;
            contexto.Asignaturas.Add(new Asignatura { Id = 99, Nombre = "Anatomía I", Horas = 32, HorasSemanales = 2 });
            contexto.Aulas.Add(new AgAula { Id = 5, Estado = AgEstadoAula.Disponible, DiasVentana = "MAR-VIE" });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, DuracionHoras = 4, Activo = 1 });
            contexto.Bloques.Add(new AgBloqueHorario { Id = 2, DiaSemanaNumero = 3, DuracionHoras = 4, Activo = 1 });
            // Sin docentes habilitados ni en contexto
            contexto.Docentes = new List<Tercero>();
            contexto.PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>();

            // Act
            var individuo = fabrica.CrearAleatorio(contexto, 1, 0);

            // Assert
            individuo.Should().NotBeNull();
            individuo.Genes.Should().HaveCount(2);
            individuo.Genes.Should().OnlyContain(g => g.EsVacante == 1 && g.IdTercero == 0 && g.MotivoVacante == AgMotivoVacante.POOL_VACIO.Codigo());
            individuo.Vacantes.Should().Be(2);
        }

        [Theory]
        [InlineData(AgMotivoVacante.Ninguno, "")]
        [InlineData(AgMotivoVacante.SIN_CANDIDATOS, "SIN_CANDIDATOS")]
        [InlineData(AgMotivoVacante.SIN_DISPONIBILIDAD, "SIN_DISPONIBILIDAD")]
        [InlineData(AgMotivoVacante.NO_CUMPLE_REQUISITOS, "NO_CUMPLE_REQUISITOS")]
        [InlineData(AgMotivoVacante.POOL_VACIO, "POOL_VACIO")]
        [InlineData(AgMotivoVacante.FUERA_DE_VENTANA, "FUERA_DE_VENTANA")]
        [InlineData(AgMotivoVacante.SIN_PLAN_OLEADA, "SIN_PLAN_OLEADA")]
        [InlineData(AgMotivoVacante.SIN_BLOQUE_NIVEL, "SIN_BLOQUE_NIVEL")]
        public void Codigo_ShouldReturnDeterministicCanonicalString(AgMotivoVacante motivo, string esperado)
        {
            motivo.Codigo().Should().Be(esperado);
        }

        [Theory]
        [InlineData("SIN_CANDIDATOS", AgMotivoVacante.SIN_CANDIDATOS)]
        [InlineData("sin_candidatos", AgMotivoVacante.SIN_CANDIDATOS)]
        [InlineData("Sin_Candidatos", AgMotivoVacante.SIN_CANDIDATOS)]
        [InlineData("POOL_VACIO", AgMotivoVacante.POOL_VACIO)]
        [InlineData("FUERA_DE_VENTANA", AgMotivoVacante.FUERA_DE_VENTANA)]
        [InlineData("SIN_PLAN_OLEADA", AgMotivoVacante.SIN_PLAN_OLEADA)]
        [InlineData("sin_plan_oleada", AgMotivoVacante.SIN_PLAN_OLEADA)]
        [InlineData("SIN_BLOQUE_NIVEL", AgMotivoVacante.SIN_BLOQUE_NIVEL)]
        [InlineData("sin_bloque_nivel", AgMotivoVacante.SIN_BLOQUE_NIVEL)]
        [InlineData("", AgMotivoVacante.Ninguno)]
        [InlineData(null, AgMotivoVacante.Ninguno)]
        [InlineData("DESCONOCIDO", AgMotivoVacante.Ninguno)]
        public void Desde_ShouldParseSafelyCaseInsensitive(string codigo, AgMotivoVacante esperado)
        {
            AgMotivoVacanteExtensions.Desde(codigo).Should().Be(esperado);
        }

        [Theory]
        [InlineData("SIN_PLAN_OLEADA", AgOrigenVacante.NoDocente)]
        [InlineData("SIN_BLOQUE_NIVEL", AgOrigenVacante.NoDocente)]
        [InlineData("SIN_CANDIDATOS", AgOrigenVacante.EstructuralDocente)]
        [InlineData("NO_CUMPLE_REQUISITOS", AgOrigenVacante.EstructuralDocente)]
        [InlineData("POOL_VACIO", AgOrigenVacante.EstructuralDocente)]
        [InlineData("SIN_DISPONIBILIDAD", AgOrigenVacante.LiberadaAg)]
        [InlineData("FUERA_DE_VENTANA", AgOrigenVacante.LiberadaAg)]
        [InlineData("OTRO_MOTIVO", AgOrigenVacante.Indeterminado)]
        [InlineData("", AgOrigenVacante.Indeterminado)]
        [InlineData(null, AgOrigenVacante.Indeterminado)]
        [InlineData("DESCONOCIDO", AgOrigenVacante.Indeterminado)]
        public void Clasificar_ShouldCategorizeCorrectly(string motivo, AgOrigenVacante esperado)
        {
            AgClasificadorVacante.Clasificar(motivo).Should().Be(esperado);
        }

        [Fact]
        public void ClaveCohorte_WhenGenesIncompletos_ShouldDiscriminarPorAulaYGen()
        {
            // Arrange (Fase 2097): 3 genes sin Semestre ni Grupo pero con distinta aula/id
            var g1 = new AgGen { Id = 1, IdAsignatura = 10, IdSede = 1, IdAula = 101, Semestre = "", Grupo = "" };
            var g2 = new AgGen { Id = 2, IdAsignatura = 10, IdSede = 1, IdAula = 102, Semestre = null, Grupo = null };
            var g3 = new AgGen { Id = 3, IdAsignatura = 10, IdSede = 1, IdAula = 101, Semestre = " ", Grupo = " " };

            // Act
            string c1 = AgClasificadorVacante.ClaveCohorte(g1);
            string c2 = AgClasificadorVacante.ClaveCohorte(g2);
            string c3 = AgClasificadorVacante.ClaveCohorte(g3);

            // Assert: no deben colapsar en una sola cohorte 10|1||
            c1.Should().Be("10|1|aula:101|gen:1");
            c2.Should().Be("10|1|aula:102|gen:2");
            c3.Should().Be("10|1|aula:101|gen:3");

            var set = new HashSet<string> { c1, c2, c3 };
            set.Should().HaveCount(3);
        }
    }
}
