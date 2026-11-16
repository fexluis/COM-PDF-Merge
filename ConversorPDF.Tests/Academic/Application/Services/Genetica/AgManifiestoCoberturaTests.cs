using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgManifiestoCoberturaTests
    {
        internal static AgContexto Contexto()
        {
            var c = new AgContexto();
            foreach (int id in new[] { 13, 14 })
            {
                c.Aulas.Add(new AgAula { Id = id, IdPeriodo = 3, IdSede = 1, IdPrograma = 1,
                    Semestre = "1", Grupo = id == 13 ? "A" : "B", Subunidad = "VALLE" });
                c.AulaAsignaturas.Add(new AgAulaAsignatura { IdAula = id, IdAsignatura = 7,
                    HorasObjetivo = 65, Requerida = 1, Subunidad = "VALLE" });
            }
            c.Asignaturas.Add(new Asignatura { Id = 7, IdPrograma = 1, Semestre = "I", IntHoraria = 64 });
            return c;
        }

        internal static AgManifiestoCobertura Crear(AgContexto c)
            => new AgConstructorManifiestoCobertura().Construir(c, new[] { 13, 14 }, 16, "VALLE",
                new DateTime(2027, 2, 5), new DateTime(2027, 5, 30)).Value;

        [Fact]
        public void OmisionNoDependeDeGenesNiIndicesYConservaDemandaOriginal()
        {
            var c = Contexto(); // No ConstruirIndices, genes ni staging.
            var m = Crear(c);
            Assert.Equal(2, m.Relaciones.Count);
            Assert.All(m.Relaciones, r => { Assert.Equal(65, r.HorasOriginales); Assert.Equal(68, r.HorasEfectivas); });
            Assert.Equal(14, Assert.Single(m.DetectarOmisiones(new[] { m.Relaciones[0].Clave })).IdAula);
            Assert.NotEqual(m.Relaciones[0].Clave.Valor, m.Relaciones[1].Clave.Valor);
            Assert.Null(c.ManifiestoCobertura);
        }

        [Fact]
        public void IdentidadIncluyeTodasLasDimensionesYHuellaEsIndependienteDelOrden()
        {
            var c = Contexto();
            var m = Crear(c);
            c.Aulas.Reverse(); c.AulaAsignaturas.Reverse();
            Assert.Equal(m.Huella, Crear(c).Huella);
            var k = m.Relaciones[0].Clave;
            var variantes = new[] {
                new AgClaveRelacionCobertura(17, k.Subunidad, 3, 1, 1, 13, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "CAUCA", 3, 1, 1, 13, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 4, 1, 1, 13, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 2, 1, 13, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 2, 13, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, 14, 7, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, 13, 8, "I", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, 13, 7, "II", "A", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, 13, 7, "I", "B", "Presencial"),
                new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, 13, 7, "I", "A", "Mixta") };
            Assert.All(variantes, v => Assert.NotEqual(k.Valor, v.Valor));
            c.AulaAsignaturas[0].HorasObjetivo = 66; // Mismo redondeo, distinto origen.
            Assert.NotEqual(m.Huella, Crear(c).Huella);
        }

        [Theory]
        [InlineData("seleccion")]
        [InlineData("subunidad")]
        [InlineData("relacion")]
        [InlineData("duplicada")]
        [InlineData("estado")]
        public void EntradaIncompletaOAmbiguaFallaSinExcluirAulas(string caso)
        {
            var c = Contexto();
            if (caso == "seleccion") c.Aulas.RemoveAt(1);
            if (caso == "subunidad") c.AulaAsignaturas[1].Subunidad = "CAUCA";
            if (caso == "relacion") c.AulaAsignaturas.RemoveAt(1);
            if (caso == "duplicada") c.AulaAsignaturas.Add(c.AulaAsignaturas[0]);
            if (caso == "estado") c.Aulas[1].Estado = AgEstadoAula.Inactiva;
            var r = new AgConstructorManifiestoCobertura().Construir(c, new[] { 13, 14 }, 16, "VALLE",
                new DateTime(2027, 2, 5), new DateTime(2027, 5, 30));
            Assert.True(r.IsFailure);
            Assert.Null(c.ManifiestoCobertura);
            Assert.Equal(65, c.AulaAsignaturas[0].HorasObjetivo);
        }

        [Fact]
        public void DisponibleMantienePrioridadActualDeHoras()
        {
            var c = Contexto();
            c.AulaAsignaturas.ForEach(r => r.HorasObjetivo = 0);
            c.Asignaturas[0].IntHoraria = 63;
            var m = Crear(c);
            Assert.All(m.Relaciones, r => { Assert.Equal(63, r.HorasOriginales); Assert.Equal(64, r.HorasEfectivas); });
        }
    }
}
