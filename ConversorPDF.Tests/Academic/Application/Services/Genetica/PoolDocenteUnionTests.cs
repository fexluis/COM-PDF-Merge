using System;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class PoolDocenteUnionTests
    {
        [Fact]
        public void EsDocenteHabilitado_WhenRegisteredWithArabic_ShouldMatchRomanQuery()
        {
            // Arrange
            var contexto = new AgContexto();
            // Docente registrado con semestre arábigo "2"
            contexto.RegistrarDocenteHabilitado(idTercero: 10, idPrograma: 1, idAsignatura: 100, semestre: "2");

            // Act & Assert
            // Búsqueda con "II" debe retornar true gracias a la normalización multi-semestre
            contexto.EsDocenteHabilitado(10, 1, 100, "II").Should().BeTrue();
            contexto.EsDocenteHabilitado(10, 1, 100, "2").Should().BeTrue();
        }

        [Fact]
        public void EsDocenteHabilitado_WhenRegisteredWithRoman_ShouldMatchArabicQuery()
        {
            // Arrange
            var contexto = new AgContexto();
            // Docente registrado con semestre romano "VIII"
            contexto.RegistrarDocenteHabilitado(idTercero: 20, idPrograma: 1, idAsignatura: 200, semestre: "VIII");

            // Act & Assert
            contexto.EsDocenteHabilitado(20, 1, 200, "8").Should().BeTrue();
            contexto.EsDocenteHabilitado(20, 1, 200, "VIII").Should().BeTrue();
        }

        [Fact]
        public void EsDocenteHabilitado_WhenRegisteredWithoutSemester_ShouldMatchAnySemester()
        {
            // Arrange
            var contexto = new AgContexto();
            // Docente general registrado sin semestre (comodín)
            contexto.RegistrarDocenteHabilitado(idTercero: 30, idPrograma: 1, idAsignatura: 300, semestre: "");

            // Act & Assert
            contexto.EsDocenteHabilitado(30, 1, 300, "V").Should().BeTrue();
            contexto.EsDocenteHabilitado(30, 1, 300, "5").Should().BeTrue();
            contexto.EsDocenteHabilitado(30, 1, 300, "X").Should().BeTrue();
        }
    }
}
