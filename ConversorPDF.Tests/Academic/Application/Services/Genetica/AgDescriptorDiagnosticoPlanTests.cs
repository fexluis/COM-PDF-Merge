using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    // Escenario productivo 2026-11-16 (aulas 12/13/14): el período no tiene plan y el terminal
    // debe nombrar el motivo real en lugar de una etiqueta genérica.
    public class AgDescriptorDiagnosticoPlanTests
    {
        private static AgClaveRelacionCobertura Clave(int asignatura = 7, int aula = 14)
            => new AgClaveRelacionCobertura(16, "VALLE", 3, 1, 1, aula, asignatura, "I", "B", "PRESENCIAL");

        private static AgRelacionCobertura Requerida(int asignatura = 7, int aula = 14)
            => new AgRelacionCobertura(Clave(asignatura, aula), 4, 4, (int)AgEstadoAula.Aprobado);

        private static AgValidacionRelacionPlan Ausente(int asignatura = 7, int aula = 14)
            => new AgValidacionRelacionPlan(Requerida(asignatura, aula), 0, 0,
                new[] { "RELACION_AUSENTE:" + Clave(asignatura, aula).Valor });

        private static AgValidacionAulaPlan Aula(params AgValidacionRelacionPlan[] relaciones)
            => new AgValidacionAulaPlan(14, AgEstadoValidacionPlan.PlanNoVerificable, relaciones);

        [Fact]
        public void TodasAusentes_InformaConteoYCodigoConcreto()
        {
            var texto = AgDescriptorDiagnosticoPlan.Describir(Aula(Ausente(1), Ausente(2), Ausente(3)));
            Assert.StartsWith("RELACIONES_AUSENTES=3/3", texto);
            Assert.Contains("RELACION_AUSENTE:", texto);
            Assert.DoesNotContain(AgDescriptorDiagnosticoPlan.SinDiagnostico, texto);
        }

        [Fact]
        public void SinDiagnosticoTipado_NoInventaCausa()
        {
            var completa = new AgValidacionRelacionPlan(Requerida(), 4, 0, new string[0]);
            var texto = AgDescriptorDiagnosticoPlan.Describir(Aula(completa));
            Assert.Equal("RELACIONES_AUSENTES=0/1;" + AgDescriptorDiagnosticoPlan.SinDiagnostico, texto);
        }

        [Fact]
        public void DiagnosticosGlobalesDelLector_SeIncluyenPorqueBloqueanElAula()
        {
            var completa = new AgValidacionRelacionPlan(Requerida(), 4, 0, new string[0]);
            var texto = AgDescriptorDiagnosticoPlan.Describir(Aula(completa),
                new[] { "VINCULACIONES_MULTIPLES:" + Clave().Valor });
            Assert.Contains("VINCULACIONES_MULTIPLES:", texto);
        }

        [Fact]
        public void ExcesoDeCodigos_SeResumeSinPerderElConteo()
        {
            var relaciones = Enumerable.Range(1, 5).Select(a => Ausente(a)).ToArray();
            var texto = AgDescriptorDiagnosticoPlan.Describir(Aula(relaciones), null, 2);
            Assert.StartsWith("RELACIONES_AUSENTES=5/5", texto);
            Assert.EndsWith("(+3)", texto);
        }

        [Fact]
        public void SalidaEsDeterministaSinImportarElOrdenDeEntrada()
        {
            var a = AgDescriptorDiagnosticoPlan.Describir(Aula(Ausente(1), Ausente(2)), new[] { "IDENTIDAD_AMBIGUA:1" });
            var b = AgDescriptorDiagnosticoPlan.Describir(Aula(Ausente(2), Ausente(1)), new[] { "IDENTIDAD_AMBIGUA:1" });
            Assert.Equal(a, b);
        }

        [Fact]
        public void AulaNula_EsErrorDeProgramacion()
            => Assert.Throws<ArgumentNullException>(() => AgDescriptorDiagnosticoPlan.Describir(null));
    }
}
