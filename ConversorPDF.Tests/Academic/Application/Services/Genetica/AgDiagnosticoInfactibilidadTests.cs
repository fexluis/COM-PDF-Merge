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
    public class AgDiagnosticoInfactibilidadTests
    {
        private const string DetalleCruceEntreCohortes = @"{
                ""indicesGenes"": [0, 1],
                ""fecha"": ""2026-02-03"",
                ""intervalos"": [
                    { ""IdAula"": 42, ""IdAsignatura"": 111, ""IdTercero"": 1043, ""inicio"": ""07:00"", ""fin"": ""11:00"" },
                    { ""IdAula"": 43, ""IdAsignatura"": 114, ""IdTercero"": 1043, ""inicio"": ""07:00"", ""fin"": ""11:00"" }
                ]
            }";

        private const string DetalleOcupacionExterna = @"{
                ""indicesGenes"": [0],
                ""fecha"": ""2026-02-03"",
                ""intervalos"": [
                    { ""IdAula"": 42, ""IdAsignatura"": 111, ""IdTercero"": 1043, ""inicio"": ""07:00"", ""fin"": ""11:00"" }
                ]
            }";

        [Fact]
        public void Describir_SinViolaciones_DevuelveVacio()
        {
            var ind = new AgIndividuo();
            var ctx = new AgContexto();

            var res = AgDiagnosticoInfactibilidad.Describir(ind, ctx);

            res.Should().BeEmpty();
        }

        [Fact]
        public void Describir_ConViolacionesPeroHardCero_DevuelveVacio()
        {
            var ind = new AgIndividuo();
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "Cruce R1", detalleJson: DetalleCruceEntreCohortes));
            ind.HardViolations = 0;

            var res = AgDiagnosticoInfactibilidad.Describir(ind, new AgContexto());

            res.Should().BeEmpty();
        }

        [Fact]
        public void Describir_ConCruceDocenteEntreCohortes_ReportaDocenteAsignaturasAulasYNoDuplicaPares()
        {
            var ind = new AgIndividuo();
            ind.Genes = new List<AgGen>
            {
                new AgGen { IdAula = 42, IdAsignatura = 111, IdTercero = 1043, IdPrograma = 7, Semestre = "II", Grupo = "A", EsVacante = 0 },
                new AgGen { IdAula = 43, IdAsignatura = 114, IdTercero = 1043, IdPrograma = 7, Semestre = "II", Grupo = "B", EsVacante = 0 }
            };
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "Cruce R1", detalleJson: DetalleCruceEntreCohortes));
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, true, "Cruce R7", detalleJson: DetalleCruceEntreCohortes));
            ind.HardViolations = 2;

            var ctx = new AgContexto();
            ctx.Asignaturas = new List<Asignatura>
            {
                new Asignatura { Id = 111, Nombre = "ESCUELAS FILOSOFICAS II" },
                new Asignatura { Id = 114, Nombre = "FUNDAMENTOS EN CIENCIAS SOCIALES" }
            };
            ctx.ConstruirIndices();

            var res = AgDiagnosticoInfactibilidad.Describir(ind, ctx);

            res.Should().NotBeEmpty();
            res[0].Should().Contain("paresDocente=1");
            res[0].Should().Contain("R1=1, R7=1");
            string lineasTexto = string.Join("\n", res);
            lineasTexto.Should().Contain("Docente 1043: pares=1");
            lineasTexto.Should().Contain("111 (ESCUELAS FILOSOFICAS II)");
            lineasTexto.Should().Contain("poolHabilitado=");
        }

        [Fact]
        public void Describir_ConR2SinTercero_NoCuentaComoParDocente()
        {
            var ind = new AgIndividuo();
            ind.Genes = new List<AgGen>
            {
                new AgGen { IdAula = 42, IdAsignatura = 111, IdTercero = 1043, IdPrograma = 7, Semestre = "II", Grupo = "A", EsVacante = 0 },
                new AgGen { IdAula = 42, IdAsignatura = 112, IdTercero = 1044, IdPrograma = 7, Semestre = "II", Grupo = "B", EsVacante = 0 }
            };
            string detalleR2 = @"{
                ""indicesGenes"": [0, 1],
                ""fecha"": ""2026-02-04"",
                ""intervalos"": [
                    { ""IdAula"": 42, ""IdAsignatura"": 111, ""IdTercero"": 1043, ""inicio"": ""07:00"", ""fin"": ""11:00"" },
                    { ""IdAula"": 42, ""IdAsignatura"": 112, ""IdTercero"": 1044, ""inicio"": ""07:00"", ""fin"": ""11:00"" }
                ]
            }";
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R2_SuperposicionAula, true, "Cruce R2", detalleJson: detalleR2));
            ind.HardViolations = 1;

            var res = AgDiagnosticoInfactibilidad.Describir(ind, new AgContexto());

            res[0].Should().Contain("paresDocente=0");
            string.Join("\n", res).Should().NotContain("Docente 0");
        }

        [Fact]
        public void Describir_ConR7OcupacionExterna_SeparaOrigen()
        {
            var ind = new AgIndividuo();
            ind.Genes = new List<AgGen>
            {
                new AgGen { IdAula = 42, IdAsignatura = 111, IdTercero = 1043, IdPrograma = 7, Semestre = "II", Grupo = "A", EsVacante = 0 }
            };
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, true, "Ocupaci\u00f3n externa del docente 1043", detalleJson: DetalleOcupacionExterna));
            ind.HardViolations = 1;

            var res = AgDiagnosticoInfactibilidad.Describir(ind, new AgContexto());

            res.Should().NotBeEmpty();
            res[0].Should().Contain("ocupacionExterna=1");
            res[0].Should().Contain("entreGenes=0");
        }

        [Fact]
        public void Describir_ConVacanteSinPool_ClasificaPoolVacio()
        {
            var ind = new AgIndividuo();
            ind.Genes = new List<AgGen>
            {
                new AgGen { IdAula = 42, IdAsignatura = 111, IdPrograma = 1, Semestre = "II", Grupo = "A", IdBloqueHorario = 10, EsVacante = 1, MotivoVacante = "SIN_DISPONIBILIDAD" }
            };
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "Violacion dura dummy"));
            ind.HardViolations = 1;

            var ctx = new AgContexto();
            ctx.PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>();

            var res = AgDiagnosticoInfactibilidad.Describir(ind, ctx);

            string texto = string.Join("\n", res);
            texto.Should().Contain("POOL_VACIO=1");
            texto.Should().Contain("motivoRegistrado=[SIN_DISPONIBILIDAD=1]");
        }

        [Fact]
        public void Describir_ConVacanteSinPlanOleada_ClasificaMotivoYNoBloqueNivel()
        {
            var ind = new AgIndividuo();
            ind.Genes = new List<AgGen>
            {
                new AgGen { IdAula = 47, IdAsignatura = 120, IdPrograma = 1, Semestre = "III", Grupo = "A", IdBloqueHorario = 0, EsVacante = 1, MotivoVacante = "SIN_PLAN_OLEADA" }
            };
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "Violacion dura dummy"));
            ind.HardViolations = 1;

            var res = AgDiagnosticoInfactibilidad.Describir(ind, new AgContexto());

            string texto = string.Join("\n", res);
            texto.Should().Contain("SIN_PLAN_OLEADA=1");
            texto.Should().Contain("SIN_BLOQUE_NIVEL=0");
        }

        [Fact]
        public void ResumenCorto_ConR15_ReportaReglaDura()
        {
            // Arrange (Fase 2088): Campeón con violación R15_VacanteExcedida
            var ind = new AgIndividuo();
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R15_VacanteExcedida, true, "Vacantes liberadas por el AG (15) exceden el déficit estructural (0)."));
            ind.HardViolations = 1;

            // Act
            string resumen = AgDiagnosticoInfactibilidad.ResumenCorto(ind, new AgContexto(), AgMotivoParada.Estancamiento, 1);

            // Assert
            var lineas = resumen.Split('\n');
            lineas.Length.Should().BeGreaterThanOrEqualTo(2);
            lineas[1].Should().Be("Reglas duras: [R15_VacanteExcedida x1]");
        }

        [Fact]
        public void Describir_ReportaReglasDurasNoDocentes()
        {
            // Arrange (Fase 2088): Campeón con violaciones duras que no involucran cruces de docentes (R15 y R16)
            var ind = new AgIndividuo();
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R15_VacanteExcedida, true, "Exceso vacantes"));
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R16_VacanteEstructuralNoConsentida, true, "Vacante sin consentimiento"));
            ind.HardViolations = 2;

            // Act
            var res = AgDiagnosticoInfactibilidad.Describir(ind, new AgContexto());

            // Assert
            res.Should().NotBeEmpty();
            res[0].Should().Contain("ReglasDuras=[R15_VacanteExcedida x1, R16_VacanteEstructuralNoConsentida x1]");
            res[0].Should().Contain("paresDocente=0");
        }

        [Fact]
        public void Describir_ParticionVacantes_SumaTotalVacantes()
        {
            // Arrange (Fase 2099): 5 genes vacantes cubriendo cada categoría de la partición exacta
            var contexto = new AgContexto();
            contexto.AsignaturasVacantesConsentidas.Add(10); // Asig 10 consentida

            var ind = new AgIndividuo();
            ind.Violaciones.Add(new AgViolacion(AgReglaViolacion.R15_VacanteExcedida, true, "Exceso vacantes"));
            ind.HardViolations = 1;
            ind.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdSede = 1, EsVacante = 1, MotivoVacante = "SIN_CANDIDATOS" });
            ind.Genes.Add(new AgGen { Id = 2, IdAsignatura = 20, IdSede = 1, EsVacante = 1, MotivoVacante = "SIN_DISPONIBILIDAD" });
            ind.Genes.Add(new AgGen { Id = 3, IdAsignatura = 30, IdSede = 1, EsVacante = 1, MotivoVacante = "SIN_CANDIDATOS" });
            ind.Genes.Add(new AgGen { Id = 4, IdAsignatura = 40, IdSede = 1, EsVacante = 1, MotivoVacante = "SIN_PLAN_OLEADA" });
            ind.Genes.Add(new AgGen { Id = 5, IdAsignatura = 50, IdSede = 1, EsVacante = 1, MotivoVacante = "" });

            // Act
            var res = AgDiagnosticoInfactibilidad.Describir(ind, contexto);

            // Assert
            string texto = string.Join("\n", res);
            texto.Should().Contain("Vacantes=5 (CONSENTIDAS=1, LIBERADAS_POR_AG=1, ESTRUCTURALES_NC=1, NO_DOCENTE=1, INDETERMINADAS=1)");
        }
    }
}