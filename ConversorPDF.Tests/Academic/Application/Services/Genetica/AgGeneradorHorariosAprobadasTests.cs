using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgGeneradorHorariosAprobadasTests
    {
        private static AgContexto CrearContextoBase(DateTime ini, DateTime fin)
        {
            var contexto = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = ini,
                    FechaFin = fin
                }
            };

            var prog = new Programa
            {
                Id = 1,
                Nombre = "Programa Test",
                Modalidad = "PREGRADO"
            };
            contexto.Programas.Add(prog);

            return contexto;
        }

        [Fact]
        public void ValidarAprobadas_WhenAulaAprobadaCabeEnPlazo_ShouldReturnSuccess()
        {
            // Arrange (17 semanas: VIE-SAB = 51 ranuras; 1 materia x 64h = 16 sesiones <= 51)
            var contexto = CrearContextoBase(new DateTime(2026, 2, 1), new DateTime(2026, 5, 30));
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-APROBADA-OK",
                Estado = AgEstadoAula.Aprobado,
                IdPrograma = 1,
                DiasVentana = "5,6",
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            var asig = new Asignatura { Id = 1, Nombre = "Materia 1", IntHoraria = 64, IdPrograma = 1 };
            contexto.Asignaturas.Add(asig);
            var aa = new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 1,
                HorasObjetivo = 64.0
            };
            contexto.AulaAsignaturas.Add(aa);
            contexto.AulaAsignaturaPorAulaAsig[(1, 1)] = aa;
            contexto.ConstruirIndices();

            // Act
            var result = AgGeneradorHorarios.ValidarAprobadas(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void ValidarAprobadas_WhenAulaAprobadaTieneDeficit_ShouldReturnFailureConI5()
        {
            // Arrange (17 semanas: VIE-SAB = 51 ranuras; 4 materias x 64h = 64 sesiones > 51)
            var contexto = CrearContextoBase(new DateTime(2026, 2, 1), new DateTime(2026, 5, 30));
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-APROBADA-DEFICIT",
                Estado = AgEstadoAula.Aprobado,
                IdPrograma = 1,
                DiasVentana = "5,6",
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            for (int i = 1; i <= 4; i++)
            {
                var asig = new Asignatura { Id = i, Nombre = "Materia " + i, IntHoraria = 64, IdPrograma = 1 };
                contexto.Asignaturas.Add(asig);
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = 64.0
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }
            contexto.ConstruirIndices();

            // Act
            var result = AgGeneradorHorarios.ValidarAprobadas(contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().StartWith("I5_AULA_APROBADA_CON_DEFICIT: AULA-APROBADA-DEFICIT; requiere revisión explícita de aprobación para replanificar. El escalador no modifica aulas aprobadas.");
        }

        [Fact]
        public void ValidarAprobadas_WhenAulaAprobadaYVentanaNull_ShouldReturnVentanaNoDefinida()
        {
            // Arrange
            var contexto = new AgContexto { Ventana = null };
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-01",
                Estado = AgEstadoAula.Aprobado,
                DiasVentana = "5,6",
                Jornada = "NOCHE"
            };
            contexto.Aulas.Add(aula);

            // Act
            var result = AgGeneradorHorarios.ValidarAprobadas(contexto);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("I1_VENTANA_NO_DEFINIDA");
        }

        [Fact]
        public void ValidarAprobadas_WhenSinAulasAprobadas_ShouldReturnSuccess()
        {
            // Arrange
            var contexto = new AgContexto { Ventana = null };
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-DISP",
                Estado = AgEstadoAula.Disponible
            };
            contexto.Aulas.Add(aula);

            // Act
            var result = AgGeneradorHorarios.ValidarAprobadas(contexto);

            // Assert
            result.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenAulaAprobadaNoCabe_ShouldAbortConI5AntesDePreVuelo()
        {
            // Arrange
            var mockVentana = new Mock<IAgValidadorVentana>();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();

            var contexto = CrearContextoBase(new DateTime(2026, 2, 1), new DateTime(2026, 5, 30));
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-DEFICIT-OPT",
                Estado = AgEstadoAula.Aprobado,
                IdPrograma = 1,
                DiasVentana = "5,6",
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            for (int i = 1; i <= 4; i++)
            {
                var asig = new Asignatura { Id = i, Nombre = "Materia " + i, IntHoraria = 64, IdPrograma = 1 };
                contexto.Asignaturas.Add(asig);
                var aa = new AgAulaAsignatura
                {
                    IdAula = 1,
                    IdAsignatura = i,
                    HorasObjetivo = 64.0
                };
                contexto.AulaAsignaturas.Add(aa);
                contexto.AulaAsignaturaPorAulaAsig[(1, i)] = aa;
            }
            contexto.ConstruirIndices();

            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);

            // Act
            var result = await sut.EjecutarOptimizacionAsync(contexto, 10, 5, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().StartWith("I5_AULA_APROBADA_CON_DEFICIT: AULA-DEFICIT-OPT; requiere revisión explícita de aprobación para replanificar. El escalador no modifica aulas aprobadas.");
            mockVentana.Verify(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [Fact]
        public async Task EjecutarOptimizacionAsync_WhenAulaAprobadaCabeEnPlazo_ShouldProceedToPreVuelo()
        {
            // Arrange
            var mockVentana = new Mock<IAgValidadorVentana>();
            var mockAptitud = new Mock<IAgEvaluadorAptitud>();

            var informe = new InformeFactibilidad();
            informe.AgregarBloqueo("Bloqueo simulado en pre-vuelo.");
            mockVentana.Setup(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(Result.Success(informe));

            var contexto = CrearContextoBase(new DateTime(2026, 2, 1), new DateTime(2026, 5, 30));
            var aula = new AgAula
            {
                Id = 1,
                AulaRef = "AULA-APROBADA-OK",
                Estado = AgEstadoAula.Aprobado,
                IdPrograma = 1,
                DiasVentana = "5,6",
                Jornada = "NOCHE",
                HorasDia = 4
            };
            contexto.Aulas.Add(aula);

            var asig = new Asignatura { Id = 1, Nombre = "Materia 1", IntHoraria = 64, IdPrograma = 1 };
            contexto.Asignaturas.Add(asig);
            var aa = new AgAulaAsignatura
            {
                IdAula = 1,
                IdAsignatura = 1,
                HorasObjetivo = 64.0
            };
            contexto.AulaAsignaturas.Add(aa);
            contexto.AulaAsignaturaPorAulaAsig[(1, 1)] = aa;
            contexto.ConstruirIndices();

            var sut = new AgGeneradorHorarios(mockVentana.Object, mockAptitud.Object);

            // Act
            var result = await sut.EjecutarOptimizacionAsync(contexto, 10, 5, 0.8, 0.1, 1, null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("bloqueos de factibilidad detectados");
            result.Error.Should().Contain("Bloqueo simulado en pre-vuelo.");
            mockVentana.Verify(v => v.ValidarFactibilidadAsync(It.IsAny<AgContexto>(), It.IsAny<CancellationToken>()), Times.Once);
        }
    }
}
