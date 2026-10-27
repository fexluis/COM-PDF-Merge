using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class AsignaturaViaticoAllocatorTests
    {
        public AsignaturaViaticoAllocatorTests()
        {
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        }

        [Fact]
        public void Ejemplo15Bloques4Sedes_AsignaPorVinculacion()
        {
            // Arrange: 4 asignaturas/sedes
            var asig1 = CrearAsignatura(1, "CALI", "", "ADMINISTRACION PUBLICA", "2026-08-01", "2026-12-15");
            var asig2 = CrearAsignatura(2, "CARTAGO", "", "ECONOMIA", "2026-08-01", "2026-12-15");
            var asig3 = CrearAsignatura(3, "EL AGUILA", "", "DERECHO", "2026-08-01", "2026-12-15");
            var asig4 = CrearAsignatura(4, "BUENAVENTURA", "", "GESTION", "2026-08-01", "2026-12-15");

            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig1, asig2, asig3, asig4 };

            // 15 bloques: 4 para Sede 1, 6 para Sede 2, 2 para Sede 3, 3 para Sede 4
            var bloques = new List<ViaticoBloque>();
            for (int i = 1; i <= 4; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 8, i * 2), 1, "CALI", ""));
            }
            for (int i = 1; i <= 6; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 9, i * 2), 2, "CARTAGO", ""));
            }
            for (int i = 1; i <= 2; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 10, i * 2), 3, "EL AGUILA", ""));
            }
            for (int i = 1; i <= 3; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 11, i * 2), 4, "BUENAVENTURA", ""));
            }

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques, avisos);

            // Assert: 4 / 6 / 2 / 3, total 15, sin continuaciones ni advertencias
            resultado.Should().HaveCount(4);
            resultado[0].Bloques.Should().HaveCount(4);
            resultado[1].Bloques.Should().HaveCount(6);
            resultado[2].Bloques.Should().HaveCount(2);
            resultado[3].Bloques.Should().HaveCount(3);
            resultado.Sum(f => f.Bloques.Count).Should().Be(15);
            resultado.Should().OnlyContain(f => !f.EsContinuacion);
            avisos.Should().BeEmpty();
        }

        [Fact]
        public void BloqueCruzaDosSedes_SeMantieneEnteroEnFilaConMasSesiones()
        {
            // Arrange: 2 asignaturas
            var asigA = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-12-15");
            var asigB = CrearAsignatura(20, "CARTAGO", "", "ASIG 2", "2026-08-01", "2026-12-15");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asigA, asigB };

            // Bloque con 2 sesiones en Cali (IdVinc=10) y 1 en Cartago (IdVinc=20)
            var bloqueMulti = new ViaticoBloque();
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 1, IdVinculacion = 10, SedeCiudad = "CALI", FechaSesion = new DateTime(2026, 9, 1), Viatica = true });
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 2, IdVinculacion = 10, SedeCiudad = "CALI", FechaSesion = new DateTime(2026, 9, 2), Viatica = true });
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 3, IdVinculacion = 20, SedeCiudad = "CARTAGO", FechaSesion = new DateTime(2026, 9, 3), Viatica = true });

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, new List<ViaticoBloque> { bloqueMulti }, avisos);

            // Assert: El bloque debe ir entero a Fila A (2 sesiones > 1 sesión), 1 sola comisión
            resultado[0].Bloques.Should().HaveCount(1);
            resultado[1].Bloques.Should().BeEmpty();
            avisos.Should().Contain(a => a.Contains("múltiples vinculaciones"));
        }

        [Fact]
        public void BloqueCruzaDosSedes_FilaMayoritariaLlena_UsaOtraVinculacionConCupo()
        {
            // Arrange: Fila A (IdVinc=10) con 8 bloques preexistentes, Fila B (IdVinc=20) con cupo
            var asigA = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-12-15");
            var asigB = CrearAsignatura(20, "CARTAGO", "", "ASIG 2", "2026-08-01", "2026-12-15");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asigA, asigB };

            var bloques = new List<ViaticoBloque>();
            for (int i = 1; i <= 8; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 8, i), 10, "CALI", ""));
            }

            // Bloque 9 cruza: 2 sesiones en Cali (IdVinc=10) y 1 en Cartago (IdVinc=20)
            var bloqueMulti = new ViaticoBloque();
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 101, IdVinculacion = 10, SedeCiudad = "CALI", FechaSesion = new DateTime(2026, 9, 1), Viatica = true });
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 102, IdVinculacion = 10, SedeCiudad = "CALI", FechaSesion = new DateTime(2026, 9, 2), Viatica = true });
            bloqueMulti.Agregar(new SesionListItem { IdSesion = 103, IdVinculacion = 20, SedeCiudad = "CARTAGO", FechaSesion = new DateTime(2026, 9, 3), Viatica = true });
            bloques.Add(bloqueMulti);

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques, avisos);

            // Assert: Fila A llena (8 bloques); bloque multi pasa entero a Fila B (1 bloque)
            resultado.Should().HaveCount(2);
            resultado[0].Bloques.Should().HaveCount(8);
            resultado[1].Bloques.Should().HaveCount(1);
            resultado[1].Bloques[0].Should().BeSameAs(bloqueMulti);
            resultado.Should().OnlyContain(f => !f.EsContinuacion);
        }

        [Fact]
        public void FilaNaturalLlena_ReasignaAOtraFilaConCupo()
        {
            // Arrange: Fila A (IdVinc=10) y Fila B (IdVinc=20). Ambas con ventana válida.
            var asigA = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-10-31");
            var asigB = CrearAsignatura(20, "CARTAGO", "", "ASIG 2", "2026-08-01", "2026-10-31");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asigA, asigB };

            // 9 bloques para IdVinc=10
            var bloques = new List<ViaticoBloque>();
            for (int i = 1; i <= 9; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 8, i * 2), 10, "CALI", ""));
            }

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques, avisos);

            // Assert: Fila A tiene 8 bloques; Bloque 9 se reasigna a Fila B por R4
            resultado.Should().HaveCount(2);
            resultado[0].Bloques.Should().HaveCount(8);
            resultado[1].Bloques.Should().HaveCount(1);
            resultado.Should().OnlyContain(f => !f.EsContinuacion);
            avisos.Should().Contain(a => a.Contains("saturación"));
        }

        [Fact]
        public void TodasLasFilasLlenas_CreaFilaDeContinuacion()
        {
            // Arrange: 1 sola asignatura con 9 bloques
            var asig = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-12-15");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig };

            var bloques = new List<ViaticoBloque>();
            for (int i = 1; i <= 9; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 8, i * 2), 10, "CALI", ""));
            }

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques, avisos);

            // Assert: Fila base con 8 bloques, fila de continuación con 1 bloque
            resultado.Should().HaveCount(2);
            resultado[0].EsContinuacion.Should().BeFalse();
            resultado[0].Bloques.Should().HaveCount(8);
            resultado[1].EsContinuacion.Should().BeTrue();
            resultado[1].Bloques.Should().HaveCount(1);
            avisos.Should().Contain(a => a.Contains("continuación"));
        }

        [Fact]
        public void NingunaFilaSuperaOchoComisiones()
        {
            // Arrange: 2 asignaturas y 20 bloques (exige filas de continuación)
            var asig1 = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-12-15");
            var asig2 = CrearAsignatura(20, "CARTAGO", "", "ASIG 2", "2026-08-01", "2026-12-15");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig1, asig2 };

            var bloques = new List<ViaticoBloque>();
            for (int i = 1; i <= 20; i++)
            {
                bloques.Add(CrearBloque(new DateTime(2026, 8, 1).AddDays(i), 10, "CALI", ""));
            }

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques);

            // Assert: Ninguna fila supera 8 comisiones
            resultado.All(f => f.Bloques.Count <= 8).Should().BeTrue();
            resultado.Sum(f => f.Bloques.Count).Should().Be(20);
        }

        [Fact]
        public void BloqueFueraDeVentana_SeEscribeYAdvierte()
        {
            // Arrange: Asignatura con ventana en agosto
            var asig = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-08-01", "2026-08-31");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig };

            // Bloque en octubre (fuera de ventana U-V)
            var bloqueFuera = CrearBloque(new DateTime(2026, 10, 15), 10, "CALI", "");
            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, new List<ViaticoBloque> { bloqueFuera }, avisos);

            // Assert: El bloque se escribe y se registra advertencia R6
            resultado[0].Bloques.Should().HaveCount(1);
            avisos.Should().Contain(a => a.Contains("fuera de la ventana"));
        }

        [Fact]
        public void VentanaIsoYddMMyyyy_SeParseanAmbos()
        {
            // Arrange: Asignatura 1 con ISO, Asignatura 2 con dd/MM/yyyy
            var asig1 = CrearAsignatura(10, "CALI", "", "ASIG 1", "2026-09-01", "2026-09-30");
            var asig2 = CrearAsignatura(20, "CARTAGO", "", "ASIG 2", "01/10/2026", "31/10/2026");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig1, asig2 };

            var b1 = CrearBloque(new DateTime(2026, 9, 15), 10, "CALI", "");
            var b2 = CrearBloque(new DateTime(2026, 10, 15), 20, "CARTAGO", "");
            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, new List<ViaticoBloque> { b1, b2 }, avisos);

            // Assert: Ambas ventanas reconocen sus fechas dentro del rango sin advertencia R6
            resultado[0].Bloques.Should().HaveCount(1);
            resultado[1].Bloques.Should().HaveCount(1);
            avisos.Should().BeEmpty();
        }

        [Fact]
        public void SinAsignaturas_SinBloques_DevuelveVacio()
        {
            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(
                new List<IReadOnlyList<SesionListItem>>(),
                new List<ViaticoBloque>());

            // Assert
            resultado.Should().BeEmpty();
        }

        [Fact]
        public void CasoRealDosCalisCartagoElAguila_15Bloques_RespetaVentanaYFila()
        {
            // Arrange: 4 asignaturas con sedes repetidas (CALI) y ventanas de fecha en formato latino dd/MM/yyyy
            var asigCaliFund = CrearAsignatura(1, "CALI", "", "FUNDAMENTOS", "10/11/2026", "02/12/2026");
            var asigCaliHist = CrearAsignatura(2, "CALI", "", "HISTORIA", "22/09/2026", "04/11/2026");
            var asigCartago  = CrearAsignatura(3, "CARTAGO", "", "ECONOMIA", "14/11/2026", "06/12/2026");
            var asigAguila   = CrearAsignatura(4, "EL AGUILA", "", "DERECHO", "12/09/2026", "25/10/2026");

            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asigCaliFund, asigCaliHist, asigCartago, asigAguila };

            // 15 bloques reales en orden cronológico
            var bloques = new List<ViaticoBloque>
            {
                CrearBloque(new DateTime(2026, 9, 19), 4, "EL AGUILA", ""),
                CrearBloque(new DateTime(2026, 9, 22), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 9, 29), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 10, 3), 4, "EL AGUILA", ""),
                CrearBloque(new DateTime(2026, 10, 13), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 10, 20), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 10, 24), 4, "EL AGUILA", ""),
                CrearBloque(new DateTime(2026, 10, 27), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 11, 3), 2, "CALI", ""),
                CrearBloque(new DateTime(2026, 11, 10), 1, "CALI", ""),
                CrearBloque(new DateTime(2026, 11, 17), 1, "CALI", ""),
                CrearBloque(new DateTime(2026, 11, 21), 3, "CARTAGO", ""),
                CrearBloque(new DateTime(2026, 11, 24), 1, "CALI", ""),
                CrearBloque(new DateTime(2026, 12, 1), 1, "CALI", ""),
                CrearBloque(new DateTime(2026, 12, 5), 3, "CARTAGO", "")
            };

            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, bloques, avisos);

            // Assert: 4 / 6 / 2 / 3, total 15, cada bloque dentro de [U, V] de su fila y sin avisos
            resultado.Should().HaveCount(4);
            resultado[0].Bloques.Should().HaveCount(4); // Cali Fundamentos
            resultado[1].Bloques.Should().HaveCount(6); // Cali Historia
            resultado[2].Bloques.Should().HaveCount(2); // Cartago
            resultado[3].Bloques.Should().HaveCount(3); // El Águila
            resultado.Sum(f => f.Bloques.Count).Should().Be(15);
            resultado.Should().OnlyContain(f => !f.EsContinuacion);
            avisos.Should().BeEmpty();
        }

        [Fact]
        public void BloqueSinVinculacionCoincidente_AdvierteSinContarSaturacion()
        {
            // Arrange: 1 asignatura con IdVinculacion = 1
            var asig = CrearAsignatura(1, "CALI", "", "ASIG 1", "2026-08-01", "2026-12-15");
            var asignaturas = new List<IReadOnlyList<SesionListItem>> { asig };

            // Bloque huérfano con IdVinculacion = 999
            var bloqueHuerfano = CrearBloque(new DateTime(2026, 9, 10), 999, "CALI", "");
            var avisos = new List<string>();

            // Act
            var resultado = AsignaturaViaticoAllocator.Distribuir(asignaturas, new List<ViaticoBloque> { bloqueHuerfano }, avisos);

            // Assert: Se asigna a la fila con cupo emitiendo advertencia de "sin vinculación coincidente" y sin la palabra "saturación"
            resultado[0].Bloques.Should().HaveCount(1);
            avisos.Should().ContainSingle(a => a.Contains("sin vinculación coincidente") && !a.Contains("saturación"));
        }

        private static IReadOnlyList<SesionListItem> CrearAsignatura(
            int idVinculacion,
            string ciudad,
            string subsede,
            string nombreAsig,
            string u,
            string v)
        {
            return new List<SesionListItem>
            {
                new SesionListItem
                {
                    IdSesion = idVinculacion * 100,
                    IdVinculacion = idVinculacion,
                    SedeCiudad = ciudad,
                    Subsede = subsede,
                    AsignaturaNombre = nombreAsig,
                    Semestre = "I",
                    VinculacionFechaInicio = u,
                    VinculacionFechaFin = v,
                    Viatica = true,
                    DocenteNombre = "DOCENTE PRUEBA"
                }
            };
        }

        private static ViaticoBloque CrearBloque(DateTime fecha, int idVinculacion, string ciudad, string subsede)
        {
            var b = new ViaticoBloque();
            b.Agregar(new SesionListItem
            {
                IdSesion = (int)fecha.Ticks % 10000,
                IdVinculacion = idVinculacion,
                SedeCiudad = ciudad,
                Subsede = subsede,
                FechaSesion = fecha,
                Viatica = true
            });
            return b;
        }
    }
}
