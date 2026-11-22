using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgInventarioRanurasTests
    {
        [Fact]
        public void Construir_Fixture132Sintetico_CalculaMetricasCanónicas()
        {
            // Arrange: Crear ventana con 86 fechas válidas, 72 ocupadas, 14 libres (4 huecos + 10 cola), M=2, ocupación 83.7%
            var fechaIni = new DateTime(2026, 2, 2);
            var fechasValidas = new List<DateTime>();
            var fechaActual = fechaIni;

            // Generar exactamente 86 fechas válidas (Lunes a Jueves)
            while (fechasValidas.Count < 86)
            {
                int diaIso = (fechaActual.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)fechaActual.DayOfWeek;
                if (diaIso >= 1 && diaIso <= 4)
                {
                    fechasValidas.Add(fechaActual);
                }
                fechaActual = fechaActual.AddDays(1);
            }

            var fechaFin = fechasValidas.Last();

            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre
                {
                    FechaInicio = fechaIni,
                    FechaFin = fechaFin,
                    PeriodoNombre = "2026-1",
                    Origen = AgOrigenVentana.Default
                },
                Aulas = new List<AgAula>
                {
                    new AgAula
                    {
                        Id = 1,
                        Nombre = "Aula 101",
                        DiasVentana = "LUN,MAR,MIE,JUE",
                        Jornada = "MANANA",
                        HorasDia = 4
                    }
                },
                Bloques = new List<AgBloqueHorario>
                {
                    new AgBloqueHorario { Id = 1, DiaSemanaNumero = 1, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 2, DiaSemanaNumero = 2, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 3, DiaSemanaNumero = 3, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 4, DiaSemanaNumero = 4, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" },
                    new AgBloqueHorario { Id = 5, DiaSemanaNumero = 5, DuracionHoras = 4.0, Activo = 1, Franja = "MANANA" }
                }
            };
            ctx.ConstruirIndices();

            // 86 válidas en total.
            // 10 últimas son cola muerta (fechasValidas[76..85]).
            // De las primeras 76 fechas: 72 ocupadas y 4 huecos intermedios (ej. índices 20, 35, 50, 65 son libres).
            var indicesHuecos = new HashSet<int> { 20, 35, 50, 65 };
            var fechasOcupadas = new List<DateTime>();
            for (int i = 0; i < 76; i++)
            {
                if (!indicesHuecos.Contains(i))
                {
                    fechasOcupadas.Add(fechasValidas[i]);
                }
            }

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen
            {
                Id = 1,
                IdAsignatura = 100,
                IdAula = 1,
                Semestre = "I",
                Grupo = "A",
                HorasGen = fechasOcupadas.Count * 4.0,
                FechasGen = fechasOcupadas,
                FechaInicioEstimada = fechasOcupadas[0],
                FechaFinEstimada = fechasOcupadas[fechasOcupadas.Count - 1],
                EsVacante = 0
            });

            string cohorte = AgCalendarioConflictos.Cohorte(ind.Genes[0]);

            // Act
            var inv = AgInventarioRanuras.Construir(ind, ctx, cohorte);

            // Assert
            inv.FechasValidas.Count.Should().Be(86);
            inv.Ocupadas.Count.Should().Be(72);
            inv.LibresReales.Count.Should().Be(14);
            inv.HuecosIntermedios.Should().Be(4);
            inv.ColaMuerta.Should().Be(10);
            inv.M.Should().Be(2);
            Math.Round(inv.Ocupacion * 100.0, 1).Should().Be(83.7);
        }

        [Fact]
        public void Construir_SinGenes_RetornaColeccionesVacias()
        {
            var ctx = new AgContexto();
            var ind = new AgIndividuo();

            var inv = AgInventarioRanuras.Construir(ind, ctx, "SIN_GRUPO");

            inv.FechasValidas.Should().BeEmpty();
            inv.Ocupadas.Should().BeEmpty();
            inv.LibresReales.Should().BeEmpty();
            inv.Liberables.Should().BeEmpty();
            inv.Ocupacion.Should().Be(0.0);
        }

        [Fact]
        public void Construir_ConGenesLegadoYFijo_NoLosIncluyeEnLiberables()
        {
            var fIni = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 3, 6);
            var ctx = new AgContexto
            {
                Ventana = new VentanaSemestre { FechaInicio = fIni, FechaFin = fFin },
                Aulas = new List<AgAula> { new AgAula { Id = 1, DiasVentana = "LUN,MAR,MIE,JUE,VIE", Jornada = "MANANA" } },
                AulaAsignaturas = new List<AgAulaAsignatura>
                {
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 10, OrigenPatron = "LEGADO" },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 20, OrigenPatron = "AUTOMATICO" },
                    new AgAulaAsignatura { IdAula = 1, IdAsignatura = 30, OrigenPatron = "AUTOMATICO" }
                }
            };
            ctx.ConstruirIndices();

            var ind = new AgIndividuo();
            ind.Genes.Add(new AgGen { Id = 1, IdAsignatura = 10, IdAula = 1, Semestre = "I", Grupo = "A", FechasGen = new List<DateTime> { fIni }, EsFijo = 0 });
            ind.Genes.Add(new AgGen { Id = 2, IdAsignatura = 20, IdAula = 1, Semestre = "I", Grupo = "A", FechasGen = new List<DateTime> { fIni.AddDays(1) }, EsFijo = 1 }); // EsFijo=1
            ind.Genes.Add(new AgGen { Id = 3, IdAsignatura = 30, IdAula = 1, Semestre = "I", Grupo = "A", FechasGen = new List<DateTime> { fIni.AddDays(2) }, EsFijo = 0 }); // Movible

            var inv = AgInventarioRanuras.Construir(ind, ctx, AgCalendarioConflictos.Cohorte(ind.Genes[0]));

            inv.Liberables.Should().ContainSingle().Which.Should().Be(fIni.AddDays(2));
        }
    }
}
