using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPlanificadorOleadasTests
    {
        private readonly AgPlanificadorOleadas _sut = new AgPlanificadorOleadas();

        [Fact]
        public void Construir_When4AsignaturasAnd16SemanasMarVie_ShouldGenerateDisjointLanesAndSequentialWaves()
        {
            // Arrange (16 semanas: del 11 de agosto al 30 de noviembre de 2026)
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 11, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 64.0, 1),
                (102, 64.0, 2),
                (103, 64.0, 3),
                (104, 64.0, 4)
            };

            // Act
            var result = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            result.IsSuccess.Should().BeTrue();
            var planes = result.Value.Planes;
            planes.Should().HaveCount(4);
            result.Value.Infactibles.Should().BeEmpty();

            result.Value.Particion.Should().HaveCount(2);                       // antes 4: ahora 2 carriles
            result.Value.Particion.SelectMany(b => b).Distinct().Count().Should().Be(4);   // días {2,3,4,5} — se mantiene
            planes.Select(p => p.Carril).Distinct().Should().HaveCount(2);      // antes 4
            planes.Select(p => p.Oleada).Distinct().OrderBy(o => o).Should().Equal(0, 1);
            foreach (var g in planes.GroupBy(p => p.Carril))
            {
                var secuencia = g.OrderBy(p => p.Oleada).ToList();
                for (int i = 1; i < secuencia.Count; i++)
                    secuencia[i].FechaInicio.Should().BeAfter(secuencia[i - 1].FechaFin);   // oleadas disjuntas
            }
        }

        [Fact]
        public void Construir_WhenVentanaDemasiadoCortaParaCapacidad_ShouldFailWithI5()
        {
            // Arrange (Solo 4 semanas de ventana para 2 oleadas de 64h = 16 semanas requeridas)
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 8, 15);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 64.0, 1),
                (102, 64.0, 2),
                (103, 64.0, 3),
                (104, 64.0, 4)
            };

            // Act
            var result = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert (Fase 1242: código vigente tras Fase 1239)
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I5_SIN_CAPACIDAD");
        }

        [Fact]
        public void Construir_WhenVentanaNoDefinida_ShouldFail()
        {
            // Arrange
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 11, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 64.0, 1)
            };

            // Act
            var result = _sut.Construir(asignaturas, "", fInicio, fFin, 4.0);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I1_VENTANA_NO_DEFINIDA");
        }

        [Fact]
        public void Construir_WhenListaAsignaturasVacia_ShouldReturnSuccessEmpty()
        {
            // Arrange
            var fInicio = new DateTime(2026, 8, 11);
            var fFin = new DateTime(2026, 11, 30);

            // Act
            var result = _sut.Construir(new List<(int, double, int)>(), "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Planes.Should().BeEmpty();
            result.Value.Infactibles.Should().BeEmpty();
        }

        [Fact]
        public void Oleadas_ConFestivosEnLaOleada_DetectaI5AntesDeConstruir()
        {
            // Arrange (Fase 1078 / Fase 1092): 4 semanas para 1 asignatura de 32h (8 sesiones de 4h).
            // En ventana MAR-VIE (carril 0 = MAR, MIE -> 8 slots teóricos).
            // Si hay un festivo en un martes de la ventana, la capacidad real cae a 7 ranuras < 8 requeridas -> infactible.
            var fInicio = new DateTime(2026, 8, 10);
            var fFin = new DateTime(2026, 9, 6);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 32.0, 1)
            };

            // Martes 18 de agosto festivo
            Func<DateTime, bool> esFestivo = d => d.Date == new DateTime(2026, 8, 18);

            // Act
            var result = _sut.Construir(asignaturas, "MAR,MIE", fInicio, fFin, 4.0, esFestivo);

            // Assert (Fase 1273: detectado antes de construir como I5_SIN_CAPACIDAD con festivos descontados)
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I5_SIN_CAPACIDAD");
        }

        [Fact]
        public void Oleadas_ValidacionDelegadaEnConstruir_DetectaI5()
        {
            // Arrange: 10 sesiones en ventana de 2 semanas MAR-VIE (capacidad 8) -> I5_SIN_CAPACIDAD
            var fInicio = new DateTime(2026, 8, 10);
            var fFin = new DateTime(2026, 8, 23); // 2 semanas = 8 ranuras para MAR-VIE
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 40.0, 1) // 40h / 4h = 10 sesiones requeridas, no caben en 8 ranuras
            };

            // Act
            var result = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("I5_SIN_CAPACIDAD");
        }

        [Fact]
        public void Oleadas_UnaAsignaturaInfactible_PlanesDelRestoSeConservan()
        {
            // Arrange (Fase 1256): Asignatura 101 con patron violado R8, Asignatura 102 sin restricciones
            var fInicio = new DateTime(2026, 8, 10);
            var fFin = new DateTime(2026, 9, 6); // 4 semanas
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 8.0, 1),
                (102, 8.0, 2)
            };
            var patrones = new Dictionary<int, string> { { 101, "1,3" } }; // Dia 3 cruza ventana LUN,MAR

            // Act
            var result = _sut.Construir(asignaturas, "LUN,MAR", fInicio, fFin, 4.0, diasPatronPorAsignatura: patrones);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().Contain(101);
            result.Value.Planes.Should().ContainSingle(p => p.IdAsignatura == 102);
        }

        [Fact]
        public void Oleada_CapacidadRanuras_NoEsIgualASesionesRequeridas()
        {
            // Arrange: 1 asig de 32h = 8 sesiones en MAR-VIE con 8 semanas de ventana
            var fInicio = new DateTime(2026, 8, 10);
            var fFin = new DateTime(2026, 10, 4); // 8 semanas
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (101, 32.0, 1)
            };

            // Act
            var result = _sut.Construir(asignaturas, "MAR-VIE", fInicio, fFin, 4.0);

            // Assert
            result.IsSuccess.Should().BeTrue();
            var plan = result.Value.Planes.Single();
            plan.SesionesRequeridas.Should().Be(8);
            plan.CapacidadRanuras.Should().Be(8);
        }

        [Fact]
        public void Construir_CuandoFestivosEstrechanLaOleadaFinal_NoMarcaInfactible()
        {
            // Arrange: caso real 2027-1 (aula 22 LUN-JUE). 6 asigs, 22 semanas, festivos que caen en lunes (Fase 1140 / Fase 1144).
            var fInicio = new DateTime(2027, 2, 2);
            var fFin = new DateTime(2027, 6, 30);
            var asignaturas = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 48.0, 1), (2, 48.0, 1), (3, 32.0, 1), (4, 32.0, 1), (5, 48.0, 1), (6, 48.0, 1)
            };
            Func<DateTime, bool> esFestivo = d =>
                d.Date == new DateTime(2027, 5, 1) || d.Date == new DateTime(2027, 5, 10) ||
                d.Date == new DateTime(2027, 5, 31) || d.Date == new DateTime(2027, 6, 7) ||
                (d.Date >= new DateTime(2027, 3, 22) && d.Date <= new DateTime(2027, 3, 28));

            // Act
            var result = _sut.Construir(asignaturas, "LUN-JUE", fInicio, fFin, 4.0, esFestivo);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Infactibles.Should().BeEmpty();
            result.Value.Planes.Should().HaveCount(6);
            result.Value.Planes.Should().OnlyContain(p => p.CapacidadRanuras >= p.SesionesRequeridas);
            result.Value.Planes.Should().OnlyContain(p => p.FechaInicio >= fInicio && p.FechaFin <= fFin);
            foreach (var g in result.Value.Planes.GroupBy(p => p.Carril))
            {
                var secuencia = g.OrderBy(p => p.Oleada).ToList();
                for (int i = 1; i < secuencia.Count; i++)
                    secuencia[i].FechaInicio.Should().BeAfter(secuencia[i - 1].FechaFin);
            }
        }

        [Fact]
        public void Construir_CuandoAsignaturaInfactible_ExponeDetalleAccionable()
        {
            // Arrange (Fase 1256): 1 asignatura con patron incompatible con ventana -> infactible con detalle
            var patrones = new Dictionary<int, string> { { 101, "2,6" } }; // Dia 6 no vive en MAR-VIE
            var result = _sut.Construir(
                new List<(int, double, int)> { (101, 32.0, 1) }, "MAR-VIE",
                new DateTime(2026, 8, 10), new DateTime(2026, 8, 23), 4.0, diasPatronPorAsignatura: patrones);

            // Assert
            result.Value.Infactibles.Should().Contain(101);
            result.Value.DetalleInfactibles.Should().ContainSingle()
                .Which.Should().Contain("Asig 101").And.Contain("R8_DiaFueraDePatron");
        }

        [Fact]
        public void Construir_WhenJueSab4hConBloques_ShouldPlanificarDosCarriles()
        {
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 1, DiaSemanaNumero = 4, Franja = "NOCHE",  DuracionHoras = 4, Activo = 1 },
                new AgBloqueHorario { Id = 2, DiaSemanaNumero = 5, Franja = "NOCHE",  DuracionHoras = 4, Activo = 1 },
                new AgBloqueHorario { Id = 3, DiaSemanaNumero = 6, Franja = "NOCHE",  DuracionHoras = 4, Activo = 1 },
                new AgBloqueHorario { Id = 4, DiaSemanaNumero = 6, Franja = "TARDE",  DuracionHoras = 4, Activo = 1 }
            };
            var sut = new AgPlanificadorOleadas();
            var res = sut.Construir(
                new List<(int, double, int)> { (101, 24.0, 1), (102, 24.0, 1), (103, 24.0, 1) },
                "JUE,SAB", new DateTime(2026, 8, 13), new DateTime(2026, 12, 3), 4.0, null, bloques, "NOCHE");
            res.IsSuccess.Should().BeTrue();
            res.Value.Particion.SelectMany(b => b).Distinct().Count().Should().Be(2);
            res.Value.Planes.Select(p => p.Carril).Distinct().Should().HaveCount(2);
        }

        [Fact]
        public void Construir_WhenCatalogoNoCubreLaVentana_ShouldDegradarSinFallarElAula()
        {
            var bloques = new List<AgBloqueHorario>
            {
                new AgBloqueHorario { Id = 1, DiaSemanaNumero = 2, Franja = "MANANA", DuracionHoras = 8, Activo = 1 }
            };
            var sut = new AgPlanificadorOleadas();

            var res = sut.Construir(
                new List<(int, double, int)> { (101, 16.0, 1) }, "LUN,MAR",
                new DateTime(2026, 8, 10), new DateTime(2026, 12, 3), 8.0, null, bloques, "NOCHE");

            res.IsSuccess.Should().BeTrue();
            res.Value.Planes.Should().HaveCount(1);
        }

        [Fact]
        public void Construir_WhenDemandaCabeEnCapacidadRanurasAunqueWMenorQueOlas_ShouldSucceed()
        {
            // Arrange (Fase 1239): 6 asignaturas de 4h (6 sesiones), W=2 semanas en VIE-SAB (3 ranuras/sem -> capacidad 6)
            var sut = new AgPlanificadorOleadas();
            var asignaturas = new List<(int, double, int)>
            {
                (1, 4.0, 1), (2, 4.0, 2), (3, 4.0, 3), (4, 4.0, 4), (5, 4.0, 5), (6, 4.0, 6)
            };

            // Act
            var res = sut.Construir(asignaturas, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 3, 15), 4.0);

            // Assert
            res.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void Construir_WhenDemandaExcedeCapacidadRanuras_ShouldFailWithI5SinCapacidad()
        {
            // Arrange (Fase 1239): 6 asignaturas de 64h (96 sesiones), W=2 semanas en VIE-SAB (capacidad 6)
            var sut = new AgPlanificadorOleadas();
            var asignaturas = new List<(int, double, int)>
            {
                (1, 64.0, 1), (2, 64.0, 2), (3, 64.0, 3), (4, 64.0, 4), (5, 64.0, 5), (6, 64.0, 6)
            };

            // Act
            var res = sut.Construir(asignaturas, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 3, 15), 4.0);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SIN_CAPACIDAD");
        }

        [Fact]
        public void Construir_Caso1_60SesionesVieSab_ShouldPlanificar20Semanas()
        {
            var asigs = Enumerable.Range(1, 15).Select(i => (i, 16.0, i)).ToList();   // 15 × 4 ses = 60
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 8, 3), 4.0);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.GroupBy(p => p.Carril).Select(g => g.Sum(p => p.SemanasOleada)).Max().Should().Be(20);
            res.Value.OcupacionPorcentaje.Should().BeGreaterThan(88.0);
        }

        [Fact]
        public void Construir_Caso2_64SesionesVieSab_ShouldPlanificar22SemanasConAltaOcupacion()
        {
            var asigs = Enumerable.Range(1, 16).Select(i => (i, 16.0, i)).ToList();   // 16 × 4 ses = 64
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 8, 10), 4.0);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.GroupBy(p => p.Carril).Select(g => g.Sum(p => p.SemanasOleada)).Max().Should().Be(22);
            res.Value.OcupacionPorcentaje.Should().BeApproximately(92.8, 1.5);
        }

        [Fact]
        public void Construir_WhenSinSemanasHabiles_ShouldFailWithI5SinSemanasHabiles()
        {
            var asigs = new List<(int, double, int)> { (1, 16.0, 1) };
            var res = _sut.Construir(asigs, "DOM", new DateTime(2026, 3, 2), new DateTime(2026, 3, 15), 4.0, d => true);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("I5_SIN_SEMANAS_HABILES");
        }

        [Fact]
        public void Construir_CuandoPatronComun_AmbasAsignaturasCompartenBanda()
        {
            var asigs = new List<(int, double, int)> { (1, 16.0, 1), (2, 16.0, 2) };
            var patrones = new Dictionary<int, string> { { 1, "5" }, { 2, "5" } };
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 4, 13), 4.0,
                diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.Should().HaveCount(2);
            res.Value.Planes.Select(p => p.Carril).Distinct().Should().HaveCount(1);
            res.Value.PatronesRenormalizados.Should().Be(0);
        }

        [Fact]
        public void Construir_CuandoInfactibleEnVieSab_SugiereJueSab()
        {
            var asigs = new List<(int, double, int)> { (1, 32.0, 1) };
            var patrones = new Dictionary<int, string> { { 1, "4,5" } }; // patron cruza a JUE
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 3, 29), 4.0, diasPatronPorAsignatura: patrones);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().Contain(1);
            res.Value.VentanaSugerida.Should().Be("JUE-SAB");
        }

        [Fact]
        public void Construir_When6AsignaturasJueSab_ShouldSecuenciarOleadasPorCarrilYNoApilar()
        {
            // Arrange: caso real prog 4 (aula JUE-SAB, MANANA, 4h, semestre 2027-1)
            var fInicio = new DateTime(2027, 2, 4);
            var fFin = new DateTime(2027, 7, 2);
            var asigs = new List<(int IdAsignatura, double Horas, int OrdenCurricular)>
            {
                (1, 48.0, 1), (2, 48.0, 2), (3, 48.0, 3), (4, 32.0, 4), (5, 32.0, 5), (6, 32.0, 6)
            };
            var patrones = new Dictionary<int, string>
            { { 1, "4,5" }, { 2, "6" }, { 3, "4,5" }, { 4, "6" }, { 5, "4,5" }, { 6, "6" } };

            // Act
            var res = _sut.Construir(asigs, "JUE-SAB", fInicio, fFin, 4.0,
                diasPatronPorAsignatura: patrones);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.Should().HaveCount(6);
            res.Value.Particion.Should().HaveCount(2);

            res.Value.Planes.Select(p => new { p.Carril, p.FechaInicio, p.FechaFin })
                .Distinct().Should().HaveCount(6, "cada asignatura debe tener su propia sub-ventana y carril");

            foreach (var g in res.Value.Planes.GroupBy(p => p.Carril))
            {
                var secuencia = g.OrderBy(p => p.Oleada).ToList();
                secuencia.Select(p => p.Oleada).Should().Equal(Enumerable.Range(0, secuencia.Count));
                for (int i = 1; i < secuencia.Count; i++)
                    secuencia[i].FechaInicio.Should().BeAfter(secuencia[i - 1].FechaFin);
            }

            // Un carril = una asignatura activa por semana (R5)
            var carriles = res.Value.Particion.Count;
            for (var w = fInicio; w <= fFin; w = w.AddDays(7))
            {
                var finSem = w.AddDays(6);
                res.Value.Planes.Count(p => p.FechaInicio <= finSem && p.FechaFin >= w)
                    .Should().BeLessOrEqualTo(carriles);
            }
        }

        [Fact]
        public void Construir_WhenCarrilesDesbalanceados_ShouldRepartirProporcionalALaUtilizacion()
        {
            var asigs = Enumerable.Range(1, 15).Select(i => (i, 16.0, i)).ToList();
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 8, 3), 4.0);

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().BeEmpty();
            res.Value.Planes.Should().HaveCount(15);
            var conteoPorCarril = res.Value.Planes.GroupBy(p => p.Carril).OrderBy(g => g.Key).Select(g => g.Count()).ToList();
            conteoPorCarril.Should().HaveCount(2);
            conteoPorCarril[0].Should().Be(10);   // carril [VIE,SÁB-M]: 2 slots/sem ⇒ 40 ses
            conteoPorCarril[1].Should().Be(5);    // carril [SÁB-T]: 1 slot/sem ⇒ 20 ses
            res.Value.Planes.Where(p => p.Carril == 1).Sum(p => p.SesionesRequeridas)
                .Should().BeLessOrEqualTo(22);
        }

        [Fact]
        public void Construir_WhenDemandaExcedeCapacidadDeUnCarril_ShouldMarcarInfactibleConDetalle()
        {
            var asigs = new List<(int, double, int)>
            {
                (1, 64.0, 1),
                (2, 64.0, 2),
                (3, 64.0, 3)
            };
            var patrones = new Dictionary<int, string> { { 1, "5" }, { 2, "5" }, { 3, "5" } };
            var res = _sut.Construir(asigs, "VIE-SAB", new DateTime(2026, 3, 2), new DateTime(2026, 6, 30), 4.0,
                diasPatronPorAsignatura: patrones);   // W=17, capacidad total 51 ≥ 48; carril 0 = 34 < 48

            res.IsSuccess.Should().BeTrue();
            res.Value.Infactibles.Should().Contain(3);           // 3ª asignatura no cabe en carril 0 (34 ses)
            res.Value.DetalleInfactibles.Should().ContainSingle()
                .Which.Should().Contain("Asig 3").And.Contain("R8_DiaFueraDePatron");
        }

        [Fact]
        public void Construir_ConFestivoDentroDeLaOleada_ShouldExtenderUnaSemanaSinPerderSesiones()
        {
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 5, 31);
            var asigs = new List<(int, double, int)> { (1, 16.0, 1) };
            Func<DateTime, bool> esFestivo = d => d.Date == new DateTime(2026, 3, 9);

            var res = _sut.Construir(asigs, "LUN", fInicio, fFin, 4.0, esFestivo: esFestivo);

            res.IsSuccess.Should().BeTrue();
            res.Value.Planes.Should().HaveCount(1);
            var plan = res.Value.Planes[0];
            plan.SesionesRequeridas.Should().Be(4);
            plan.SemanasOleada.Should().Be(4);                          // 4 lunes hábiles
            plan.FechaFin.Should().Be(new DateTime(2026, 4, 5));        // 5 semanas calendario (festivo absorbido)
        }

        [Fact]
        public void Construir_WhenVentanaDe2Dias_ShouldUnSoloCarril()
        {
            var fInicio = new DateTime(2026, 3, 2);
            var fFin = new DateTime(2026, 5, 31);
            var asigs = new List<(int, double, int)> { (1, 16.0, 1), (2, 16.0, 2) };

            var res = _sut.Construir(asigs, "LUN-MAR", fInicio, fFin, 4.0);

            res.IsSuccess.Should().BeTrue();
            res.Value.Particion.Should().HaveCount(1);
            res.Value.Planes.Select(p => p.Carril).Distinct().Should().Equal(0);
            res.Value.Planes.Select(p => p.Oleada).OrderBy(o => o).Should().Equal(0, 1);
        }
    }
}
