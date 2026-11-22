using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgConcurrenciaCohorteTests
    {
        [Fact]
        public void Calcular_SinIntervalos_RetornaFactibleSinEvaluacion()
        {
            var res = AgConcurrenciaCohorte.Calcular(new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>(), 2);

            res.Evaluado.Should().BeFalse();
            res.Factible.Should().BeTrue();
            res.MaxConcurrentes.Should().Be(0);
            res.CarrilesDisponibles.Should().Be(2);
        }

        [Fact]
        public void Calcular_CuandoConcurrentesNoExcedenCarriles_RetornaFactibleTrue()
        {
            var intervalos = new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                ("C1", 101, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15)),
                ("C1", 102, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15))
            };

            var res = AgConcurrenciaCohorte.Calcular(intervalos, 2);

            res.Evaluado.Should().BeTrue();
            res.Factible.Should().BeTrue();
            res.MaxConcurrentes.Should().Be(2);
            res.CarrilesDisponibles.Should().Be(2);
        }

        [Fact]
        public void Calcular_CuandoConcurrentesExcedenCarriles_RetornaFactibleFalse()
        {
            var intervalos = new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                ("C1", 101, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15)),
                ("C1", 102, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15)),
                ("C1", 103, new DateTime(2026, 3, 5), new DateTime(2026, 3, 10))
            };

            var res = AgConcurrenciaCohorte.Calcular(intervalos, 2);

            res.Evaluado.Should().BeTrue();
            res.Factible.Should().BeFalse();
            res.MaxConcurrentes.Should().Be(3);
            res.CarrilesDisponibles.Should().Be(2);
        }

        [Fact]
        public void DetectarSobrecargas_IdentificaTramosConExceso()
        {
            var intervalos = new List<(int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                (101, new DateTime(2026, 3, 1), new DateTime(2026, 3, 10)),
                (102, new DateTime(2026, 3, 1), new DateTime(2026, 3, 10)),
                (103, new DateTime(2026, 3, 5), new DateTime(2026, 3, 8))
            };

            var sobrecargas = AgConcurrenciaCohorte.DetectarSobrecargas(intervalos, 2);

            sobrecargas.Should().ContainSingle();
            sobrecargas[0].TramoIni.Should().Be(new DateTime(2026, 3, 5));
            sobrecargas[0].TramoFin.Should().Be(new DateTime(2026, 3, 8));
            sobrecargas[0].MaxConcurrentes.Should().Be(3);
        }

        [Fact]
        public void Calcular_ConMultiplesCohortes_LanzaArgumentException()
        {
            var intervalos = new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                ("C1", 101, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15)),
                ("C2", 102, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15))
            };

            Action act = () => AgConcurrenciaCohorte.Calcular(intervalos, 2);

            act.Should().Throw<ArgumentException>()
                .WithMessage("*sola cohorte*");
        }

        [Fact]
        public void Calcular_ConIntervaloInvertido_NoDeclaraFactible()
        {
            var intervalos = new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                ("C1", 101, new DateTime(2026, 3, 15), new DateTime(2026, 3, 1))
            };

            var res = AgConcurrenciaCohorte.Calcular(intervalos, 2);

            res.Evaluado.Should().BeFalse();
            res.Factible.Should().BeFalse();
        }

        [Fact]
        public void Calcular_ConCohorteVaciaYNoVacia_LanzaArgumentException()
        {
            var intervalos = new List<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)>
            {
                ("", 101, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15)),
                ("C1", 102, new DateTime(2026, 3, 1), new DateTime(2026, 3, 15))
            };

            Action act = () => AgConcurrenciaCohorte.Calcular(intervalos, 2);

            act.Should().Throw<ArgumentException>()
                .WithMessage("*sola cohorte*");
        }
    }
}
