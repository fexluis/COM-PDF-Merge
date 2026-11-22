using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgDimensionesPerfilAulaTests
    {
        [Fact]
        public void Resolver_PrevalecePerfilSobreAula_WhenDifiere()
        {
            var aula = new AgAula
            {
                AulaRef = "1|1|1|I|A|MIXTA",
                DiasVentana = "MAR-VIE",
                Jornada = "TARDE",
                HorasDia = 4
            };
            var perfil = new AgContextoSeleccionDto
            {
                AulaRef = aula.AulaRef,
                DiasVentana = "1,2,3,4,5",
                Jornada = "TARDE",
                HorasDia = 4
            };

            var res = AgDimensionesPerfilResolver.Aplicar(new[] { aula }, new[] { perfil }, out var diffs);

            res.Single().DiasVentana.Should().Be("1,2,3,4,5");
            diffs.Should().ContainSingle();
        }

        [Fact]
        public void Resolver_NoModificaAula_WhenCoincideConPerfil()
        {
            var aula = new AgAula { AulaRef = "1|1|1|I|A|MIXTA", DiasVentana = "1,2,3,4,5", Jornada = "NOCHE", HorasDia = 4 };
            var perfil = new AgContextoSeleccionDto { AulaRef = aula.AulaRef, DiasVentana = "1,2,3,4,5", Jornada = "NOCHE", HorasDia = 4 };

            var res = AgDimensionesPerfilResolver.Aplicar(new[] { aula }, new[] { perfil }, out var diffs);

            res.Single().DiasVentana.Should().Be("1,2,3,4,5");
            diffs.Should().BeEmpty();
        }

        [Fact]
        public void Detectar_RetornaDiferencias_WhenDimensionesDifieren()
        {
            var aula = new AgAula { AulaRef = "1|1|1|I|A|MIXTA", DiasVentana = "MAR-VIE", Jornada = "NOCHE", HorasDia = 4, Subunidad = "SEDE1" };
            var perfil = new AgContextoSeleccionDto { AulaRef = aula.AulaRef, DiasVentana = "1,2,3,4,5", Jornada = "MANANA", HorasDia = 4 };

            var diffs = AgDimensionesPerfilResolver.Detectar(new[] { aula }, new[] { perfil });

            diffs.Should().ContainSingle();
            diffs[0].AulaRef.Should().Be(aula.AulaRef);
            diffs[0].DiasVentana.Should().Be("1,2,3,4,5");
            diffs[0].Jornada.Should().Be("MANANA");
            diffs[0].Subunidad.Should().Be("SEDE1");
        }

        [Fact]
        public void Resolver_PerfilInvalido_NoModificaAulaNiGeneraDiff()
        {
            var aula = new AgAula { AulaRef = "1|1|1|I|A|MIXTA", DiasVentana = "MAR-VIE", Jornada = "NOCHE", HorasDia = 4 };
            var perfilInvalido = new AgContextoSeleccionDto { AulaRef = aula.AulaRef, DiasVentana = "1,2,3,4,5", Jornada = "", HorasDia = 0 };

            var res = AgDimensionesPerfilResolver.Aplicar(new[] { aula }, new[] { perfilInvalido }, out var diffs);
            var detectadas = AgDimensionesPerfilResolver.Detectar(new[] { aula }, new[] { perfilInvalido });

            res.Single().DiasVentana.Should().Be("MAR-VIE");
            res.Single().Jornada.Should().Be("NOCHE");
            res.Single().HorasDia.Should().Be(4);
            diffs.Should().BeEmpty();
            detectadas.Should().BeEmpty();
        }
    }
}
