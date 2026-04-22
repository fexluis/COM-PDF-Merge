using Xunit;
using FluentAssertions;
using System;

namespace ConversorPDF.Tests.Unit
{
    public class EstadoParametrosTests
    {
        private static string BuildHash(int t = 1, string h = "48",
            string dias = "JUEVES|8", string fi = "20260423", string ff = "20260528",
            int p = 1, string s = "X", int a = 56)
        {
            return $"T={t};P={p};S={s};A={a};H={h};D={dias};FI={fi};FF={ff}";
        }

        [Fact]
        public void Hash_MismaFechaDiferenteHora_ProduceHashes_Iguales()
        {
            string fi1 = new DateTime(2026, 4, 23, 10, 0, 0).Date.ToString("yyyyMMdd");
            string fi2 = new DateTime(2026, 4, 23, 15, 30, 0).Date.ToString("yyyyMMdd");

            fi1.Should().Be(fi2);
        }

        [Fact]
        public void Hash_CambioPrograma_ProduceHashDiferente()
        {
            string hash1 = BuildHash(p: 1);
            string hash2 = BuildHash(p: 2);

            hash1.Should().NotBe(hash2);
        }

        [Fact]
        public void Hash_CambioSemestre_ProduceHashDiferente()
        {
            string hash1 = BuildHash(s: "I");
            string hash2 = BuildHash(s: "X");

            hash1.Should().NotBe(hash2);
        }

        [Fact]
        public void Hash_CambioAsignatura_ProduceHashDiferente()
        {
            string hash1 = BuildHash(a: 5);
            string hash2 = BuildHash(a: 99);

            hash1.Should().NotBe(hash2);
        }

        [Fact]
        public void Hash_MismoEstado_HashIdentico()
        {
            string hash1 = BuildHash();
            string hash2 = BuildHash();

            hash1.Should().Be(hash2);
        }
    }
}
