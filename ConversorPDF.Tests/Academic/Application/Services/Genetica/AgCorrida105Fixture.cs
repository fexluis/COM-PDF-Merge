using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public static class AgCorrida105Fixture
    {
        public const int TotalAulasEsperadas = 35;
        public const int TotalAsignaturasEsperadas = 102;
        public const int TotalBloquesEsperados = 32;
        public const int TotalDocentesEsperados = 220;
        public const int TotalGenesEsperados = 410;

        public static AgContexto CrearContexto105(int seed = 105)
        {
            var rng = new Random(seed);
            var contexto = new AgContexto
            {
                IdSede = 1,
                Subunidad = "FACULTAD_PREGRADO",
                Ventana = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 2), new DateTime(2026, 6, 20), AgOrigenVentana.Presencial).Value
            };

            // 1. 35 Aulas
            for (int a = 1; a <= TotalAulasEsperadas; a++)
            {
                var aula = new AgAula
                {
                    Id = a,
                    IdSede = 1,
                    Nombre = $"Aula-{a:D2}",
                    Estado = AgEstadoAula.Disponible,
                    Capacidad = 35 + (a % 10),
                    Semestre = ((a % 8) + 1).ToString()
                };
                contexto.Aulas.Add(aula);
                contexto.AulasPorId[a] = aula;
            }

            // 2. 102 Asignaturas
            for (int s = 1; s <= TotalAsignaturasEsperadas; s++)
            {
                var asig = new Asignatura
                {
                    Id = s,
                    Nombre = $"Asignatura-{s:D3}",
                    Horas = (s % 3 == 0) ? 64 : 48,
                    Semestre = ((s % 8) + 1).ToString()
                };
                contexto.Asignaturas.Add(asig);
                contexto.AsignaturasPorId[s] = asig;
            }

            // 3. 32 Bloques Horarios
            var dias = new[] { 1, 2, 3, 4, 5, 6, 7 }; // Lunes a Domingo
            var franjas = new[]
            {
                ("07:00", "11:00"),
                ("09:00", "13:00"),
                ("14:00", "18:00"),
                ("16:00", "20:00"),
                ("18:00", "22:00")
            };

            int bloqueId = 1;
            for (int d = 0; d < dias.Length && bloqueId <= TotalBloquesEsperados; d++)
            {
                for (int f = 0; f < franjas.Length && bloqueId <= TotalBloquesEsperados; f++)
                {
                    var (ini, fin) = franjas[f];
                    var bloque = new AgBloqueHorario
                    {
                        Id = bloqueId,
                        DiaSemanaNumero = dias[d],
                        HoraInicio = ini,
                        HoraFin = fin,
                        DuracionHoras = 4,
                        Activo = 1
                    };
                    contexto.Bloques.Add(bloque);
                    contexto.BloquesPorId[bloqueId] = bloque;
                    bloqueId++;
                }
            }

            // 4. 220 Docentes y Habilitaciones en Pool
            for (int d = 1; d <= TotalDocentesEsperados; d++)
            {
                var docente = new Tercero
                {
                    Id = d,
                    Nombre = $"Docente-{d:D3}",
                    Activo = 1
                };
                contexto.Docentes.Add(docente);
                contexto.DocentesPorId[d] = docente;
            }

            // Asignar al menos 2-3 docentes habilitados por asignatura
            for (int s = 1; s <= TotalAsignaturasEsperadas; s++)
            {
                var habilitados = new List<TerceroAsignatura>();
                int cantDocentes = 2 + (s % 3);
                for (int h = 0; h < cantDocentes; h++)
                {
                    int docId = ((s * 3 + h) % TotalDocentesEsperados) + 1;
                    habilitados.Add(new TerceroAsignatura
                    {
                        IdTercero = docId,
                        IdAsignatura = s,
                        Activo = 1,
                        Prioridad = (h == 0) ? 1 : 2,
                        VecesAsignada = 0
                    });
                }
                contexto.PoolPorAsignatura[s] = habilitados;
            }

            contexto.ConstruirIndices();
            return contexto;
        }

        public static AgIndividuo CrearIndividuo105(AgContexto contexto, int seed = 105)
        {
            var rng = new Random(seed);
            var ind = new AgIndividuo();

            for (int g = 1; g <= TotalGenesEsperados; g++)
            {
                int aulaId = ((g - 1) % TotalAulasEsperadas) + 1;
                int asigId = ((g - 1) % TotalAsignaturasEsperadas) + 1;
                int bloqueId = ((g - 1) % TotalBloquesEsperados) + 1;

                int docId = 0;
                int esVacante = 1;
                string motivo = "VACANTE_ESTRUCTURAL";

                if (contexto.PoolPorAsignatura.TryGetValue(asigId, out var pool) && pool.Count > 0)
                {
                    docId = pool[rng.Next(pool.Count)].IdTercero;
                    esVacante = 0;
                    motivo = string.Empty;
                }

                ind.Genes.Add(new AgGen
                {
                    Id = g,
                    IdIndividuo = 1,
                    IdAula = aulaId,
                    IdAsignatura = asigId,
                    IdBloqueHorario = bloqueId,
                    IdTercero = docId,
                    EsVacante = esVacante,
                    MotivoVacante = motivo,
                    HorasSemanales = 4,
                    HorasGen = 4,
                    TotalHorasAsignatura = 48,
                    FechaInicioEstimada = new DateTime(2026, 2, 2).AddDays((g % 14)),
                    FechaFinEstimada = new DateTime(2026, 6, 1).AddDays(-(g % 10)),
                    DentroDeVentana = 1,
                    Estado = AgEstadoGen.Valido,
                    ClaveGrupo = $"GRUPO_{aulaId}_{asigId}"
                });
            }

            return ind;
        }

        public static string CalcularHashIndividuo(AgIndividuo individuo)
        {
            if (individuo == null || individuo.Genes == null) return string.Empty;

            var sb = new StringBuilder();
            foreach (var g in individuo.Genes.OrderBy(x => x.Id))
            {
                sb.Append($"{g.Id}:{g.IdAula}:{g.IdAsignatura}:{g.IdBloqueHorario}:{g.IdTercero}:{g.EsVacante};");
            }

            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        private static string DiaNombre(int dia)
        {
            switch (dia)
            {
                case 1: return "Lunes";
                case 2: return "Martes";
                case 3: return "Miércoles";
                case 4: return "Jueves";
                case 5: return "Viernes";
                case 6: return "Sábado";
                default: return "Domingo";
            }
        }
    }
}
