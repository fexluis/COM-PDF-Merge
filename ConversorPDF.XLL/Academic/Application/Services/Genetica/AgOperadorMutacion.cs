using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Operador de mutación para introducir diversidad estocástica en la población del AG.
    /// </summary>
    public class AgOperadorMutacion
    {
        private readonly Random _random;

        public AgOperadorMutacion(Random random = null)
        {
            _random = random ?? new Random();
        }

        public void Mutar(AgIndividuo individuo, double tasaMutacion, AgContexto contexto)
        {
            if (individuo == null) throw new ArgumentNullException(nameof(individuo));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            var aulasValidas = contexto.Aulas.Where(a => contexto.PuedeGenerarAula(a.Id)).ToList();
            var bloquesValidos = contexto.Bloques.Where(b => b.Activo == 1).ToList();

            // Fase 508: Docente/vacante coherente por ClaveGrupo (no por gen)
            var grupos = individuo.Genes
                .Where(g => g.Estado != AgEstadoGen.Fijado && (contexto.ManifiestoCobertura == null || contexto.PuedeGenerarAula(g.IdAula)))
                .GroupBy(g => string.IsNullOrWhiteSpace(g.ClaveGrupo) ? $"{g.IdAsignatura}|{g.IdAula}|{g.Grupo}" : g.ClaveGrupo)
                .ToList();

            foreach (var grp in grupos)
            {
                var genes = grp.ToList();
                int docCanon = genes[0].IdTercero;
                int vacCanon = genes[0].EsVacante;
                string motCanon = genes[0].MotivoVacante ?? string.Empty;

                // Fase 519, 533 & 1815: Selección y aplicación grupal con verificación de disponibilidad fechada.
                // Si ningún candidato alterno resulta compatible con el grupo y sus fechas, se conserva la asignación existente.
                if (_random.NextDouble() < tasaMutacion
                    && vacCanon == 0
                    && contexto.PoolPorAsignatura.TryGetValue(genes[0].IdAsignatura, out var poolAsig))
                {
                    var habilitados = poolAsig
                        .Where(p => p.Activo == 1 && (contexto.IdTerceroVacante <= 0 || p.IdTercero != contexto.IdTerceroVacante))
                        .ToList();
                    if (habilitados.Count > 0)
                    {
                        var orden = habilitados
                            .OrderBy(p => p.VecesAsignada)
                            .ThenBy(p => p.IdTercero)
                            .ToList();

                        foreach (var cand in orden)
                        {
                            if (cand.IdTercero == docCanon) continue;
                            if (AgAsignadorTercero.DisponibleGrupo(cand.IdTercero, genes[0], individuo, contexto))
                            {
                                docCanon = cand.IdTercero;
                                vacCanon = 0;
                                motCanon = string.Empty;
                                break;
                            }
                        }
                        // Si ningún candidato alterno resulta compatible, se conserva docCanon previo
                    }
                }
                else if (vacCanon == 1 && contexto.IdTerceroVacante > 0)
                {
                    docCanon = contexto.IdTerceroVacante;
                }

                // Fase 519 & 1815: sellar TODOS los genes del ClaveGrupo (incondicional)
                foreach (var g in genes)
                {
                    g.IdTercero = docCanon;
                    g.EsVacante = vacCanon;
                    g.MotivoVacante = vacCanon == 1 ? motCanon : string.Empty;
                }
            }

            foreach (var gen in individuo.Genes)
            {
                if (gen.Estado == AgEstadoGen.Fijado || (contexto.ManifiestoCobertura != null && !contexto.PuedeGenerarAula(gen.IdAula))) continue;
                if (gen.EsFijoTemporal) continue;   // F658: congela aula y bloque; docente ya mutó en bucle L28-73

                // Mutación de Aula — restringida al mismo grupo (Fase 456)
                if (_random.NextDouble() < tasaMutacion && aulasValidas.Count > 0)
                {
                    if (contexto.ManifiestoCobertura != null)
                    {
                        // La identidad de relación/aula está sellada por el manifiesto.
                    }
                    else if (contexto.IdAulaCorrida > 0)
                    {
                        gen.IdAula = contexto.IdAulaCorrida;
                    }
                    else
                    {
                        // Fase 456, Fase 486 & Fase 664: Restringir candidatas al mismo grupo, semestre y programa de la cohorte
                        var candidatas = aulasValidas.Where(a =>
                                string.Equals(a.Grupo ?? "A", gen.Grupo ?? "A", StringComparison.OrdinalIgnoreCase)
                             && string.Equals(AgFormatoAsignatura.SemestreRomano(a.Semestre),
                                              AgFormatoAsignatura.SemestreRomano(gen.Semestre),
                                              StringComparison.OrdinalIgnoreCase)
                             && (a.IdPrograma <= 0 || gen.IdPrograma <= 0 || a.IdPrograma == gen.IdPrograma))
                            .ToList();
                        if (candidatas.Count > 0)
                        {
                            var elegida = candidatas[_random.Next(candidatas.Count)];
                            gen.IdAula = elegida.Id;
                            string semNorm = AgFormatoAsignatura.SemestreRomano(elegida.Semestre);
                            if (semNorm != "SIN ASIGNAR")
                            {
                                gen.Semestre = semNorm;
                            }
                        }
                    }
                }

                // Mutación de Bloque Horario (Fase 456 / Fase 1060 / Fase 1114): preserva duración exacta de nivel y días del carril
                bool multiDia = !string.IsNullOrWhiteSpace(gen.ClaveGrupo) && !string.IsNullOrWhiteSpace(gen.PatronSemanal);
                // Fase 1114: duración objetivo según nivel del aula (D4'), nunca la del gen
                double duracionObjetivo;
                AgAula aulaMut = null;
                if (gen.IdAula > 0 && contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(gen.IdAula, out aulaMut))
                {
                    var progMut = contexto.Programas?.FirstOrDefault(p => p.Id == aulaMut.IdPrograma);
                    duracionObjetivo = AgUnidadesSemanales.HorasSesionAula(aulaMut, AgNivelAcademicoCatalogo.Desde(progMut?.Modalidad));
                }
                else
                {
                    duracionObjetivo = 4.0;
                }

                // Fase 1132 / Fase 1168: resolver franjas permitidas según AgRanuras.FranjasDe
                string franjaActual = (contexto?.BloquesPorId != null && gen.IdBloqueHorario > 0
                    && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bAct)) ? bAct.Franja : null;
                var franjasPermitidas = AgRanuras.FranjasDe(gen.DiaSemana, aulaMut?.Jornada, duracionObjetivo);

                if (!multiDia && _random.NextDouble() < tasaMutacion && bloquesValidos.Count > 0)
                {
                    var candidatos = bloquesValidos.Where(b =>
                        Math.Abs(b.DuracionHoras - duracionObjetivo) < 0.001 &&
                        (franjasPermitidas.Count == 0 || franjasPermitidas.Any(f => string.Equals(f, b.Franja, StringComparison.OrdinalIgnoreCase))))
                        .OrderByDescending(b => string.Equals(b.Franja, franjaActual, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (candidatos.Count == 0)
                    {
                        candidatos = bloquesValidos.Where(b => Math.Abs(b.DuracionHoras - duracionObjetivo) < 0.001).ToList();
                    }

                    if (candidatos.Count > 0)
                    {
                        var bloqueNuevo = candidatos[_random.Next(candidatos.Count)];
                        gen.AsignarBloque(bloqueNuevo);
                    }
                }
                else if (multiDia && _random.NextDouble() < tasaMutacion && gen.DiaSemana >= 1)
                {
                    var permitidos = !string.IsNullOrWhiteSpace(gen.DiasPatron)
                        ? gen.DiasPatron.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(s => int.TryParse(s, out int d) ? d : 0)
                                        .Where(d => d >= 1 && d <= 7).ToList()
                        : new List<int> { gen.DiaSemana };

                    if (permitidos.Contains(gen.DiaSemana))
                    {
                        var candidatosMismoDia = bloquesValidos
                            .Where(b => b.DiaSemanaNumero == gen.DiaSemana &&
                                        Math.Abs(b.DuracionHoras - duracionObjetivo) < 0.001 &&
                                        (franjasPermitidas.Count == 0 || franjasPermitidas.Any(f => string.Equals(f, b.Franja, StringComparison.OrdinalIgnoreCase))))
                            .OrderByDescending(b => string.Equals(b.Franja, franjaActual, StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (candidatosMismoDia.Count == 0)
                        {
                            candidatosMismoDia = bloquesValidos
                                .Where(b => b.DiaSemanaNumero == gen.DiaSemana && Math.Abs(b.DuracionHoras - duracionObjetivo) < 0.001)
                                .ToList();
                        }

                        if (candidatosMismoDia.Count > 0)
                        {
                            gen.AsignarBloque(candidatosMismoDia[_random.Next(candidatosMismoDia.Count)]);
                        }
                    }
                }
            }

            AgCoherenciaGrupo.Sellar(individuo, contexto); // Fase 521
        }
    }
}
