using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Contexto global en memoria con catálogos, índices rápidos y parámetros de la optimización del AG.
    /// Diseñado para búsquedas O(1) de alta velocidad durante la evaluación cromosómica.
    /// </summary>
    public class AgContexto
    {
        /// <summary>Manifiesto explícito opcional de preparación; null conserva los flujos anteriores.
        /// No habilita conservación, generación ni autorización por sí mismo.</summary>
        public AgManifiestoCobertura ManifiestoCobertura { get; set; }

        /// <summary>Admite genes únicamente para relaciones seleccionadas y replanificación ratificada;
        /// sin manifiesto preserva el filtro histórico Disponible.</summary>
        public bool PuedeGenerarAula(int idAula)
        {
            var aula = Aulas?.SingleOrDefault(a => a.Id == idAula);
            if (aula == null) return false;
            if (ManifiestoCobertura == null) return aula.Estado == ConversorPDF.XLL.Academic.Domain.Enums.AgEstadoAula.Disponible;
            if (!ManifiestoCobertura.Relaciones.Any(r => r.Clave.IdAula == idAula)) return false;
            if (aula.Estado == ConversorPDF.XLL.Academic.Domain.Enums.AgEstadoAula.Disponible) return true;
            // Regla general: el estado Aprobado sólo puede conservar lo que existe. Sin ninguna fila
            // productiva para la cohorte no hay origen que preservar y el aula es demanda obligatoria.
            // Requiere snapshot no nulo para acreditar ausencia comprobada; snapshot nulo nunca autoriza (fail-closed).
            if (aula.Estado == ConversorPDF.XLL.Academic.Domain.Enums.AgEstadoAula.Aprobado
                && PlanAprobadoSnapshot != null
                && !TieneOrigenConservable(idAula)) return true;
            var c = CandidatosCobertura.SingleOrDefault(x => x.IdAula == idAula);
            return aula.Estado == ConversorPDF.XLL.Academic.Domain.Enums.AgEstadoAula.Aprobado
                && c != null && c.Desde != c.Hacia
                && c.HuellaManifiesto == ManifiestoCobertura.Huella
                && c.HuellaOrigen == PlanAprobadoSnapshot?.HuellaOrigen
                && c.Inicio == ManifiestoCobertura.Inicio && c.Fin == ManifiestoCobertura.Fin;
        }

        /// <summary>Indica si el aula tiene alguna fila productiva reconciliada (requerida o extra) en el
        /// snapshot leído. Es la única condición que habilita conservar, comparar o ratificar un origen;
        /// un aula aprobada sin origen se planifica como demanda nueva.</summary>
        public bool TieneOrigenConservable(int idAula)
        {
            var snapshot = PlanAprobadoSnapshot;
            if (snapshot == null) return false;
            return snapshot.Relaciones.Concat(snapshot.RelacionesExtra)
                .Any(r => r.ClavesCandidatas.Any(k => k.IdAula == idAula));
        }

        /// <summary>Proyección fija de sesiones conservadas, nunca genes ni filas de staging.</summary>
        public IReadOnlyCollection<AgSesionCobertura> SesionesConservadas { get; set; } =
            Array.Empty<AgSesionCobertura>();

        /// <summary>Origen canónico capturado antes de cambios; null si no hay aulas aprobadas.</summary>
        public AgPlanAprobadoSnapshot PlanAprobadoSnapshot { get; set; }

        /// <summary>Decisión candidata aislada; no modifica ag_aula ni concede generación/promoción T5/T6.</summary>
        public AgCandidatoPlanAprobado CandidatoRatificacion { get; set; }

        /// <summary>Decisiones publicadas por aula de la corrida, sin sesiones productivas mutables.</summary>
        public IReadOnlyCollection<AgCandidatoPlanAprobado> CandidatosCobertura { get; private set; } =
            Array.Empty<AgCandidatoPlanAprobado>();

        /// <summary>Descarta decisiones de una preparación anterior al reconstruir el contexto.</summary>
        public void ReiniciarCandidatosCobertura()
        {
            CandidatosCobertura = Array.Empty<AgCandidatoPlanAprobado>();
            CandidatoRatificacion = null;
        }

        /// <summary>Registra una decisión única por aula; no sustituye la prueba de publicación del confirmador.</summary>
        public void RegistrarCandidatoCobertura(AgCandidatoPlanAprobado candidato)
        {
            if (candidato == null || CandidatosCobertura.Any(c => c.IdAula == candidato.IdAula))
                throw new InvalidOperationException("CANDIDATO_AULA_DUPLICADO");
            CandidatosCobertura = CandidatosCobertura.Concat(new[] { candidato }).ToList().AsReadOnly();
            CandidatoRatificacion = candidato;
        }

        /// <summary>Diagnóstico ligado al manifiesto; no concede ratificación ni autorización T4.</summary>
        public AgValidacionPlanAprobado ValidacionPlanAprobado { get; set; }

        public int IdPeriodo { get; set; }
        public int IdPrograma { get; set; }
        public int IdSede { get; set; }
        public int IdAulaCorrida { get; set; }
        public int IdTerceroVacante { get; set; }
        public string Subunidad { get; set; }
        public string SemestreActual { get; set; } = string.Empty;
        public string GrupoActual { get; set; } = string.Empty;
        public VentanaSemestre Ventana { get; set; }
        public AgPesosFitness Pesos { get; set; }
        /// <summary> Instancia de traza diagnóstica asociada al contexto de ejecución (Fase 481). </summary>
        public AgDiagnosticoTraza Traza { get; set; }
        /// <summary> Identificador de corrida activa para delimitación de alcance de diagnóstico (Fase 585). </summary>
        public long IdCorrida { get; set; }

        /// <summary> Identificadores de sedes seleccionadas para ejecuciones multi-sede (Fase 525). </summary>
        public List<int> SedesAlcance { get; set; } = new List<int>();

        /// <summary> Identificadores de programas seleccionados para ejecuciones multi-programa (Fase 525). </summary>
        public List<int> ProgramasAlcance { get; set; } = new List<int>();

        public List<AgAula> Aulas { get; set; }
        public List<AgAula> AulasAlcance { get; set; }
        public List<AgBloqueHorario> Bloques { get; set; }
        public List<TerceroAsignatura> PoolDocentes { get; set; }
        public List<Tercero> Docentes { get; set; }
        public List<Asignatura> Asignaturas { get; set; }
        public List<Programa> Programas { get; set; }
        public List<Sede> Sedes { get; set; }
        public List<AgAulaAsignatura> AulaAsignaturas { get; set; }
        public List<(DateTime Inicio, DateTime Fin)> Festivos { get; set; }
        public List<(DateTime Inicio, DateTime Fin)> Presenciales { get; set; }

        private HashSet<int> _vacantesConsentidas = new HashSet<int>();
        /// <summary>
        /// FASE 1871/1877: Asignaturas autorizadas por el usuario para admitir docente VACANTE sin fallar el certificado de cobertura.
        /// </summary>
        public HashSet<int> AsignaturasVacantesConsentidas
        {
            get => _vacantesConsentidas ?? (_vacantesConsentidas = new HashSet<int>());
            set => _vacantesConsentidas = value ?? new HashSet<int>();
        }
        public bool PermiteVacantes => AsignaturasVacantesConsentidas != null && AsignaturasVacantesConsentidas.Count > 0;

        // Índices O(1) para evaluación rápida
        public Dictionary<int, AgAula> AulasPorId { get; private set; }
        public Dictionary<int, AgBloqueHorario> BloquesPorId { get; private set; }
        public Dictionary<int, Asignatura> AsignaturasPorId { get; private set; }
        public Dictionary<int, Programa> ProgramasPorId { get; private set; }
        public Dictionary<int, double> UnidadHorasPorAulaId { get; private set; }
        public Dictionary<int, AgAulaAsignatura> AulaAsignaturasPorId { get; private set; }
        /// <summary> Índice O(1) de asignaturas por aula y asignatura para soporte multi-aula y multi-sede (Fase 547). </summary>
        public Dictionary<(int IdAula, int IdAsignatura), AgAulaAsignatura> AulaAsignaturaPorAulaAsig { get; private set; }
        public Dictionary<int, Tercero> DocentesPorId { get; private set; }
        public Dictionary<int, Sede> SedesPorId { get; private set; }
        public HashSet<long> PoolDocenteKeys { get; private set; }
        public Dictionary<int, List<TerceroAsignatura>> PoolPorAsignatura { get; set; }
        public Dictionary<int, List<(DateTime Fecha, FranjaHoraria Franja)>> OcupacionExterna { get; private set; }
        private HashSet<DateTime> _diasFestivos;
        private IReadOnlyCollection<(long IdSesion, int IdTercero, DateTime Fecha, FranjaHoraria Franja)> _ocupacionesCanonicas =
            Array.Empty<(long, int, DateTime, FranjaHoraria)>();
        private IReadOnlyCollection<(int IdTercero, DateTime Fecha, FranjaHoraria Franja)> _ocupacionesStaging =
            Array.Empty<(int, DateTime, FranjaHoraria)>();

        public AgContexto()
        {
            Pesos = AgPesosFitness.PorDefecto();
            Aulas = new List<AgAula>();
            AulasAlcance = new List<AgAula>();
            Bloques = new List<AgBloqueHorario>();
            PoolDocentes = new List<TerceroAsignatura>();
            Docentes = new List<Tercero>();
            Asignaturas = new List<Asignatura>();
            Programas = new List<Programa>();
            Sedes = new List<Sede>();
            AulaAsignaturas = new List<AgAulaAsignatura>();
            Festivos = new List<(DateTime Inicio, DateTime Fin)>();
            Presenciales = new List<(DateTime Inicio, DateTime Fin)>();

            AulasPorId = new Dictionary<int, AgAula>();
            BloquesPorId = new Dictionary<int, AgBloqueHorario>();
            AsignaturasPorId = new Dictionary<int, Asignatura>();
            ProgramasPorId = new Dictionary<int, Programa>();
            UnidadHorasPorAulaId = new Dictionary<int, double>();
            AulaAsignaturasPorId = new Dictionary<int, AgAulaAsignatura>();
            AulaAsignaturaPorAulaAsig = new Dictionary<(int IdAula, int IdAsignatura), AgAulaAsignatura>();
            DocentesPorId = new Dictionary<int, Tercero>();
            SedesPorId = new Dictionary<int, Sede>();
            PoolDocenteKeys = new HashSet<long>();
            PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>();
            OcupacionExterna = new Dictionary<int, List<(DateTime, FranjaHoraria)>>();
            _diasFestivos = new HashSet<DateTime>();
            Subunidad = string.Empty;
        }

        public void ConstruirIndices()
        {
            AulasPorId = Aulas.ToDictionary(a => a.Id, a => a);
            BloquesPorId = Bloques.ToDictionary(b => b.Id, b => b);
            AsignaturasPorId = Asignaturas.ToDictionary(a => a.Id, a => a);
            DocentesPorId = Docentes.ToDictionary(d => d.Id, d => d);
            SedesPorId = Sedes.ToDictionary(s => s.Id, s => s);
            ProgramasPorId = Programas != null ? Programas.ToDictionary(p => p.Id, p => p) : new Dictionary<int, Programa>();

            UnidadHorasPorAulaId = new Dictionary<int, double>();
            if (Aulas != null)
            {
                foreach (var a in Aulas)
                {
                    var prog = (ProgramasPorId != null && ProgramasPorId.TryGetValue(a.IdPrograma, out var p)) ? p : null;
                    var nivel = AgNivelAcademicoCatalogo.Desde(prog?.Modalidad);
                    UnidadHorasPorAulaId[a.Id] = AgUnidadesSemanales.HorasSesionAula(a, nivel);
                }
            }

            AulaAsignaturasPorId = new Dictionary<int, AgAulaAsignatura>();
            AulaAsignaturaPorAulaAsig = AulaAsignaturas != null
                ? AulaAsignaturas
                    .GroupBy(x => (x.IdAula, x.IdAsignatura))
                    .ToDictionary(g => g.Key, g => g.FirstOrDefault())
                : new Dictionary<(int IdAula, int IdAsignatura), AgAulaAsignatura>();

            if (AulaAsignaturas != null)
            {
                foreach (var aa in AulaAsignaturas)
                {
                    AulaAsignaturasPorId[aa.IdAsignatura] = aa;
                }
            }

            PoolDocenteKeys = new HashSet<long>(PoolDocentes.Where(p => p.Activo == 1).Select(p => p.Id));

            PoolPorAsignatura = new Dictionary<int, List<TerceroAsignatura>>();
            foreach (var item in PoolDocentes.Where(p => p.Activo == 1))
            {
                if (!PoolPorAsignatura.TryGetValue(item.IdAsignatura, out var lista))
                {
                    lista = new List<TerceroAsignatura>();
                    PoolPorAsignatura[item.IdAsignatura] = lista;
                }
                lista.Add(item);
            }

            // Priorizar candidatos del pool: primero los del semestre actual (exacto o romano/arábigo), luego por prioridad y carga
            string semRomanoActual = AgFormatoAsignatura.SemestreRomano(SemestreActual);
            foreach (var kvp in PoolPorAsignatura.ToList())
            {
                PoolPorAsignatura[kvp.Key] = kvp.Value
                    .OrderBy(p => (string.Equals(p.Semestre, SemestreActual, StringComparison.OrdinalIgnoreCase) ||
                                  (!string.IsNullOrWhiteSpace(semRomanoActual) && semRomanoActual != "SIN ASIGNAR" &&
                                   string.Equals(AgFormatoAsignatura.SemestreRomano(p.Semestre), semRomanoActual, StringComparison.OrdinalIgnoreCase))) ? 0 : 1)
                    .ThenBy(p => p.Prioridad)
                    .ThenBy(p => p.VecesAsignada)
                    .ToList();
            }

            _diasFestivos = new HashSet<DateTime>();
            foreach (var f in Festivos)
            {
                for (var d = f.Inicio.Date; d <= f.Fin.Date; d = d.AddDays(1))
                {
                    _diasFestivos.Add(d);
                }
            }

            // Fase 1125 / Fase 1128: Invalidar índice de bloques ante reconstrucción
            InvalidarIndiceBloques();
        }

        // Fase 1125 / Fase 1127: Caché indexado de bloques por (DiaSemanaNumero, DuracionHoras) -> lista de bloques
        private Dictionary<(int, double), List<AgBloqueHorario>> _indiceBloquesDiaDur;

        /// <summary>
        /// Invalida el caché indexado de bloques cuando cambia la colección Bloques (Fase 1128).
        /// </summary>
        public void InvalidarIndiceBloques() => _indiceBloquesDiaDur = null;

        /// <summary>
        /// Busca un bloque horario activo de forma determinista en O(1) por día y duración,
        /// priorizando la franja preferida del aula (Fase 1125, Fase 1127).
        /// </summary>
        public bool TryGetBloque(int dia, double dur, string franjaPreferida, out AgBloqueHorario b)
        {
            b = null;
            if (_indiceBloquesDiaDur == null)
            {
                _indiceBloquesDiaDur = (Bloques ?? new List<AgBloqueHorario>())
                    .Where(x => x.Activo == 1)
                    .GroupBy(x => (x.DiaSemanaNumero, Math.Round(x.DuracionHoras, 3)))
                    .ToDictionary(g => g.Key, g => g.ToList());
            }

            if (!_indiceBloquesDiaDur.TryGetValue((dia, Math.Round(dur, 3)), out var lista) || lista == null || lista.Count == 0)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(franjaPreferida))
            {
                b = lista.FirstOrDefault(x => string.Equals(x.Franja, franjaPreferida, StringComparison.OrdinalIgnoreCase));
            }

            if (b == null)
            {
                b = lista[0];
            }

            return b != null;
        }

        /// <summary>
        /// Sobrecarga de compatibilidad sin franja preferida (Fase 1127).
        /// </summary>
        public bool TryGetBloque(int dia, double dur, out AgBloqueHorario b)
            => TryGetBloque(dia, dur, null, out b);

        public bool EsDocenteHabilitado(int idTercero, int idPrograma, int idAsignatura, string semestre)
        {
            if (idTercero <= 0) return false;

            // 1. Clave exacta con el semestre recibido
            long key = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre ?? string.Empty);
            if (PoolDocenteKeys.Contains(key)) return true;

            // 2. Clave con semestre normalizado a Romano (F378 / F399)
            string semRomano = AgFormatoAsignatura.SemestreRomano(semestre);
            if (!string.IsNullOrWhiteSpace(semRomano) && semRomano != "SIN ASIGNAR" && semRomano != semestre)
            {
                long keyRomano = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semRomano);
                if (PoolDocenteKeys.Contains(keyRomano)) return true;
            }

            // 3. Clave con semestre normalizado a Arábigo (F378 / F399)
            string semArabigo = AgFormatoAsignatura.SemestreArabigo(semestre);
            if (!string.IsNullOrWhiteSpace(semArabigo) && semArabigo != semestre)
            {
                long keyArabigo = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semArabigo);
                if (PoolDocenteKeys.Contains(keyArabigo)) return true;
            }

            // 4. Revisar si existe habilitación genérica (semestre comodín "")
            long wildcardKey = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, string.Empty);
            return PoolDocenteKeys.Contains(wildcardKey);
        }

        public void RegistrarDocenteHabilitado(int idTercero, int idPrograma, int idAsignatura, string semestre = "")
        {
            long key = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre ?? string.Empty);
            PoolDocenteKeys.Add(key);
            PoolDocentes.Add(new TerceroAsignatura
            {
                Id = key,
                IdTercero = idTercero,
                IdPrograma = idPrograma,
                IdAsignatura = idAsignatura,
                Semestre = semestre ?? string.Empty,
                Activo = 1
            });
        }

        public void CargarOcupacionExterna(IEnumerable<(int IdTercero, DateTime Fecha, FranjaHoraria Franja)> ocupaciones)
        {
            OcupacionExterna = new Dictionary<int, List<(DateTime Fecha, FranjaHoraria Franja)>>();
            if (ocupaciones == null) return;

            foreach (var o in ocupaciones)
            {
                if (o.IdTercero <= 0) continue;
                if (!OcupacionExterna.TryGetValue(o.IdTercero, out var l))
                {
                    l = new List<(DateTime Fecha, FranjaHoraria Franja)>();
                    OcupacionExterna[o.IdTercero] = l;
                }
                l.Add((o.Fecha.Date, o.Franja));
            }
        }

        /// <summary>Captura ocupaciones externas con identidad canónica sin duplicar sesiones de staging.</summary>
        public void CargarOcupacionesIdentificadas(
            IEnumerable<(long IdSesion, int IdTercero, DateTime Fecha, FranjaHoraria Franja)> canonicas,
            IEnumerable<(int IdTercero, DateTime Fecha, FranjaHoraria Franja)> staging)
        {
            _ocupacionesCanonicas = canonicas.ToList().AsReadOnly();
            _ocupacionesStaging = staging.ToList().AsReadOnly();
            ExcluirOcupacionesCanonicas(Array.Empty<long>());
        }

        /// <summary>Excluye sólo sesiones canónicas seleccionadas; las restantes siguen bloqueando al docente.</summary>
        public void ExcluirOcupacionesCanonicas(IEnumerable<long> idsSeleccionados)
        {
            var ids = new HashSet<long>(idsSeleccionados);
            CargarOcupacionExterna(_ocupacionesCanonicas.Where(o => !ids.Contains(o.IdSesion))
                .Select(o => (o.IdTercero, o.Fecha, o.Franja)).Concat(_ocupacionesStaging));
        }

        public bool TerceroOcupado(int idTercero, DateTime fecha, FranjaHoraria franja)
        {
            if (idTercero <= 0) { return false; }
            if (OcupacionExterna == null || !OcupacionExterna.TryGetValue(idTercero, out var lista)) { return false; }
            foreach (var o in lista)
            {
                if (o.Fecha.Date == fecha.Date && o.Franja.SeSolapaCon(franja)) { return true; }
            }
            return false;
        }

        public bool EsFestivo(DateTime fecha) => _diasFestivos != null && _diasFestivos.Contains(fecha.Date);

        /// <summary>
        /// Filtra el pool docente y la lista de docentes por un conjunto de identificadores permitidos.
        /// Preserva siempre el centinela vacante si existe.
        /// </summary>
        public void FiltrarPoolDocentesPorTerceros(IEnumerable<int> idsPermitidos)
        {
            if (idsPermitidos == null) return;
            var set = new HashSet<int>(idsPermitidos);
            if (set.Count == 0) return;

            if (IdTerceroVacante > 0)
            {
                set.Add(IdTerceroVacante);
            }

            if (PoolDocentes != null)
            {
                PoolDocentes = PoolDocentes.Where(p => set.Contains(p.IdTercero)).ToList();
            }

            if (Docentes != null)
            {
                Docentes = Docentes.Where(d => set.Contains(d.Id)).ToList();
            }

            if (Traza != null)
            {
                Traza.Dato("TercerosFiltradosEnPool", $"Docentes={Docentes?.Count ?? 0}, Pool={PoolDocentes?.Count ?? 0}");
            }
        }
    }
}
