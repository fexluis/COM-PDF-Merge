using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Identidad completa de demanda; el aula no se deduce de una asignatura o semestre.</summary>
    public sealed class AgClaveRelacionCobertura
    {
        /// <summary>Territorial explícita de la selección.</summary>
        public int IdTerritorial { get; }
        /// <summary>Partición territorial, nunca comodín vacío.</summary>
        public string Subunidad { get; }
        /// <summary>Período seleccionado.</summary>
        public int IdPeriodo { get; }
        /// <summary>Sede seleccionada.</summary>
        public int IdSede { get; }
        /// <summary>Programa seleccionado.</summary>
        public int IdPrograma { get; }
        /// <summary>Aula AG seleccionada, ausente del esquema productivo.</summary>
        public int IdAula { get; }
        /// <summary>Asignatura requerida.</summary>
        public int IdAsignatura { get; }
        /// <summary>Semestre canónico de la cohorte.</summary>
        public string Semestre { get; }
        /// <summary>Grupo canónico de la cohorte.</summary>
        public string Grupo { get; }
        /// <summary>Modalidad canónica, sin inferencia desde el estado Aprobado.</summary>
        public string Modalidad { get; }
        /// <summary>Representación completa no ambigua para comparación ordinal.</summary>
        public string Valor { get; }

        /// <summary>Crea una clave con semestre previamente normalizado por el constructor de manifiesto.</summary>
        public AgClaveRelacionCobertura(int idTerritorial, string subunidad, int idPeriodo, int idSede,
            int idPrograma, int idAula, int idAsignatura, string semestre, string grupo, string modalidad)
        {
            if (new[] { idTerritorial, idPeriodo, idSede, idPrograma, idAula, idAsignatura }.Any(i => i <= 0)
                || new[] { subunidad, semestre, grupo, modalidad }.Any(string.IsNullOrWhiteSpace))
                throw new ArgumentException("IDENTIDAD_INCOMPLETA");
            IdTerritorial = idTerritorial; Subunidad = Texto(subunidad); IdPeriodo = idPeriodo;
            IdSede = idSede; IdPrograma = idPrograma; IdAula = idAula; IdAsignatura = idAsignatura;
            Semestre = Texto(semestre); Grupo = Texto(grupo);
            Modalidad = Texto(ModalidadTutoriaCatalog.NormalizarTutoria(modalidad));
            Valor = AgOrigenCanonico.Campos(IdTerritorial, Subunidad, IdPeriodo, IdSede, IdPrograma,
                IdAula, IdAsignatura, Semestre, Grupo, Modalidad);
        }

        internal static string Texto(string s) => (s ?? string.Empty).Trim().ToUpperInvariant();
    }

    /// <summary>Demanda inmutable, previa a cualquier gen, staging o decisión de conservación.</summary>
    public sealed class AgRelacionCobertura
    {
        /// <summary>Identidad completa requerida.</summary>
        public AgClaveRelacionCobertura Clave { get; }
        /// <summary>Demanda sin redondear según la prioridad vigente.</summary>
        public double HorasOriginales { get; }
        /// <summary>Demanda de AgEvaluadorDemandaVentana.HorasObjetivo.</summary>
        public double HorasEfectivas { get; }
        /// <summary>Estado observado; no acredita aprobación productiva ni autorización.</summary>
        public int EstadoAula { get; }
        /// <summary>Crea una relación pendiente de revisión; no concede consentimiento.</summary>
        public AgRelacionCobertura(AgClaveRelacionCobertura clave, double horasOriginales, double horasEfectivas, int estadoAula)
        {
            Clave = clave ?? throw new ArgumentNullException(nameof(clave));
            if (!Valida(horasOriginales) || !Valida(horasEfectivas)) throw new ArgumentException("DEMANDA_INVALIDA");
            HorasOriginales = horasOriginales; HorasEfectivas = horasEfectivas; EstadoAula = estadoAula;
        }
        private static bool Valida(double h) => h > 0 && !double.IsInfinity(h) && !double.IsNaN(h);
    }

    /// <summary>Snapshot del alcance requerido independiente de genes y resultados.</summary>
    public sealed class AgManifiestoCobertura
    {
        /// <summary>Versión del contrato serializable.</summary>
        public int Version => 1;
        /// <summary>Inicio fijo de la ventana.</summary>
        public DateTime Inicio { get; }
        /// <summary>Fin fijo, sin ampliación implícita.</summary>
        public DateTime Fin { get; }
        /// <summary>Relaciones requeridas ordenadas, sin duplicados.</summary>
        public IReadOnlyList<AgRelacionCobertura> Relaciones { get; }
        /// <summary>Claves productivas declaradas no requeridas (extras): se conservan inmutables y ocupan
        /// aula y docente, pero no suman demanda requerida ni se replantean.</summary>
        public IReadOnlyList<AgClaveRelacionCobertura> RelacionesExtra { get; }
        /// <summary>Huella de identidad, demanda original/efectiva, estados, ventana y extras declaradas.</summary>
        public string Huella { get; }
        private readonly HashSet<string> _extras;
        /// <summary>Copia y valida el manifiesto; nunca completa selecciones faltantes por fallback.</summary>
        public AgManifiestoCobertura(DateTime inicio, DateTime fin, IEnumerable<AgRelacionCobertura> relaciones,
            IEnumerable<AgClaveRelacionCobertura> extras = null)
        {
            if (fin.Date < inicio.Date) throw new ArgumentException("VENTANA_INVALIDA");
            Inicio = inicio.Date; Fin = fin.Date;
            var lista = (relaciones ?? throw new ArgumentNullException(nameof(relaciones))).ToList();
            if (lista.Count == 0 || lista.Any(r => r == null) || lista.GroupBy(r => r.Clave.Valor).Any(g => g.Count() != 1))
                throw new ArgumentException("MANIFIESTO_VACIO_O_DUPLICADO");
            Relaciones = lista.OrderBy(r => r.Clave.Valor, StringComparer.Ordinal).ToList().AsReadOnly();
            // Decisión humana 2026-11-15 «B: extras como ocupación fija»: colección explícita y sellada.
            var declaradas = (extras ?? Enumerable.Empty<AgClaveRelacionCobertura>()).ToList();
            if (declaradas.Any(k => k == null) || declaradas.GroupBy(k => k.Valor).Any(g => g.Count() != 1)
                || declaradas.Any(k => Relaciones.Any(r => r.Clave.Valor == k.Valor)))
                throw new ArgumentException("EXTRAS_INVALIDAS");
            RelacionesExtra = declaradas.OrderBy(k => k.Valor, StringComparer.Ordinal).ToList().AsReadOnly();
            _extras = new HashSet<string>(RelacionesExtra.Select(k => k.Valor), StringComparer.Ordinal);
            Huella = AgOrigenCanonico.Huella(AgOrigenCanonico.Campos(Version, Inicio, Fin) +
                string.Concat(Relaciones.Select(r => AgOrigenCanonico.Campos(r.Clave.Valor,
                    r.HorasOriginales, r.HorasEfectivas, r.EstadoAula))) +
                string.Concat(RelacionesExtra.Select(k => AgOrigenCanonico.Campos("EXTRA", k.Valor))));
        }
        /// <summary>Indica si la clave está declarada como extra no requerida.</summary>
        public bool EsExtra(AgClaveRelacionCobertura clave) => clave != null && _extras.Contains(clave.Valor);
        /// <summary>Indica si el valor canónico está declarado como extra no requerida.</summary>
        public bool EsExtra(string valor) => !string.IsNullOrEmpty(valor) && _extras.Contains(valor);
        /// <summary>Detecta relaciones enteras omitidas aunque genes y staging coincidan entre sí.</summary>
        public IReadOnlyList<AgClaveRelacionCobertura> DetectarOmisiones(IEnumerable<AgClaveRelacionCobertura> presentes)
        {
            var claves = new HashSet<string>((presentes ?? throw new ArgumentNullException(nameof(presentes))).Select(k => k.Valor), StringComparer.Ordinal);
            return Relaciones.Where(r => !claves.Contains(r.Clave.Valor)).Select(r => r.Clave).ToList().AsReadOnly();
        }
    }

    // Codificación con longitudes y tipos: preserva null, contenido bruto y límites sin depender de cultura.
    internal static class AgOrigenCanonico
    {
        internal static string Campos(params object[] valores)
        {
            return string.Concat(valores.Select(v => {
                string s = v == null || v == DBNull.Value ? "null" : v is DateTime d
                    ? "date:" + d.ToString("O", CultureInfo.InvariantCulture) : v is byte[] bytes
                    ? "bytes:" + Convert.ToBase64String(bytes) : v is double n
                    ? "double:" + n.ToString("R", CultureInfo.InvariantCulture)
                    : v.GetType().FullName + ":" + Convert.ToString(v, CultureInfo.InvariantCulture);
                return s.Length.ToString(CultureInfo.InvariantCulture) + ":" + s;
            }));
        }
        internal static string Huella(string origen)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(origen))).Replace("-", string.Empty);
        }
    }
}
