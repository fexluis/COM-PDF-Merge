using System;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO que representa una rama completa seleccionada en el árbol jerárquico de 6 niveles:
    /// Periodo -> Sede -> Programa -> Semestre -> Grupo -> Modalidad.
    /// Cada instancia de este DTO materializa una tupla de contexto S para la generación de aulas.
    /// </summary>
    public class AgContextoSeleccionDto
    {
        public int IdPeriodo { get; set; }
        public int IdSede { get; set; }
        public int IdPrograma { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public string Modalidad { get; set; }
        public int Capacidad { get; set; }
        public string TipoAula { get; set; }
        public string CaracteristicasJson { get; set; }
        public string Codigo { get; set; }
        public string Nombre { get; set; }
        public int EsVirtual { get; set; }
        public string Jornada { get; set; } = string.Empty;
        public int HorasDia { get; set; }
        public string DiasVentana { get; set; } = string.Empty;

        public AgContextoSeleccionDto()
        {
            Semestre = string.Empty;
            Grupo = "A";
            Modalidad = ModalidadTutoriaCatalog.Presencial;
            TipoAula = "ESTANDAR";
            CaracteristicasJson = "{}";
            Codigo = string.Empty;
            Nombre = string.Empty;
            Jornada = string.Empty;
            HorasDia = 0;
            DiasVentana = string.Empty;
        }

        public AgContextoSeleccionDto(int idPeriodo, int idSede, int idPrograma, string semestre, string grupo, string modalidad)
        {
            IdPeriodo = idPeriodo;
            IdSede = idSede;
            IdPrograma = idPrograma;
            Semestre = semestre ?? string.Empty;
            Grupo = string.IsNullOrWhiteSpace(grupo) ? "A" : grupo.Trim().ToUpperInvariant();
            Modalidad = ModalidadTutoriaCatalog.NormalizarTutoria(modalidad);
        }

        /// <summary>
        /// Modalidad efectiva calculada: si EsVirtual > 0, se fuerza a "Plataforma".
        /// De lo contrario, se retorna la Modalidad normalizada (por defecto "Presencial") (Fases 237 y 785).
        /// </summary>
        public string ModalidadEfectiva => EsVirtual > 0 
            ? ModalidadTutoriaCatalog.Plataforma 
            : ModalidadTutoriaCatalog.NormalizarTutoria(Modalidad);

        /// <summary>
        /// Clave natural unívoca de la tupla S (AulaRef).
        /// Incluye ModalidadEfectiva para evitar colisiones entre variantes presenciales y virtuales del mismo grupo.
        /// Utiliza AgAulaFactory.ConstruirAulaRef para garantizar sincronía total con la entidad de aula (Fases 231 y 237).
        /// </summary>
        public string AulaRef => AgAulaFactory.ConstruirAulaRef(IdPeriodo, IdSede, IdPrograma, Semestre, Grupo, ModalidadEfectiva);

        public override string ToString() => AulaRef;
    }
}
