using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Fábrica para materializar entidades AgAula a partir de AgContextoSeleccionDto.
    /// Garantiza que AulaRef contenga la Modalidad (evitando colisión de tuplas S) y que los campos canónicos queden sincronizados.
    /// </summary>
    public static class AgAulaFactory
    {
        /// <summary>
        /// Generador canónico de la clave natural AulaRef (idperiodo|idsede|idprograma|semestre|grupo|modalidad).
        /// Normaliza la modalidad al código canónico antes de convertir a mayúsculas invariantes (Fases 237 y 786).
        /// </summary>
        public static string ConstruirAulaRef(int idPeriodo, int idSede, int idPrograma, string semestre, string grupo, string modalidad)
        {
            string modNormalizada = ModalidadTutoriaCatalog.NormalizarTutoria(modalidad);
            return $"{idPeriodo}|{idSede}|{idPrograma}|{(semestre ?? "").Trim().ToUpperInvariant()}|{(grupo ?? "").Trim().ToUpperInvariant()}|{modNormalizada.Trim().ToUpperInvariant()}";
        }

        /// <summary>
        /// Extrae la clave de cohorte lógica (idperiodo|idsede|semestre|grupo|modalidad) omitiendo idprograma para permitir match con homónimos (Fase 645).
        /// </summary>
        public static string ExtraerCohorte(string aulaRef)
        {
            if (string.IsNullOrWhiteSpace(aulaRef)) return string.Empty;
            var p = aulaRef.Split('|');
            if (p.Length < 6) return string.Empty;
            return (p[0] + "|" + p[1] + "|" + p[3] + "|" + p[4] + "|" + p[5]).ToUpperInvariant();
        }

        public static AgAula CrearDesdeDto(AgContextoSeleccionDto dto, VentanaSemestre ventana = null, string subunidad = "")
        {
            if (dto == null) throw new ArgumentNullException(nameof(dto));

            string mod = ModalidadTutoriaCatalog.NormalizarTutoria(dto.Modalidad);
            int esVirtual = (dto.EsVirtual > 0 || ModalidadTutoriaCatalog.EsPlataforma(mod)) ? 1 : 0;
            string modVinculacion = esVirtual == 1 ? ModalidadTutoriaCatalog.Plataforma : mod;
            string grupo = string.IsNullOrWhiteSpace(dto.Grupo) ? "A" : dto.Grupo.Trim().ToUpperInvariant();
            string sem = string.IsNullOrWhiteSpace(dto.Semestre) ? "I" : dto.Semestre.Trim().ToUpperInvariant();

            int capacidad = dto.Capacidad > 0 ? dto.Capacidad : (esVirtual == 1 ? 999 : 30);
            string tipoAula = !string.IsNullOrWhiteSpace(dto.TipoAula) 
                ? dto.TipoAula 
                : (esVirtual == 1 ? "VIRTUAL" : "ESTANDAR");
            string caracteristicasJson = !string.IsNullOrWhiteSpace(dto.CaracteristicasJson) ? dto.CaracteristicasJson : "{}";

            // Fase 237: Recomputar AulaRef después de forzar modalidad para sincronía total con la tupla S
            string aulaRefCalculada = ConstruirAulaRef(dto.IdPeriodo, dto.IdSede, dto.IdPrograma, sem, grupo, modVinculacion);
            string codigo = !string.IsNullOrWhiteSpace(dto.Codigo) ? dto.Codigo : aulaRefCalculada;
            string nombre = !string.IsNullOrWhiteSpace(dto.Nombre) ? dto.Nombre : aulaRefCalculada;

            return new AgAula
            {
                Id = 0,
                IdSede = dto.IdSede,
                IdPeriodo = dto.IdPeriodo,
                IdPrograma = dto.IdPrograma,
                Codigo = codigo ?? string.Empty,
                Nombre = nombre ?? string.Empty,
                Capacidad = capacidad,
                TipoAula = tipoAula ?? "ESTANDAR",
                EsVirtual = esVirtual,
                Estado = AgEstadoAula.Disponible,
                CaracteristicasJson = caracteristicasJson ?? "{}",
                Semestre = sem,
                Grupo = grupo,
                ModalidadVinculacion = modVinculacion,
                FechaInicio = ventana != null ? ventana.FechaInicio.ToString("yyyy-MM-dd") : string.Empty,
                FechaFin = ventana != null ? ventana.FechaFin.ToString("yyyy-MM-dd") : string.Empty,
                SemestreInicio = ventana != null ? ventana.FechaInicio.ToString("yyyy-MM-dd") : string.Empty,
                SemestreFin = ventana != null ? ventana.FechaFin.ToString("yyyy-MM-dd") : string.Empty,
                VentanaOrigen = ventana != null ? ventana.Origen.ToString() : string.Empty,
                AulaRef = aulaRefCalculada,
                Subunidad = subunidad ?? string.Empty,
                Jornada = dto.Jornada ?? string.Empty,
                HorasDia = dto.HorasDia,
                DiasVentana = dto.DiasVentana ?? string.Empty
            };
        }
    }
}
