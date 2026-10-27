using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador bidireccional entre AgContextoSeleccionDto (tupla S) y la entidad persistente AgArbolNodo (Fase 239).
    /// Garantiza la sincronización de AulaRef, ModalidadEfectiva y atributos de aula.
    /// </summary>
    public static class AgArbolMapper
    {
        public static AgArbolNodo Mapear(AgContextoSeleccionDto dto, long idArbol, string subunidad)
        {
            if (dto == null) throw new ArgumentNullException(nameof(dto));

            string modEfectiva = dto.ModalidadEfectiva;
            int esVirtual = (dto.EsVirtual > 0 || ModalidadTutoriaCatalog.EsPlataforma(modEfectiva)) ? 1 : 0;
            string aulaRef = AgAulaFactory.ConstruirAulaRef(dto.IdPeriodo, dto.IdSede, dto.IdPrograma, dto.Semestre, dto.Grupo, modEfectiva);

            return new AgArbolNodo
            {
                Id = 0,
                IdArbol = idArbol,
                IdPeriodo = dto.IdPeriodo,
                IdSede = dto.IdSede,
                IdPrograma = dto.IdPrograma,
                Semestre = dto.Semestre ?? string.Empty,
                Grupo = string.IsNullOrWhiteSpace(dto.Grupo) ? "A" : dto.Grupo.Trim().ToUpperInvariant(),
                Modalidad = modEfectiva,
                AulaRef = aulaRef,
                Capacidad = dto.Capacidad > 0 ? dto.Capacidad : (esVirtual == 1 ? 999 : 30),
                TipoAula = !string.IsNullOrWhiteSpace(dto.TipoAula) ? dto.TipoAula : (esVirtual == 1 ? "VIRTUAL" : "ESTANDAR"),
                CaracteristicasJson = !string.IsNullOrWhiteSpace(dto.CaracteristicasJson) ? dto.CaracteristicasJson : "{}",
                Codigo = dto.Codigo ?? string.Empty,
                Nombre = dto.Nombre ?? string.Empty,
                EsVirtual = esVirtual,
                Habilitado = 1,
                Subunidad = subunidad ?? string.Empty,
                Jornada = dto.Jornada ?? string.Empty,
                HorasDia = dto.HorasDia,
                DiasVentana = dto.DiasVentana ?? string.Empty
            };
        }

        public static AgContextoSeleccionDto Mapear(AgArbolNodo nodo)
        {
            if (nodo == null) throw new ArgumentNullException(nameof(nodo));

            string mod = ModalidadTutoriaCatalog.NormalizarTutoria(nodo.Modalidad);
            int esVirtual = nodo.EsVirtual > 0 || ModalidadTutoriaCatalog.EsPlataforma(mod) ? 1 : 0;
            string modEfectiva = esVirtual == 1 ? ModalidadTutoriaCatalog.Plataforma : mod;

            return new AgContextoSeleccionDto
            {
                IdPeriodo = nodo.IdPeriodo,
                IdSede = nodo.IdSede,
                IdPrograma = nodo.IdPrograma,
                Semestre = nodo.Semestre ?? string.Empty,
                Grupo = string.IsNullOrWhiteSpace(nodo.Grupo) ? "A" : nodo.Grupo.Trim().ToUpperInvariant(),
                Modalidad = modEfectiva,
                Capacidad = nodo.Capacidad > 0 ? nodo.Capacidad : (esVirtual == 1 ? 999 : 30),
                TipoAula = !string.IsNullOrWhiteSpace(nodo.TipoAula) ? nodo.TipoAula : (esVirtual == 1 ? "VIRTUAL" : "ESTANDAR"),
                CaracteristicasJson = !string.IsNullOrWhiteSpace(nodo.CaracteristicasJson) ? nodo.CaracteristicasJson : "{}",
                Codigo = nodo.Codigo ?? string.Empty,
                Nombre = nodo.Nombre ?? string.Empty,
                EsVirtual = esVirtual,
                Jornada = nodo.Jornada ?? string.Empty,
                HorasDia = nodo.HorasDia,
                DiasVentana = nodo.DiasVentana ?? string.Empty
            };
        }
    }
}
