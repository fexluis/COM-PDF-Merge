using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Interfaz para consultas especializadas de vinculaciones.
    /// </summary>
    public interface IVinculacionQueryRepository
    {
        /// <summary>
        /// Obtiene una lista detallada de sesiones filtrada por sede, tercero, período, viática y rango de fechas.
        /// </summary>
        /// <param name="idSede">Identificador opcional de la sede.</param>
        /// <param name="idTercero">Identificador opcional del tercero (docente).</param>
        /// <param name="idPeriodo">Identificador opcional del período académico.</param>
        /// <param name="idPrograma">Identificador opcional del programa académico.</param>
        /// <param name="modalidadPrograma">Modalidad del programa (PREGRADO o POSGRADO).</param>
        /// <param name="subunidad">Nombre de la territorial (subunidad) para filtrar sedes y terceros.</param>
        /// <param name="semestre">Número de semestre a filtrar.</param>
        /// <param name="soloViaticas">Si es true, solo retorna sesiones marcadas con viática.</param>
        /// <param name="fechaDesde">Fecha inicial opcional del rango de búsqueda (inclusive).</param>
        /// <param name="fechaHasta">Fecha final opcional del rango de búsqueda (inclusive).</param>
        /// <param name="limit">Número máximo de registros a retornar.</param>
        /// <param name="offset">Número de registros a omitir (paginación).</param>
        /// <param name="ct">Token de cancelación.</param>
        Task<Result<IReadOnlyCollection<SesionListItem>>> GetSesionesAsync(
            int? idSede, 
            int? idTercero, 
            int? idPeriodo,
            int? idPrograma,
            string modalidadPrograma,
            string subunidad,
            string semestre,
            bool soloViaticas,
            DateTime? fechaDesde,
            DateTime? fechaHasta,
            int limit,
            int offset,
            CancellationToken ct);

        /// <summary>
        /// Obtiene una lista detallada de sesiones filtrada adicionalmente por corrida (Modo AG / Staging).
        /// </summary>
        Task<Result<IReadOnlyCollection<SesionListItem>>> GetSesionesAsync(
            int? idSede, 
            int? idTercero, 
            int? idPeriodo,
            int? idPrograma,
            string modalidadPrograma,
            string subunidad,
            string semestre,
            bool soloViaticas,
            DateTime? fechaDesde,
            DateTime? fechaHasta,
            int limit,
            int offset,
            CancellationToken ct,
            long? idCorrida);
    }
}
