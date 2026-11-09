using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
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

        /// <summary>
        /// Obtiene las corridas del staging AG con su volumen de vinculaciones y sesiones (Fase 1489).
        /// Incluye corridas sin staging materializado y staging huérfano sin cabecera en ag_corrida.
        /// </summary>
        /// <param name="idPeriodo">Filtro opcional por período.</param>
        /// <param name="idPrograma">Filtro opcional por programa.</param>
        /// <param name="subunidad">Filtro opcional por territorial (comparación tolerante).</param>
        /// <param name="ct">Token de cancelación.</param>
        Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
            int? idPeriodo,
            int? idPrograma,
            string subunidad,
            CancellationToken ct);

        /// <summary>
        /// Construye el resumen de diagnóstico del staging de una corrida (Fases 1489/1499).
        /// Separa lo que contiene la corrida (sin filtros) de lo que devuelve la consulta con filtros.
        /// Recibe parámetros primitivos para no invertir la dependencia Domain -> Application.
        /// </summary>
        Task<Result<AgStagingResumen>> GetResumenStagingAsync(
            long idCorrida,
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
            CancellationToken ct);

        /// <summary>
        /// Elimina el staging completo de una corrida (vinculaciones, sesiones, genealogía,
        /// individuos, bitácoras, violaciones, promociones, salvedades y cabecera en ag_corrida).
        /// Operación transaccional e irreversible (Fases 1516/1517).
        /// </summary>
        /// <returns>Conteos de sesiones y vinculaciones eliminadas.</returns>
        Task<Result<(int Sesiones, int Vinculaciones)>> EliminarCorridaStagingAsync(
            long idCorrida,
            CancellationToken ct);
    }
}
