using System;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Fábrica para la creación de encabezados de vinculación en staging (ag_vinculaciones) a partir de aulas y genes.
    /// </summary>
    public static class AgEncabezadoVinculacionFactory
    {
        public static AgVinculacion Crear(AgGen gen, AgAula aula, long idCorrida, long idIndividuo, string subunidad)
        {
            if (gen == null) throw new ArgumentNullException(nameof(gen));

            return new AgVinculacion
            {
                IdCorrida = idCorrida,
                IdIndividuo = idIndividuo,
                IdTercero = gen.IdTercero,
                IdPrograma = gen.IdPrograma,
                Idsede = gen.IdSede,
                IdAsignatura = gen.IdAsignatura,
                IdPeriodo = gen.IdPeriodo > 0 ? gen.IdPeriodo : (aula != null ? aula.IdPeriodo : 0),
                Semestre = !string.IsNullOrWhiteSpace(gen.Semestre) ? gen.Semestre : (aula?.Semestre ?? string.Empty),
                Grupo = !string.IsNullOrWhiteSpace(gen.Grupo) ? gen.Grupo : (aula?.Grupo ?? "A"),
                TotalHoras = gen.TotalHorasAsignatura,
                FechaInicio = gen.FechaInicioEstimada.HasValue ? gen.FechaInicioEstimada.Value.ToString("yyyy-MM-dd") : (aula?.FechaInicio ?? string.Empty),
                ModalidadVinculacion = aula != null && aula.EsVirtual == 1
                    ? ModalidadTutoriaCatalog.Plataforma
                    : ModalidadTutoriaCatalog.NormalizarTutoria(aula?.ModalidadVinculacion),
                Estado = AgEstadoVinculacion.Borrador,
                IdTerminal = 1,
                Subunidad = subunidad ?? string.Empty
            };
        }
    }
}
