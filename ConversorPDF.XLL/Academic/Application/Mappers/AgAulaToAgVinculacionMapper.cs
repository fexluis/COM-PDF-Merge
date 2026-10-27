using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador estático de atributos de AgAula sobre AgVinculacion.
    /// </summary>
    public static class AgAulaToAgVinculacionMapper
    {
        public static void AplicarAtributosAula(AgVinculacion vinculacion, AgAula aula)
        {
            if (vinculacion == null || aula == null) return;

            if (vinculacion.IdPeriodo <= 0) vinculacion.IdPeriodo = aula.IdPeriodo;
            if (vinculacion.IdPrograma <= 0) vinculacion.IdPrograma = aula.IdPrograma;
            if (string.IsNullOrWhiteSpace(vinculacion.Semestre)) vinculacion.Semestre = aula.Semestre;
            if (string.IsNullOrWhiteSpace(vinculacion.Grupo)) vinculacion.Grupo = aula.Grupo;
            if (string.IsNullOrWhiteSpace(vinculacion.FechaInicio)) vinculacion.FechaInicio = aula.FechaInicio;
            if (string.IsNullOrWhiteSpace(vinculacion.FechaFin)) vinculacion.FechaFin = aula.FechaFin;
            vinculacion.ModalidadVinculacion = ModalidadTutoriaCatalog.NormalizarTutoria(aula.ModalidadVinculacion);
        }
    }
}
