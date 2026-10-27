using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador estático de AgVinculacion (staging) a Vinculacion (producción).
    /// </summary>
    public static class AgVinculacionToVinculacionMapper
    {
        public static Vinculacion MapToProduccion(AgVinculacion agVinc)
        {
            if (agVinc == null) return null;

            DateTime? fInicio = null;
            if (DateTime.TryParse(agVinc.FechaInicio, out DateTime fi)) fInicio = fi;

            DateTime? fFin = null;
            if (DateTime.TryParse(agVinc.FechaFin, out DateTime ff)) fFin = ff;

            return new Vinculacion
            {
                IdTercero = agVinc.IdTercero,
                IdPrograma = agVinc.IdPrograma,
                IdSede = agVinc.IdSede,
                IdAsignatura = agVinc.IdAsignatura,
                IdPeriodo = agVinc.IdPeriodo,
                Semestre = agVinc.Semestre,
                Grupo = agVinc.Grupo ?? string.Empty,
                TotalHoras = (int)Math.Round(agVinc.TotalHoras),
                DiaHorasSemanales = agVinc.DiaHorasSemanales ?? string.Empty,
                FechaInicio = fInicio,
                FechaFin = fFin,
                Modalidad = ModalidadTutoriaCatalog.NormalizarTutoria(agVinc.ModalidadVinculacion),
                IsTerminalAereo = agVinc.IsTerminalAereo,
                IdTerminales = agVinc.IsTerminalAereo == 1 ? agVinc.IdTerminal : 0,
                GastoTerminalAerea = agVinc.IsTerminalAereo == 1 ? agVinc.GastoTerminalAerea : 0,
                EsVacante = agVinc.EsVacante,
                MotivoVacante = agVinc.MotivoVacante ?? string.Empty
            };
        }
    }
}
