using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Implementación desacoplada del cálculo de combinaciones seleccionadas para el motor AG.
    /// Realiza la expansión combinatoria de los 6 niveles jerárquicos produciendo tuplas canónicas.
    /// </summary>
    public class AgTreeViewCheckService : ITreeViewCheckService
    {
        public IReadOnlyCollection<AgContextoSeleccionDto> ExpandirCombinaciones(
            IEnumerable<int> periodosIds,
            IEnumerable<int> sedesIds,
            IEnumerable<int> programasIds,
            IEnumerable<string> semestres,
            IEnumerable<string> grupos,
            IEnumerable<string> modalidades)
        {
            var pList = (periodosIds ?? Enumerable.Empty<int>()).Distinct().ToList();
            var sList = (sedesIds ?? Enumerable.Empty<int>()).Distinct().ToList();
            var progList = (programasIds ?? Enumerable.Empty<int>()).Distinct().ToList();
            var semList = (semestres ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            var grpList = (grupos ?? Enumerable.Empty<string>()).Where(g => !string.IsNullOrWhiteSpace(g)).Distinct().ToList();
            var modList = (modalidades ?? Enumerable.Empty<string>()).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct().ToList();

            if (!pList.Any() || !sList.Any() || !progList.Any() || !semList.Any() || !grpList.Any() || !modList.Any())
            {
                return new List<AgContextoSeleccionDto>().AsReadOnly();
            }

            var combinaciones = new List<AgContextoSeleccionDto>();

            foreach (var p in pList)
            {
                foreach (var s in sList)
                {
                    foreach (var prog in progList)
                    {
                        foreach (var sem in semList)
                        {
                            foreach (var grp in grpList)
                            {
                                foreach (var mod in modList)
                                {
                                    combinaciones.Add(new AgContextoSeleccionDto(p, s, prog, sem, grp, mod));
                                }
                            }
                        }
                    }
                }
            }

            return combinaciones.AsReadOnly();
        }
    }
}
