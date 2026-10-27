using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Unit.Academic
{
    public class GridConfigurationTests
    {
        [Fact]
        public void SesionListItem_ColumnasMapeanDataPropertyName()
        {
            // Verifica que todas las columnas del grid tengan un mapeo válido en el modelo SesionListItem
            var propiedades = typeof(SesionListItem).GetProperties().Select(p => p.Name).ToList();
            
            var columnasEsperadas = new[] 
            { 
                "ProgramaDetalle", 
                "DocenteNombre", 
                "Semestre", 
                "AsignaturaNombre", 
                "SedeCiudad", 
                "DiaSemana", 
                "Horas", 
                "HorarioSesion", 
                "RangoFechas", 
                "Viatica", 
                "IdVinculacion" 
            };
            
            foreach (var col in columnasEsperadas)
            {
                propiedades.Should().Contain(col, $"la columna '{col}' no tiene una propiedad correspondiente en SesionListItem");
            }
        }
    }
}
