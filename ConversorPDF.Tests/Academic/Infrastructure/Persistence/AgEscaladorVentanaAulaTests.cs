using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgEscaladorVentanaAulaTests
    {
        [Fact]
        public void DecisionAplica_CuandoCapacidadCrece_DevuelveTrue()
        {
            var decision = new AgEscaladorVentanaAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "VIE-SAB",
                Hacia = "JUE-SAB",
                DemandaSesiones = 10,
                CapacidadAntes = 6,
                CapacidadDespues = 12,
                Motivo = "demanda > capacidad"
            };

            decision.Aplica.Should().BeTrue();
        }

        [Fact]
        public void DecisionAplica_CuandoCapacidadNoCrece_DevuelveFalse()
        {
            var decision = new AgEscaladorVentanaAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "VIE-SAB",
                Hacia = "VIE-SAB",
                DemandaSesiones = 10,
                CapacidadAntes = 6,
                CapacidadDespues = 6,
                Motivo = "sin cambio"
            };

            decision.Aplica.Should().BeFalse();
        }

        [Fact]
        public void DecisionAplica_CuandoHaciaEsVacio_DevuelveFalse()
        {
            var decision = new AgEscaladorVentanaAula.Decision
            {
                AulaRef = "A101",
                Subunidad = "",
                Desde = "4,5,6,7",
                Hacia = "",
                DemandaSesiones = 20,
                CapacidadAntes = 16,
                CapacidadDespues = 16,
                Motivo = "tope de escalera"
            };

            decision.Aplica.Should().BeFalse();
        }
    }
}
