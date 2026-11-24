using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resultado global emitido tras la ejecución del motor del Algoritmo Genético.
    /// </summary>
    public class AgResultadoCorrida
    {
        public long IdCorrida { get; set; }
        public AgEstadoCorrida Estado { get; set; }
        /// <summary>
        /// Motivo del corte o finalización del bucle evolutivo del AG (Umbral, Estancamiento o MaxGeneraciones).
        /// </summary>
        public AgMotivoParada MotivoParada { get; set; }
        /// <summary>
        /// Conteo de generaciones consecutivas sin mejora del mejor fitness histórico alcanzado al momento del corte.
        /// </summary>
        public int Estancadas { get; set; }
        public int GeneracionesEjecutadas { get; set; }
        public long TiempoTotalMs { get; set; }
        public AgIndividuo MejorIndividuo { get; set; }
        public List<AgBitacora> HistorialConvergencia { get; set; }
        public List<AgPerfilGeneracion> PerfilesGeneracion { get; set; }
        public long TiempoFactibilidadMs { get; set; }
        public long TiempoIndicesMs { get; set; }
        public long TiempoPoblacionInicialMs { get; set; }
        public long TiempoPrimerCicloMs { get; set; }
        public long TiempoCampeonFinalMs { get; set; }
        public string Mensaje { get; set; }

        // Telemetría de reparación R5 (Fase 2329)
        public int ReparacionR5Intentos { get; set; }
        public int ReparacionR5HardPre { get; set; }
        public int ReparacionR5HardPost { get; set; }
        public double ReparacionR5DesperdicioPre { get; set; }
        public double ReparacionR5DesperdicioPost { get; set; }
        public int ReparacionR5RippleIntentos { get; set; }
        public int ReparacionR5RippleSinCola { get; set; }

        // Telemetría de reparación R6 (Fase 2370)
        public int ReparacionR6Intentos { get; set; }
        public int ReparacionR6HardPre { get; set; }
        public int ReparacionR6HardPost { get; set; }

        private bool? _esFactible;
        public bool EsFactible
        {
            get => _esFactible ?? (MejorIndividuo != null && MejorIndividuo.EsFactible);
            set => _esFactible = value;
        }

        public int Vacantes => MejorIndividuo?.Vacantes ?? 0;

        public AgResultadoCorrida()
        {
            Estado = AgEstadoCorrida.Iniciada;
            HistorialConvergencia = new List<AgBitacora>();
            PerfilesGeneracion = new List<AgPerfilGeneracion>();
            Mensaje = string.Empty;
        }

        public class AgPerfilGeneracion
        {
            public int Generacion { get; set; }
            public long DuracionTotalMs { get; set; }
            public long TiempoSeleccionMs { get; set; }
            public long TiempoCruceNetoMs { get; set; }
            public long TiempoMutacionNetoMs { get; set; }
            public long TiempoSelladoTotalMs { get; set; }
            public long TiempoEvaluacionMs { get; set; }
            public int EvaluacionesAptitud { get; set; }
            public int HijosCreados { get; set; }
            public int HijosIncorporados { get; set; }
            public int ElitesClonados { get; set; }
            public int EvaluacionesExtraPorCupoImpar { get; set; }
            public int Vacantes { get; set; }
            public int FueraDeVentana { get; set; }
        }
    }
}
