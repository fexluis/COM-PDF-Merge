using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// Resultado inmutable de la resolución de selección en el selector de corridas del Módulo R (Fase 1606).
    /// </summary>
    public sealed class AgSeleccionModuloR
    {
        public CorridaStagingItem ItemSeleccionado { get; }
        public int IndiceSeleccionado { get; }
        public bool RequiereCargaInicial { get; }
        public bool CorridaActivaNoEncontrada { get; }

        public AgSeleccionModuloR(
            CorridaStagingItem itemSeleccionado,
            int indiceSeleccionado,
            bool requiereCargaInicial,
            bool corridaActivaNoEncontrada)
        {
            ItemSeleccionado = itemSeleccionado;
            IndiceSeleccionado = indiceSeleccionado;
            RequiereCargaInicial = requiereCargaInicial;
            CorridaActivaNoEncontrada = corridaActivaNoEncontrada;
        }
    }
}
