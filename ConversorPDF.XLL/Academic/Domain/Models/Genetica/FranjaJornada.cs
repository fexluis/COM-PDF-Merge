namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa una tupla canónica de resolución de franja (hora inicio, hora fin, jornada y duración).
    /// </summary>
    public readonly struct FranjaJornada
    {
        public string HoraInicio { get; }
        public string HoraFin { get; }
        public string Franja { get; }
        public double DuracionHoras { get; }

        public FranjaJornada(string ini, string fin, string franja, double dur)
        {
            HoraInicio = ini;
            HoraFin = fin;
            Franja = franja;
            DuracionHoras = dur;
        }
    }
}
