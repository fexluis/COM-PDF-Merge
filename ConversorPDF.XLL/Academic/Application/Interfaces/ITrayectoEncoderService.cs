namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface ITrayectoEncoderService
    {
        string GetOrCreateCodigoTrayecto(string territorial, string trayecto, decimal? valorTransporte);
    }
}
