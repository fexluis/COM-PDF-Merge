using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de extracción de datos de documentos para firma.
    /// </summary>
    public interface IFirmaExtractService
    {
        /// <summary>
        /// Extrae la información de una obligación presupuestal desde una hoja de Excel.
        /// </summary>
        /// <param name="ws">Hoja de Excel (dynamic).</param>
        /// <returns>Modelo con los datos extraídos.</returns>
        ExtraerDatos ObligacionData(dynamic ws);

        /// <summary>
        /// Extrae datos de Orden de Pago (Boletines).
        /// </summary>
        ExtraerDatos OrdeDePagoData(dynamic ws);

        /// <summary>
        /// Extrae datos de Certificado de Disponibilidad Presupuestal (CDP).
        /// </summary>
        ExtraerDatos CdpData(dynamic ws);

        /// <summary>
        /// Extrae datos de Registro Presupuestal (RP).
        /// </summary>
        ExtraerDatos RpData(dynamic ws);

        /// <summary>
        /// Extrae datos de Verificado (Comisión Servicio).
        /// </summary>
        ExtraerDatos VerificadosData(dynamic ws);
    }
}
