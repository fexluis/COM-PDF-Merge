using System;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio para la extracción de datos de documentos para firma.
    /// </summary>
    public class FirmaExtractService : IFirmaExtractService
    {
        /// <inheritdoc/>
        public ExtraerDatos ObligacionData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xDep = ws.Range["B32"].Value?.ToString();

            if (ws.Range["F16"].Value?.ToString() == "Estado:")
            {
                extract.xNumero = ws.Range["C15"].Value?.ToString() ?? "";
                extract.xFechaStr = ws.Range["I15"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["I22"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["C22"].Value?.ToString() ?? "";
                extract.xRP = ws.Range["AE16"].Value?.ToString() ?? "";
                extract.xObjeto = ws.Range["C38"].Value?.ToString() ?? "";
                extract.xCategoria = FirmaPDFsHelper.GetCategoria(extract.xObjeto);
            }
            else
            {
                extract.xNumero = ws.Range["D15"].Value?.ToString() ?? "";
                extract.xFechaStr = ws.Range["K15"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["K22"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["D22"].Value?.ToString() ?? "";
                extract.xRP = ws.Range["AN16"].Value?.ToString() ?? "";

                dynamic rngObjeto = ws.Range["B1:B300"].Find("Objeto:");
                if (rngObjeto != null)
                {
                    extract.xObjeto = ws.Range["D" + rngObjeto.Row].Value?.ToString() ?? "";
                    extract.xCategoria = FirmaPDFsHelper.GetCategoria(extract.xObjeto);
                }
                else
                {
                    extract.xObjeto = "DESCONOCIDO";
                    extract.xCategoria = "XXX";
                }
            }

            extract.xUnifiedFileName = $"OB-{extract.xNumero}-{extract.xCategoria}-{FirmaPDFsHelper.SanitizarNombre(extract.xTercero)}";
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos OrdeDePagoData(dynamic ws)
        {
            var extract = new ExtraerDatos();

            if (ws.Range["X16"].Value != null)
            {
                extract.xFechaStr = ws.Range["J15"].Value?.ToString() ?? "";
                extract.xOB = ws.Range["X16"].Value?.ToString() ?? "";
                extract.xOP = ws.Range["B15"].Value?.ToString() ?? "";
                extract.xEstado = ws.Range["J16"].Value?.ToString() ?? "";
                extract.xTercero = ws.Range["J30"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["B30"].Value?.ToString() ?? "";
                
                extract.xFechaStr = extract.xFechaStr.Replace("-", "");
                extract.xTercero = (extract.xEstado == "Anulada") ? "Anulada" : extract.xTercero;
                extract.xUnifiedFileName = extract.xFechaStr + "_" + extract.xOP + "_" + extract.xOB + "_";
            }
            else
            {
                extract.xFechaStr = ws.Range["I15"].Value?.ToString() ?? "";
                extract.xOB = ws.Range["U16"].Value?.ToString() ?? "";
                extract.xOP = ws.Range["B15"].Value?.ToString() ?? "";
                extract.xEstado = ws.Range["I16"].Value?.ToString() ?? "";
                extract.xIdentificacion = ws.Range["B30"].Value?.ToString() ?? "";
                
                extract.xFechaStr = extract.xFechaStr.Replace("-", "");

                extract.xTercero = string.IsNullOrEmpty(ws.Range["I30"].Value?.ToString()) ? ws.Range["I31"].Value?.ToString() : ws.Range["I30"].Value?.ToString();
                extract.xTercero = (extract.xEstado == "Anulada") ? "Anulada" : extract.xTercero;
                extract.xUnifiedFileName = extract.xFechaStr + "_" + extract.xOP + "_" + extract.xOB + "_";
            }

            extract.xUnifiedFileName = FirmaPDFsHelper.SanitizarNombre(extract.xUnifiedFileName);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos CdpData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["E18"].Value?.ToString() ?? "";
            extract.xCDP = ws.Range["B18"].Value?.ToString() ?? "";
            extract.xNumero = extract.xCDP;
            
            string prefijo = extract.xFechaStr.Replace("-", "") + "_CDP_" + extract.xCDP + "_";
            extract.xUnifiedFileName = FirmaPDFsHelper.SanitizarNombre(prefijo);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos RpData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["I16"].Value?.ToString() ?? "";
            extract.xRP = ws.Range["F16"].Value?.ToString() ?? "";
            extract.xTercero = ws.Range["I20"].Value?.ToString() ?? "";
            extract.xNumero = extract.xRP;

            string prefijo = extract.xFechaStr.Replace("-", "") + "_RP_" + extract.xRP + "_";
            extract.xUnifiedFileName = FirmaPDFsHelper.SanitizarNombre(prefijo + extract.xTercero);
            return extract;
        }

        /// <inheritdoc/>
        public ExtraerDatos VerificadosData(dynamic ws)
        {
            var extract = new ExtraerDatos();
            extract.xFechaStr = ws.Range["N15"].Value?.ToString() ?? "";
            extract.xNumero = ws.Range["C15"].Value?.ToString() ?? "";
            extract.xTercero = ws.Range["A23"].Value?.ToString() ?? "";

            string prefijo = extract.xFechaStr.Replace("-", "") + "_V_" + extract.xNumero + "_";
            extract.xUnifiedFileName = FirmaPDFsHelper.SanitizarNombre(prefijo + extract.xTercero);
            return extract;
        }

    }
}
