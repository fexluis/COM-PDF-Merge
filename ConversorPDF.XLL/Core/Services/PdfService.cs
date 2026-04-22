using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio encargado de la manipulación de documentos PDF utilizando la librería PDFSharp.
    /// </summary>
    public class PdfService : IPdfService
    {
        /// <summary>
        /// Combina una serie de archivos PDF existentes en un nuevo documento único.
        /// Este proceso importa las páginas de cada archivo de origen de forma secuencial.
        /// </summary>
        /// <param name="archivos">Colección de rutas completas a los archivos PDF de entrada.</param>
        /// <param name="rutaSalida">Ruta absoluta donde se guardará el PDF resultante.</param>
        /// <exception cref="InvalidOperationException">Se lanza si no se agregan páginas al documento de salida (ej. si todos los archivos de origen son inválidos).</exception>
        public void CombinarPDFs(IEnumerable<string> archivos, string rutaSalida)
        {
            var dir = Path.GetDirectoryName(rutaSalida);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var outputDoc = new PdfDocument())
            {
                outputDoc.Info.Title = "PDF combinado";
                foreach (var file in archivos)
                {
                    if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
                        continue;

                    using (var inputDoc = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                    {
                        int count = inputDoc.PageCount;
                        for (int i = 0; i < count; i++)
                            outputDoc.AddPage(inputDoc.Pages[i]);
                    }
                }

                if (outputDoc.PageCount == 0)
                    throw new InvalidOperationException("No se agregaron páginas al PDF de salida.");

                outputDoc.Save(rutaSalida);
            }
        }
    }
}