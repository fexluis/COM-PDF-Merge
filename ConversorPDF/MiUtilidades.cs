using System;
using System.IO;
using System.Runtime.InteropServices;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ConversorPDF
{
    [ComVisible(true)]
    [Guid("4F97F332-3B6C-4B6E-8B7C-5AF9F4EEA4C9")]
    [ProgId("ConversorPDF.MiUtilidades")]
    [ClassInterface(ClassInterfaceType.None)]
    public class MiUtilidades : IMiUtilidades
    {
        public string Saludar(string nombre)
        {
            return "¡Hola, " + nombre + "! Desde ConversorPDF .NET COM.";
        }

        public double Sumar(double a, double b)
        {
            return a + b;
        }

        public void CombinarArchivos(object entradas, string salida)
        {
            try
            {
                string[] listaArchivos = null;

                if (entradas == null)
                    throw new COMException("El argumento 'entradas' es nulo.");

                if (entradas is string[] arrStr)
                {
                    listaArchivos = arrStr;
                }
                else if (entradas is object[] arrObj)
                {
                    listaArchivos = new string[arrObj.Length];
                    for (int i = 0; i < arrObj.Length; i++)
                        listaArchivos[i] = arrObj[i]?.ToString();
                }
                else
                {
                    // Intento de conversión genérica si viene como Variant encapsulado (SAFEARRAY)
                    try 
                    {
                        listaArchivos = (string[])entradas;
                    }
                    catch
                    {
                        throw new COMException("El argumento 'entradas' no se pudo convertir a Array de Strings. Tipo recibido: " + entradas.GetType().FullName);
                    }
                }

                if (listaArchivos == null || listaArchivos.Length == 0)
                    throw new COMException("No hay archivos de entrada.");

                var dir = Path.GetDirectoryName(salida);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (var outputDoc = new PdfDocument())
                {
                    outputDoc.Info.Title = "PDF combinado";
                    foreach (var file in listaArchivos)
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
                        throw new COMException("No se agregaron páginas al PDF de salida.");

                    outputDoc.Save(salida);
                }
            }
            catch (COMException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new COMException("Error al combinar PDFs: " + ex.Message);
            }
        }

        public string GetOneDriveLocalPath(string path)
        {
            return OneDriveHelper.GetOneDriveLocalPath(path);
        }

        public string GetCurrentOfficeUser()
        {
            return OneDriveHelper.GetCurrentOfficeUser();
        }

        public string GetFirmaBase64(string nombreFirma)
        {
            if (string.Equals(nombreFirma, "MONTANO", StringComparison.OrdinalIgnoreCase))
                return FirmasHelper.FirmaMontano;
            else if (string.Equals(nombreFirma, "VILLEGAS", StringComparison.OrdinalIgnoreCase))
                return FirmasHelper.FirmaVillegas;
            else
                return FirmasHelper.FirmaLuis;
        }

        public string Base64ToTempFile(string base64String, string nombreBase = "Imagen")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(base64String) || base64String.Length < 100)
                    return "";

                byte[] bytes = Convert.FromBase64String(base64String);

                string tempFolder = Path.GetTempPath();
                string cleanName = string.Join("_", nombreBase.Split(Path.GetInvalidFileNameChars()));
                string fileName = $"{cleanName}_{DateTime.Now:yyyyMMdd_HHmmss}_{new Random().Next(10000)}.png";
                string fullPath = Path.Combine(tempFolder, fileName);

                File.WriteAllBytes(fullPath, bytes);

                if (File.Exists(fullPath))
                    return fullPath;

                return "";
            }
            catch (Exception ex)
            {
                // In COM, returning empty or throwing depends on contract, returning empty for safety as VBA does
                return "";
            }
        }

        public string FirmaToTempFile(string nombreFirma)
        {
            string base64Str = GetFirmaBase64(nombreFirma);
            if (string.IsNullOrEmpty(base64Str)) return string.Empty;

            string nombreBase = "Firma_" + nombreFirma;
            return Base64ToTempFile(base64Str, nombreBase);
        }

        public int ConvertirObligacionContador(object activeWorkbook, string userFirma)
        {
            return ContadorUtilidades.ConvertirObligacionContador(activeWorkbook, userFirma, this);
        }

        public int UnificarObligacionContador(object activeWorkbook, string docType)
        {
            return ContadorUtilidades.UnificarObligacionContador(activeWorkbook, docType, this);
        }

        public string GenerarFirmaBase64Txt()
        {
            return GeneradorFirmasHelper.GenerarStringBase64DesdeImagen();
        }
    }
}
