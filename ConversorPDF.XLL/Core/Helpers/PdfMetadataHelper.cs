using System;
using System.IO;
using Microsoft.Win32;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Interfaces;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Orquestador para la lectura de metadatos de archivos PDF seleccionados por el usuario.
    /// </summary>
    public static class PdfMetadataHelper
    {
        /// <summary>
        /// Inicia el flujo de usuario para seleccionar múltiples archivos PDF y registrar un resumen de sus metadatos.
        /// </summary>
        /// <remarks>
        /// Utiliza <see cref="Microsoft.Win32.OpenFileDialog"/> para la selección y <see cref="PdfMetadataService"/> 
        /// para extraer la información estructurada. Los resultados se registran en <see cref="LoggerHelper"/>.
        /// </remarks>
        public static void LeerMetadatosPDF()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "Seleccionar archivos PDF para leer metadatos",
                    Filter = "Archivos PDF (*.pdf)|*.pdf|Todos los archivos (*.*)|*.*",
                    Multiselect = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    IPdfMetadataService metadataSvc = new PdfMetadataService();
                    int count = 0;

                    foreach (string filePath in openFileDialog.FileNames)
                    {
                        if (!File.Exists(filePath)) continue;

                        try
                        {
                            string resumen = metadataSvc.ObtenerResumenMetadatos(filePath);
                            LoggerHelper.Info($"Resumen de metadatos para {Path.GetFileName(filePath)}:\n{resumen}");
                            count++;
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Error($"Error al leer metadatos de {filePath}", ex);
                        }
                    }

                    if (count > 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"Se han procesado {count} archivo(s). Consulte el log para ver los detalles.",
                            App.APP_NAME,
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error en LeerMetadatosPDF", ex);
                System.Windows.Forms.MessageBox.Show(
                    $"Error al procesar metadatos: {ex.Message}",
                    App.APP_NAME,
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
