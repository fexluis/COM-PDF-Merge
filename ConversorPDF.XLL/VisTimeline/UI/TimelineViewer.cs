using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ConversorPDF.XLL.VisTimeline.Application.Services;
using ConversorPDF.XLL.VisTimeline.Domain.Models;
using ConversorPDF.XLL.VisTimeline.Infrastructure.Helpers;
using ConversorPDF.XLL.VisTimeline.Infrastructure.Services;

namespace ConversorPDF.XLL.VisTimeline.UI
{
    public static class TimelineViewer
    {
        public static void Mostrar(List<TimelineItem> items, List<TimelineGroup> groups, string cssRules)
        {
            try
            {
                // Ruta base en Mis Documentos
                string baseDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "VisTimeline");
                
                string assetsDir = Path.Combine(baseDir, "assets");
                Directory.CreateDirectory(assetsDir);

                // Extraer recursos embebidos (caché en disco)
                string ns = "ConversorPDF.XLL.VisTimeline.Resources";
                ResourceExtractor.ExtractToDocuments($"{ns}.vis-data.min.js", assetsDir, "vis-data.min.js");
                ResourceExtractor.ExtractToDocuments($"{ns}.vis-timeline-graph2d.min.js", assetsDir, "vis-timeline-graph2d.min.js");
                ResourceExtractor.ExtractToDocuments($"{ns}.vis-timeline-graph2d.min.css", assetsDir, "vis-timeline-graph2d.min.css");
                ResourceExtractor.ExtractToDocuments($"{ns}.lucide.min.js", assetsDir, "lucide.min.js");

                // Generar HTML con rutas locales
                string html = TimelineHtmlGenerator.GenerarHtml(items, groups, cssRules, assetsDir);

                // Guardar HTML en la raz de la carpeta de trabajo
                string htmlPath = Path.Combine(baseDir, "viewer.html");
                File.WriteAllText(htmlPath, html, new UTF8Encoding(true));

                // Abrir en navegador
                Process.Start(new ProcessStartInfo(htmlPath) { UseShellExecute = true });

                MessageBox.Show(
                    "El Timeline se ha generado y se abrir en su navegador predeterminado.\n\n" +
                    $"Carpeta de trabajo: {baseDir}",
                    "VisTimeline",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al visualizar el timeline: {ex.Message}",
                    "Error VisTimeline",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
