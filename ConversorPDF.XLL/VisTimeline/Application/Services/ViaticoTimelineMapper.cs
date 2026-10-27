using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.VisTimeline.Domain.Models;

namespace ConversorPDF.XLL.VisTimeline.Application.Services
{
    public static class ViaticoTimelineMapper
    {
        private static readonly string[] PaletaSedes = {
            "#2E7D32", "#1565C0", "#E65100", "#6A1B9A",
            "#C62828", "#00838F", "#4E342E", "#37474F",
            "#AD1457", "#1B5E20", "#FF6F00", "#283593"
        };

        private static readonly string[] PaletaBordes = {
            "#1B5E20", "#0D47A1", "#BF360C", "#4A148C",
            "#B71C1C", "#006064", "#3E2723", "#263238",
            "#880E4F", "#094409", "#E65100", "#1A237E"
        };

        public static (List<TimelineItem> items, List<TimelineGroup> groups, string cssRules)
            Mapear(
                List<SesionListItem> sesiones,
                Dictionary<string, Tercero> terceros,
                bool incluirDatosBancarios = false)
        {
            var items = new List<TimelineItem>();
            var groups = new List<TimelineGroup>();

            // NIVEL 1: Docente  →  NIVEL 2: Sede  →  ITEMS: Bloques consecutivos
            foreach (var grupoDocente in sesiones.GroupBy(s => s.DocenteNombre).OrderBy(g => g.Key))
            {
                string docenteNombre = grupoDocente.Key;
                string docenteSlug   = Slugify(docenteNombre);
                string docenteId     = $"doc-{docenteSlug}";
                string docenteCorto  = TruncateStrict(docenteNombre, 15);   // ✅ 15 chars

                var sedesDelDocente = grupoDocente
                    .Select(s => s.SedeCiudad ?? "SIN SEDE")
                    .Distinct()
                    .OrderBy(s => s)
                    .ToList();

                var nestedIds = new List<string>();

                // ─── GRUPOS HIJO: Sedes ───
                foreach (var sede in sedesDelDocente)
                {
                    string sedeSlug = Slugify(sede);
                    string sedeId   = $"{docenteId}-{sedeSlug}";
                    nestedIds.Add(sedeId);

                    groups.Add(new TimelineGroup
                    {
                        id        = sedeId,
                        content   = $"📍 {sede.ToUpper()}",
                        className = $"grupo-sede-{sedeSlug}",
                        treeLevel = 2
                    });
                }

                // ─── GRUPO PADRE: Docente (nombre corto) ───
                groups.Add(new TimelineGroup
                {
                    id           = docenteId,
                    content      = $"👤 {docenteCorto.ToUpper()}",
                    className    = "grupo-docente",
                    nestedGroups = nestedIds,
                    showNested   = true,
                    treeLevel    = 1
                });

                // ─── ITEMS: Bloques consecutivos (no sesiones sueltas) ───
                foreach (var grupoSede in grupoDocente.GroupBy(s => s.SedeCiudad ?? "SIN SEDE").OrderBy(g => g.Key))
                {
                    string sede     = grupoSede.Key;
                    string sedeSlug = Slugify(sede);
                    string sedeId   = $"{docenteId}-{sedeSlug}";

                    var bloques = ViaticoBloque.AgruparConsecutivos(grupoSede.ToList());

                    foreach (var bloque in bloques)
                    {
                        if (!bloque.FechaMinima.HasValue) continue;  // seguridad

                        DateTime fechaIni = bloque.FechaMinima.Value;
                        DateTime fechaFin = bloque.FechaFinComision ?? bloque.FechaMaxima ?? fechaIni;

                        bool esRango = fechaFin.Date > fechaIni.Date;
                        string tipo  = esRango ? "range" : "box";
                        string end   = esRango
                            ? fechaFin.AddDays(1).ToString("yyyy-MM-dd")
                            : null;

                        // ✅ Contenido: Abreviaturas únicas del bloque + Total Horas
                        var abrevs = bloque.Sesiones
                            .Select(s => !string.IsNullOrWhiteSpace(s.AsignaturaAbreviatura)
                                ? s.AsignaturaAbreviatura.Trim().ToUpper()
                                : TruncateContent(s.AsignaturaNombre, 5).ToUpper())
                            .Distinct();

                        string abrevsText = string.Join(", ", abrevs);
                        abrevsText = TruncateContent(abrevsText, 12);

                        string content = $"{abrevsText} | {bloque.TotalHoras}h";

                        // ✅ Tooltip: Docente → Sede → detalle por sesión
                        string tooltip = BuildBloqueTooltip(
                            bloque, docenteNombre, sede, terceros, incluirDatosBancarios);

                        items.Add(new TimelineItem
                        {
                            id        = Guid.NewGuid().ToString("N"),  // evitar colisión de IDs
                            content   = content,
                            title     = tooltip,
                            start     = fechaIni.ToString("yyyy-MM-dd"),
                            end       = end,
                            group     = sedeId,
                            type      = tipo,
                            className = $"sede-{sedeSlug}"
                        });
                    }
                }
            }

            var todasLasSedes = sesiones
                .Select(s => s.SedeCiudad ?? "SIN SEDE")
                .Distinct()
                .OrderBy(s => s)
                .ToList();

            string cssRules = BuildCss(todasLasSedes);
            return (items, groups, cssRules);
        }

        // ═══════════════════════════════════════════════════════════════
        //  TOOLTIP DE BLOQUE: Docente / Sede / Sem | Asig | Xh | Día
        // ═══════════════════════════════════════════════════════════════
        private static string BuildBloqueTooltip(
            ViaticoBloque bloque,
            string docenteNombre,
            string sedeNombre,
            Dictionary<string, Tercero> terceros,
            bool incluirBancario)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<div style='font-family:Segoe UI,sans-serif;font-size:12.5px;line-height:1.55;max-width:360px;padding:2px'>");

            // Cabecera: Docente + Sede
            sb.AppendLine($"<div style='font-size:13px;font-weight:700;color:#fbbf24;margin-bottom:4px'>👤 {HtmlEncode(docenteNombre)}</div>");
            sb.AppendLine($"<div style='color:#94a3b8;margin-bottom:6px'>📍 {HtmlEncode(sedeNombre)}</div>");

            // Rango del bloque
            string rango = bloque.FechaMinima.HasValue
                ? bloque.FechaMinima.Value.ToString("dd/MM/yyyy")
                : "?";
            DateTime? fFinBloque = bloque.FechaFinComision ?? bloque.FechaMaxima;
            if (fFinBloque.HasValue && fFinBloque.Value.Date > bloque.FechaMinima.Value.Date)
                rango += $" → {fFinBloque.Value:dd/MM/yyyy}";

            sb.AppendLine($"<div style='font-size:11px;color:#64748b;margin-bottom:8px'>🗓️ {rango} &nbsp;|&nbsp; <b>{bloque.TotalHoras}h</b> &nbsp;|&nbsp; {bloque.DuracionTotal} días</div>");

            sb.AppendLine("<hr style='border:0;border-top:1px solid rgba(148,163,184,0.2);margin:6px 0'/>");

            // Detalle por sesión: Semestre | Asignatura | Xh | Día
            foreach (var sesion in bloque.Sesiones.OrderBy(s => s.FechaSesion))
            {
                string semestre = HtmlEncode(sesion.Semestre ?? "");
                string asignatura = HtmlEncode(sesion.AsignaturaNombre ?? "");
                string horas = $"{sesion.Horas}h";
                string dia = HtmlEncode(sesion.DiaSemana ?? "");

                sb.AppendLine($"<div style='display:flex;gap:8px;padding:2px 0;border-bottom:1px solid rgba(148,163,184,0.06)'>");
                sb.AppendLine($"  <span style='color:#60a5fa;min-width:28px'>{semestre}</span>");
                sb.AppendLine($"  <span style='flex:1;color:#e2e8f0'>{asignatura}</span>");
                sb.AppendLine($"  <span style='color:#fbbf24;font-weight:600;min-width:32px;text-align:right'>{horas}</span>");
                sb.AppendLine($"  <span style='color:#94a3b8;min-width:56px;text-align:right;font-size:11px'>{dia}</span>");
                sb.AppendLine($"</div>");
            }

            // Datos bancarios (si aplica)
            if (incluirBancario && bloque.Sesiones.FirstOrDefault() is SesionListItem primera)
            {
                if (!string.IsNullOrEmpty(primera.CedulaTercero) &&
                    terceros.TryGetValue(primera.CedulaTercero, out var tercero) && tercero != null)
                {
                    sb.AppendLine("<hr style='border:0;border-top:1px solid rgba(148,163,184,0.2);margin:8px 0'/>");
                    sb.AppendLine($"<div style='font-size:11px;color:#94a3b8'>🏦 {HtmlEncode(tercero.EntidadBancaria)} — {HtmlEncode(tercero.TipoCuentaBancaria)} #{MaskAccount(tercero.CuentaBancaria)}</div>");
                }
            }

            sb.AppendLine("</div>");
            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════
        private static string Slugify(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return "sin-dato";
            return texto.ToLower()
                        .Replace(" ", "-")
                        .Replace("á","a").Replace("é","e")
                        .Replace("í","i").Replace("ó","o")
                        .Replace("ú","u").Replace("ñ","n");
        }

        private static string TruncateContent(string texto, int maxLen)
        {
            if (string.IsNullOrEmpty(texto)) return "S/N";
            return texto.Length <= maxLen ? texto : texto.Substring(0, maxLen - 1) + "…";
        }

        private static string TruncateStrict(string texto, int maxLen)
        {
            if (string.IsNullOrEmpty(texto)) return "S/N";
            return texto.Length <= maxLen ? texto : texto.Substring(0, maxLen) + "…";
        }

        private static string HtmlEncode(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;");
        }

        private static string MaskAccount(string cuenta)
        {
            if (string.IsNullOrEmpty(cuenta) || cuenta.Length <= 4) return cuenta;
            return "****" + cuenta.Substring(cuenta.Length - 4);
        }

        private static string BuildCss(List<string> sedes)
        {
            var sb = new StringBuilder();

            // Grupo Docente (padre)
            sb.AppendLine(@"
.vis-label.grupo-docente {
    background: #0f172a;
    color: #fbbf24;
    font-weight: 700;
    font-size: 12px;
    letter-spacing: 0.3px;
    border-bottom: 1px solid #334155;
}
.vis-label.grupo-docente .vis-inner {
    padding: 7px 8px;
}");

            // Grupo Sede (hijo)
            for (int i = 0; i < sedes.Count; i++)
            {
                string slug  = Slugify(sedes[i]);
                string bg    = PaletaSedes[i % PaletaSedes.Length];
                string borde = PaletaBordes[i % PaletaBordes.Length];

                sb.AppendLine($@"
.vis-item.sede-{slug} {{
    background-color: {bg};
    border-color: {borde};
    color: #ffffff;
    font-weight: 600;
    font-size: 11px;
    letter-spacing: 0.2px;
    border-radius: 5px;
}}
.vis-item.sede-{slug}.vis-selected {{
    background-color: {borde};
    border-color: #fff;
    box-shadow: 0 0 0 3px rgba(255,255,255,0.15), 0 4px 12px rgba(0,0,0,0.4);
}}
.vis-label.grupo-sede-{slug} {{
    border-left: 4px solid {bg};
    padding-left: 12px;
    font-weight: 600;
    color: #e2e8f0;
    font-size: 12px;
    background: #1e293b;
}}");
            }
            return sb.ToString();
        }
    }
}
