using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using ConversorPDF.XLL.VisTimeline.Domain.Models;

namespace ConversorPDF.XLL.VisTimeline.Infrastructure.Services
{
    public static class TimelineHtmlGenerator
    {
        public static string GenerarHtml(
            List<TimelineItem> items,
            List<TimelineGroup> groups,
            string cssRules,
            string assetsDir)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            string itemsJson  = serializer.Serialize(items);
            string groupsJson = serializer.Serialize(groups);

            // Ventana inicial: desde 2 días antes del primer bloque hasta 60 días después
            DateTime? fechaMin = items
                .Select(i => { DateTime.TryParse(i.start, out var d); return (DateTime?)d; })
                .Where(d => d.HasValue)
                .Min();

            string startStr = fechaMin.HasValue 
                ? fechaMin.Value.AddDays(-2).ToString("yyyy-MM-dd") 
                : "";
            string endStr = fechaMin.HasValue 
                ? fechaMin.Value.AddDays(60).ToString("yyyy-MM-dd") 
                : "";

            // Rutas locales seguras para HTML
            string visDataJs      = new Uri(Path.Combine(assetsDir, "vis-data.min.js")).AbsoluteUri;
            string visTimelineJs  = new Uri(Path.Combine(assetsDir, "vis-timeline-graph2d.min.js")).AbsoluteUri;
            string visTimelineCss = new Uri(Path.Combine(assetsDir, "vis-timeline-graph2d.min.css")).AbsoluteUri;
            string lucideJs       = new Uri(Path.Combine(assetsDir, "lucide.min.js")).AbsoluteUri;

            return $@"<!DOCTYPE html>
<html lang=""es"">
<head>
  <meta charset=""UTF-8""/>
  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0""/>
  <title>VisTimeline — Viáticos (Supreme Edition)</title>
  <link rel=""stylesheet"" href=""{visTimelineCss}""/>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: #0b1120;
      color: #e2e8f0;
      overflow: hidden;
      height: 100vh; width: 100vw;
    }}
    
    /* ── TOOLBAR ── */
    #toolbar {{
      position: fixed; top: 10px; right: 14px; z-index: 999;
      display: flex; gap: 6px; padding: 6px 10px;
      background: rgba(15,23,42,0.85); backdrop-filter: blur(10px);
      border: 1px solid rgba(148,163,184,0.15); border-radius: 10px;
      box-shadow: 0 8px 24px rgba(0,0,0,0.5);
    }}
    #toolbar button {{
      display: inline-flex; align-items: center; gap: 5px;
      padding: 6px 14px; border-radius: 6px;
      border: 1px solid rgba(148,163,184,0.2);
      background: rgba(255,255,255,0.04); color: #cbd5e1;
      cursor: pointer; font-family: inherit; font-size: 12px; font-weight: 600;
      transition: all .15s ease;
    }}
    #toolbar button:hover {{
      background: #3b82f6; border-color: #3b82f6; color: #fff;
      transform: translateY(-1px);
    }}
    #toolbar button.btn-accent {{
      background: rgba(251, 191, 36, 0.1);
      border-color: rgba(251, 191, 36, 0.3);
      color: #fbbf24;
    }}
    #toolbar button.btn-accent:hover {{
      background: #fbbf24; color: #0f172a; border-color: #fbbf24;
    }}
    
    #visualization {{
      width: 100vw; height: 100vh; padding-top: 56px;
    }}

    /* ── CURSOR ── */
    .vis-timeline, .vis-timeline .vis-center,
    .vis-timeline .vis-content, .vis-timeline .vis-panel {{
      cursor: grab !important;
    }}
    .vis-timeline:active, .vis-timeline.vis-move,
    .vis-timeline.vis-move .vis-center,
    .vis-timeline.vis-move .vis-content {{
      cursor: grabbing !important;
    }}

    /* ── TOOLTIP FIX (Z-INDEX + OVERFLOW) ── */
    .vis-timeline, .vis-panel.vis-center, .vis-itemset {{
      overflow: visible !important;
    }}
    .vis-tooltip {{
      background: #1e293b !important;
      border: 1px solid rgba(148,163,184,0.2) !important;
      border-radius: 10px !important; color: #f1f5f9 !important;
      box-shadow: 0 12px 40px rgba(0,0,0,0.6) !important;
      padding: 12px 16px !important;
      font-family: 'Segoe UI', sans-serif !important;
      font-size: 12.5px !important; line-height: 1.6 !important;
      max-width: 380px;
      z-index: 1000000 !important;
      position: fixed !important;
      pointer-events: none;
    }}

    /* ── EJES Y GRID (MÁS VISIBLES) ── */
    .vis-time-axis .vis-text {{
      color: #64748b !important;
      font-family: 'JetBrains Mono', 'Consolas', monospace !important;
      font-size: 10px;
    }}
    .vis-time-axis .vis-text.vis-major {{
      color: #94a3b8 !important; font-size: 11px; font-weight: 600;
      text-transform: capitalize;
    }}
    .vis-time-axis .vis-grid.vis-vertical {{
      border-color: rgba(148, 163, 184, 0.15) !important; /* Más visible */
    }}
    .vis-time-axis .vis-grid.vis-minor {{
      border-color: rgba(148, 163, 184, 0.08) !important;
    }}

    /* ── FIX TRANSICIONES RESIDUALES ── */
    .vis-group,
    .vis-label,
    .vis-inner,
    .vis-nested-group,
    .vis-label.vis-nesting-group {{
      transition: none !important;
      animation: none !important;
    }}

    /* ── GRUPOS ── */
    .vis-labelset .vis-label {{
      border-bottom: 1px solid rgba(148,163,184,0.08);
    }}
    .vis-labelset .vis-label:hover {{
      background: rgba(255, 255, 255, 0.03);
    }}
    .vis-labelset .vis-label .vis-inner {{
      padding: 5px 8px;
    }}

    /* ── ITEMS (RANGO EXACTO) ── */
    .vis-item {{
      border-radius: 5px; border-width: 1.5px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.35);
    }}
    .vis-item .vis-item-content {{
      padding: 3px 0 !important; /* Rango exacto: sin padding horizontal */
      text-align: center;
      width: 100%;
    }}
    .vis-item.vis-selected {{ z-index: 100; }}

    {cssRules}
  </style>
</head>
<body>

<div id=""toolbar"">
  <button class=""btn-accent"" onclick=""toggleAllNested(true)"" id=""btn-expand-all"">⏵ Expandir todo</button>
  <button class=""btn-accent"" onclick=""toggleAllNested(false)"" id=""btn-collapse-all"">⏷ Colapsar todo</button>
  <div style=""width:1px; height:20px; background:rgba(148,163,184,0.15); margin:0 4px""></div>
  <button onclick=""fitAll()"">⟲ Ver todo</button>
  <button onclick=""zoomIn()"">🔍+</button>
  <button onclick=""zoomOut()"">🔍−</button>
</div>

<div id=""visualization""></div>

<!-- Scripts Locales -->
<script src=""{visDataJs}""></script>
<script src=""{visTimelineJs}""></script>
<script src=""{lucideJs}""></script>

<script>
  if (typeof lucide !== 'undefined') {{
    lucide.createIcons();
  }}

  // 1. Truncar nombres de docentes a 15 caracteres (en JS para asegurar consistencia si el C# falla)
  function truncarNombre(nombre, max = 15) {{
    const limpio = nombre.replace(/^👤\s*/, ''); 
    if (limpio.length <= max) return '👤 ' + limpio;
    return '👤 ' + limpio.substring(0, max).trim() + '…';
  }}

  window.__TIMELINE_DATA__ = {{
    items:  {itemsJson},
    groups: {groupsJson}.map(g => {{
      if (g.treeLevel === 1 && g.content) {{
        g.content = truncarNombre(g.content);
      }}
      return g;
    }})
  }};

  if (typeof vis !== 'undefined') {{
    var data      = window.__TIMELINE_DATA__;
    var container = document.getElementById('visualization');

    // Usar DataSet para grupos para que expandir/colapsar sea reactivo
    var itemsDS  = new vis.DataSet(data.items);

    // ── Definiciones base (INMUTABLES)
    var CHILD_GROUPS = [];
    var PARENT_GROUPS = [];

    data.groups.forEach(function(g) {{
      if (g.treeLevel === 1 && g.nestedGroups && g.nestedGroups.length > 0) {{
        PARENT_GROUPS.push(g);
      }} else {{
        CHILD_GROUPS.push(g);
      }}
    }});

    // ── Estado canónico propio
    var parentShowNested = {{}};
    PARENT_GROUPS.forEach(function(g) {{
      parentShowNested[g.id] = g.showNested === true;
    }});

    // ── Función que construye siempre un DataSet nuevo
    function buildGroupDataSet() {{
      var rows = [];
      CHILD_GROUPS.forEach(function(g) {{
        rows.push(Object.assign({{}}, g));           // hijos sin modificar
      }});
      PARENT_GROUPS.forEach(function(g) {{
        var copy = Object.assign({{}}, g);
        copy.showNested = parentShowNested[g.id] === true;  // explícito
        rows.push(copy);
      }});
      return new vis.DataSet(rows);                  // DataSet NUEVO
    }}

    var options = {{
      locale: 'es',
      start: '{startStr}',
      end:   '{endStr}',

      /* Escala: días numerados pequeños */
      timeAxis: {{ scale: 'day', step: 1 }},
      format: {{
        minorLabels: {{ day: 'D' }},
        majorLabels: {{ month: 'MMMM YYYY', day: 'dddd D MMMM' }}
      }},

      stack: true,
      showCurrentTime: true,
      zoomKey: 'ctrlKey',
      moveable: true,
      horizontalScroll: true,
      verticalScroll: true,
      orientation: {{ axis: 'top', item: 'top' }},
      groupOrder: 'content',

      zoomMin:  1 * 24 * 60 * 60 * 1000,   // 1 día mínimo
      zoomMax:  365 * 24 * 60 * 60 * 1000, // 1 año máximo

      margin: {{ 
        item: {{ horizontal: 0, vertical: 3 }}, // Rango exacto: horizontal 0
        axis: 6 
      }},
      tooltip: {{ followMouse: true, overflowMethod: 'flip' }}
    }};

    var timeline = new vis.Timeline(container, itemsDS, buildGroupDataSet(), options);

    timeline.fit({{ animation: {{ duration: 700, easingFunction: 'easeInOutQuad' }} }});

    // ── FUNCIONES DE CONTROL ──
    window.fitAll  = function() {{ timeline.fit({{ animation: true }}); }};
    window.zoomIn  = function() {{ timeline.zoomIn(0.35); }};
    window.zoomOut = function() {{ timeline.zoomOut(0.35); }};

    window.toggleAllNested = function(visible) {{
      var w = timeline.getWindow();                  // guardar ventana
      Object.keys(parentShowNested).forEach(function(id) {{
          parentShowNested[id] = visible;           // actualizar estado
      }});
      timeline.setGroups(buildGroupDataSet());       // reconstruir DOM
      timeline.setWindow(w.start, w.end, {{ animation: false }}); // restaurar
    }};

    // ── TOOLTIP MUTATION OBSERVER ──
    const tooltipObserver = new MutationObserver((mutations) => {{
      mutations.forEach((mutation) => {{
        mutation.addedNodes.forEach((node) => {{
          if (node.nodeType === 1 && node.classList.contains('vis-tooltip')) {{
            node.style.zIndex = '1000000';
            node.style.position = 'fixed';
          }}
        }});
      }});
    }});
    tooltipObserver.observe(document.body, {{ childList: true, subtree: true }});

    // Refuerzo periódico
    setInterval(() => {{
      document.querySelectorAll('.vis-tooltip').forEach(tt => {{
        tt.style.zIndex = '1000000';
        tt.style.position = 'fixed';
      }});
    }}, 500);

    // Atajos
    document.addEventListener('keydown', e => {{
      if (e.altKey && e.key.toLowerCase() === 'e') toggleAllNested(true);
      if (e.altKey && e.key.toLowerCase() === 'c') toggleAllNested(false);
      if (e.key.toLowerCase() === 'f' && !e.ctrlKey) fitAll();
    }});
  }}
</script>
</body>
</html>";
        }
    }
}
