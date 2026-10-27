using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Formulario modal especializado para la selección y filtrado de terceros y docentes
    /// en el generador de horarios por Algoritmo Genético.
    /// </summary>
    public partial class AgFormTerceros : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly HashSet<int> _seleccionados;
        private List<AgTerceroSeleccionDto> _todosLosTerceros = new List<AgTerceroSeleccionDto>();
        private BindingList<AgTerceroSeleccionDto> _vistaFiltrada;

        public IReadOnlyCollection<int> TercerosSeleccionadosIds => _seleccionados.ToList();

        public AgFormTerceros(Func<IDbConnection> connectionFactory, IEnumerable<int> seleccionInicial = null)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _seleccionados = seleccionInicial != null ? new HashSet<int>(seleccionInicial) : new HashSet<int>();

            InitializeComponent();
            ConfigurarEventos();
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Fase 1471: dejar explicito que "habilitadas" se refiere a la subunidad activa.
            string subActiva = (SyncContext.ActiveSubunidad ?? string.Empty).Trim();
            chkSoloHabilitadas.Text = string.IsNullOrWhiteSpace(subActiva)
                ? "Solo asignaturas habilitadas"
                : $"Solo asignaturas habilitadas en {subActiva}";

            await CargarDatosAsync();
        }

        private void ConfigurarEventos()
        {
            txtBuscar.TextChanged += (s, e) => AplicarFiltros();
            // Fase 1444: Evitar que Enter en el buscador dispare AcceptButton cerrando el diálogo prematuramente
            txtBuscar.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.SuppressKeyPress = true;
                AplicarFiltros();
            };

            // Fase 1444: En la grilla, Enter confirma edición de celda y no el diálogo
            dgvTerceros.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) e.SuppressKeyPress = true;
            };

            cmbTerritorial.SelectedIndexChanged += (s, e) => AplicarFiltros();
            chkSoloHabilitadas.CheckedChanged += (s, e) => AplicarFiltros();
            chkMostrarInactivos.CheckedChanged += (s, e) => AplicarFiltros();

            btnMarcarTodos.Click += (s, e) => MarcarVisibles(true);
            btnDesmarcarTodos.Click += (s, e) => MarcarVisibles(false);

            dgvTerceros.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvTerceros.IsCurrentCellDirty)
                {
                    dgvTerceros.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };

            // Fase 1427: Leer directamente el valor de la celda de selección para evitar desfase del DTO en modo enlazado
            dgvTerceros.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex != colSeleccion.Index) return;
                if (!(dgvTerceros.Rows[e.RowIndex].DataBoundItem is AgTerceroSeleccionDto dto)) return;

                bool marcado = dgvTerceros.Rows[e.RowIndex].Cells[e.ColumnIndex].Value is bool b && b;
                dto.Seleccionado = marcado;
                if (marcado) _seleccionados.Add(dto.Id); else _seleccionados.Remove(dto.Id);
                ActualizarConteo();
            };

            // Fase 1471: trazabilidad de selecciones que no entraran al pool AG (sin habilitacion local).
            btnAceptar.Click += (s, e) =>
            {
                var fueraDePool = _todosLosTerceros
                    .Where(t => t.Seleccionado && t.HabilitacionesLocales == 0)
                    .Select(t => t.Id)
                    .ToList();

                if (fueraDePool.Count > 0)
                {
                    LoggerHelper.Warning($"[AgFormTerceros] {fueraDePool.Count} seleccionados sin habilitacion local: [{string.Join(",", fueraDePool)}]");
                }
            };
        }

        private async Task CargarDatosAsync()
        {
            try
            {
                // Fase 1471: la lista de terceros es multi-territorial (igual que FormTerceros).
                // El pool AG sigue acotado a la subunidad activa en AgFormGeneradorHorarios (L912-936).
                string sub = (SyncContext.ActiveSubunidad ?? string.Empty).Trim();

                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    const string sqlTerceros = @"
SELECT id AS Id, nombre AS Nombre, cedula AS Cedula, tipo AS Tipo,
       categoria AS Categoria, departamento AS Departamento, residencia AS Residencia,
       subunidad AS Subunidad, activo AS Activo
FROM terceros
WHERE (cedula != '1111111111' AND UPPER(TRIM(nombre)) != 'VACANTE')
ORDER BY nombre";

                    var terceros = (await db.QueryAsync<Tercero>(sqlTerceros)).ToList();

                    string filtroSubLocal = string.IsNullOrWhiteSpace(sub)
                        ? string.Empty
                        : " AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)";

                    string sqlHabilitadosLocal = @"
SELECT idtercero AS IdTercero, COUNT(*) AS Total
FROM terceros_asignaturas
WHERE activo = 1" + filtroSubLocal + @"
GROUP BY idtercero";

                    string sqlHabilitadosOtras = @"
SELECT idtercero AS IdTercero, COUNT(*) AS Total
FROM terceros_asignaturas
WHERE activo = 1
  AND subunidad IS NOT NULL AND TRIM(subunidad) <> '' AND subunidad <> @Sub
GROUP BY idtercero";

                    var locales = (await db.QueryAsync<ConteoHabilitaciones>(sqlHabilitadosLocal, new { Sub = sub }))
                        .ToDictionary(c => c.IdTercero, c => c.Total);

                    // Fase 1475: Si no hay subunidad activa, locales ya incluye todo el padrón; no duplicar como otras
                    var otras = string.IsNullOrWhiteSpace(sub)
                        ? new Dictionary<int, int>()
                        : (await db.QueryAsync<ConteoHabilitaciones>(sqlHabilitadosOtras, new { Sub = sub }))
                            .ToDictionary(c => c.IdTercero, c => c.Total);

                    const string sqlTerritoriales = @"
SELECT DISTINCT subunidad
FROM territoriales
WHERE subunidad IS NOT NULL AND TRIM(subunidad) <> ''";

                    var catalogo = (await db.QueryAsync<string>(sqlTerritoriales)).ToList();

                    _todosLosTerceros = terceros.Select(t => new AgTerceroSeleccionDto
                    {
                        Id = t.Id,
                        Nombre = t.Nombre,
                        Cedula = t.Cedula,
                        Tipo = t.Tipo,
                        Categoria = t.Categoria,
                        Departamento = t.Departamento,
                        Residencia = t.Residencia,
                        Subunidad = t.Subunidad,
                        Activo = t.Activo == 1,
                        HabilitacionesLocales = locales.TryGetValue(t.Id, out int nLocal) ? nLocal : 0,
                        HabilitacionesOtras = otras.TryGetValue(t.Id, out int nOtras) ? nOtras : 0,
                        Seleccionado = _seleccionados.Contains(t.Id)
                    }).ToList();

                    var valoresTerceros = _todosLosTerceros
                        .SelectMany(t => AgTerritorialCatalog.Separar(t.Departamento)
                            .Concat(AgTerritorialCatalog.Separar(t.Subunidad)));

                    var territoriales = AgTerritorialCatalog.Construir(catalogo, valoresTerceros);

                    cmbTerritorial.Items.Clear();
                    cmbTerritorial.Items.Add(AgTerritorialCatalog.Todas);
                    foreach (var ter in territoriales)
                    {
                        cmbTerritorial.Items.Add(ter);
                    }
                    cmbTerritorial.SelectedIndex = 0;

                    LoggerHelper.Info($"[AgFormTerceros] terceros={_todosLosTerceros.Count}; territoriales={territoriales.Count}; habilitadosLocales({(string.IsNullOrWhiteSpace(sub) ? "sin-subunidad" : sub)})={locales.Count}; habilitadosOtras={otras.Count}");
                }

                AplicarFiltros();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormTerceros] Error cargando datos de terceros", ex);
                MessageBox.Show(this, $"Error al cargar lista de terceros: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AplicarFiltros()
        {
            string territorial = cmbTerritorial.SelectedIndex > 0 ? cmbTerritorial.SelectedItem?.ToString() : null;
            string busqueda = txtBuscar.Text;
            bool soloHabilitadas = chkSoloHabilitadas.Checked;
            bool mostrarInactivos = chkMostrarInactivos.Checked;

            var filtrados = AgTerceroFiltro.Filtrar(
                _todosLosTerceros,
                territorial,
                busqueda,
                soloHabilitadas,
                mostrarInactivos).ToList();

            _vistaFiltrada = new BindingList<AgTerceroSeleccionDto>(filtrados);
            dgvTerceros.DataSource = _vistaFiltrada;
            ActualizarConteo();
        }

        private void MarcarVisibles(bool marcar)
        {
            if (_vistaFiltrada == null) return;

            foreach (var item in _vistaFiltrada)
            {
                item.Seleccionado = marcar;
                if (marcar)
                {
                    _seleccionados.Add(item.Id);
                }
                else
                {
                    _seleccionados.Remove(item.Id);
                }
            }

            // Fase 1428: En modo enlazado, actualizar explícitamente el valor de la celda de selección e invalidar la grilla para refrescar los checks
            foreach (DataGridViewRow row in dgvTerceros.Rows)
            {
                if (row.DataBoundItem is AgTerceroSeleccionDto dto)
                {
                    row.Cells[colSeleccion.Index].Value = dto.Seleccionado;
                }
            }
            dgvTerceros.Invalidate();
            ActualizarConteo();
        }

        private void ActualizarConteo()
        {
            int total = _todosLosTerceros?.Count ?? 0;
            int seleccionados = _seleccionados.Count;
            int visibles = _vistaFiltrada?.Count ?? 0;
            int conLocal = _todosLosTerceros?.Count(t => t.HabilitacionesLocales > 0) ?? 0;
            int soloOtras = _todosLosTerceros?.Count(t => t.HabilitacionesLocales == 0 && t.HabilitacionesOtras > 0) ?? 0;

            lblConteo.Text = $"{seleccionados} seleccionados (visibles: {visibles} de {total}) | habilitados locales: {conLocal} | solo en otras: {soloOtras}";
        }

        /// <summary>Fila de conteo de habilitaciones agrupadas por tercero (Fase 1471).</summary>
        private sealed class ConteoHabilitaciones
        {
            public int IdTercero { get; set; }
            public int Total { get; set; }
        }
    }
}
