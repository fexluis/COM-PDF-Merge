using System;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormConfig : Form
    {
        private readonly EncryptedSQLiteContext _context;
        private readonly ConfigRepository _repo;
        private Config _currentConfig;

        /// <summary>
        /// Inicializa el formulario. 
        /// Fase 16/12: El contexto _context es el singleton de SqliteHelper; NO debe ser dispuesto por este form.
        /// Cambio (Fase 1 de refinamiento): Layout estándar con dropdowns de base de datos.
        /// </summary>
        public FormConfig()
        {
            InitializeComponent();
            _context = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext();
            _repo = new ConfigRepository(_context);
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            SetupSyncIndicatorUI();
            CargarConfiguracion();
        }

        private void SetupSyncIndicatorUI()
        {
            bool isSyncing = false;
            string syncText = "";

            try
            {
                var resolver = new ConversorPDF.XLL.Core.Sync.SyncPathResolver();
                string path = resolver.GetSyncRootFolder();
                isSyncing = true;
                syncText = $"Sincronizando: {path}";
            }
            catch (Exception)
            {
                isSyncing = false;
                syncText = "Sincronización Inactiva: Ruta csync no encontrada.";
            }

            // Draw status circle in panel1
            this.panel1.Paint += (s, ev) =>
            {
                using (var brush = new System.Drawing.SolidBrush(isSyncing ? System.Drawing.Color.LimeGreen : System.Drawing.Color.Red))
                using (var pen = new System.Drawing.Pen(System.Drawing.Color.DarkGray))
                {
                    ev.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    ev.Graphics.FillEllipse(brush, 10, 18, 12, 12);
                    ev.Graphics.DrawEllipse(pen, 10, 18, 12, 12);
                }
            };

            // Add label with path text next to circle
            var lblSyncPath = new Label
            {
                Text = syncText,
                AutoSize = true,
                Location = new System.Drawing.Point(28, 17),
                ForeColor = System.Drawing.Color.DimGray,
                Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Regular)
            };
            this.panel1.Controls.Add(lblSyncPath);

            // Force redraw of panel
            this.panel1.Invalidate();
        }

        private void CargarConfiguracion()
        {
            var db = _context.Open();

            // Llenar combo Periodos
            // Corrección: Se utiliza ConversorPDF.XLL.Academic.Domain.Models.LookupItem ya que el namespace Application.DTOs no contiene este modelo.
            var periodos = db.Query<ConversorPDF.XLL.Academic.Domain.Models.LookupItem>(
                "SELECT idperiodo as Id, periodo as DisplayName FROM periodos ORDER BY idperiodo DESC").ToList();
            cmbPeriodos.DataSource = periodos;
            cmbPeriodos.DisplayMember = "DisplayName";
            cmbPeriodos.ValueMember = "Id";

            // Llenar combo Subunidad
            var subunidades = db.Query<string>("SELECT DISTINCT subunidad FROM territoriales WHERE subunidad IS NOT NULL AND subunidad != '' ORDER BY subunidad").ToList();

            // Cambio Fase 14: Suspender evento para evitar queries prematuras y race conditions
            cmbCiudadBase.SelectedIndexChanged -= cmbCiudadBase_SelectedIndexChanged;
            cmbSubunidad.SelectedIndexChanged -= cmbSubunidad_SelectedIndexChanged;

            cmbSubunidad.DataSource = subunidades;
            cmbSubunidad.Enabled = false; // La subunidad de sincronización es de solo lectura en esta pantalla

            // Llenar combo Aeropuerto Base
            var terminales = db.Query<string>("SELECT DISTINCT terminal FROM terminales WHERE id > 1 AND terminal IS NOT NULL AND terminal != '' ORDER BY terminal").ToList();
            cmbAeropuertoBase.DataSource = terminales;

            // Llenar combo Ciudad Base
            var ciudades = db.Query<string>("SELECT DISTINCT ciudadSiif FROM sedes WHERE ciudadSiif IS NOT NULL AND ciudadSiif != '' ORDER BY ciudadSiif").ToList();
            cmbCiudadBase.DataSource = ciudades;

            var result = _repo.GetConfig();
            if (result.IsSuccess)
            {
                _currentConfig = result.Value;
                MapToUi();
            }
            else
            {
                MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGuardar.Enabled = false;
            }

            // Fase 11: Reactivar eventos incondicionalmente para mantener UI funcional
            cmbCiudadBase.SelectedIndexChanged += cmbCiudadBase_SelectedIndexChanged;
            cmbSubunidad.SelectedIndexChanged += cmbSubunidad_SelectedIndexChanged;
        }

        private void MapToUi()
        {
            // Fase 3 y 6: SelectedValue falla si Dapper mapea INTEGER de SQLite como long (Fase 6)
            // o si el valor no existe lanzando InvalidOperationException (Fase 3).
            // Se itera manualmente comparando item.Id (int) con _currentConfig.IdPeriodo.
            if (_currentConfig.IdPeriodo > 0)
            {
                foreach (ConversorPDF.XLL.Academic.Domain.Models.LookupItem item in cmbPeriodos.Items)
                {
                    if (item.Id == _currentConfig.IdPeriodo)
                    {
                        cmbPeriodos.SelectedItem = item;
                        break;
                    }
                }
            }

            // Cambio Fase 25: Prevenir ArgumentNullException si la propiedad es nula y se evalúa Contains()
            if (!string.IsNullOrEmpty(_currentConfig.Subunidad) && cmbSubunidad.Items.Contains(_currentConfig.Subunidad))
            {
                cmbSubunidad.SelectedItem = _currentConfig.Subunidad;
                // Fase 15: Sincronizar txtIdTerritorial con BD al cargar para asegurar consistencia con la tabla territoriales
                var db = _context.Open();
                var id = db.QueryFirstOrDefault<int>("SELECT id FROM territoriales WHERE subunidad = @Sub LIMIT 1", new { Sub = _currentConfig.Subunidad });
                txtIdTerritorial.Text = id > 0 ? id.ToString() : _currentConfig.IdTerritorial.ToString();
            }
            else
            {
                txtIdTerritorial.Text = _currentConfig.IdTerritorial.ToString();
            }
            // Prevención de ArgumentNullException con strings nulos provenientes de BD.
            txtFechaInicio.Text = _currentConfig.FechaInicio ?? string.Empty;
            txtFechaFin.Text = _currentConfig.FechaFin ?? string.Empty;
            txtCdp.Text = _currentConfig.Cdp ?? string.Empty;
            // Formato "0" evita notación científica (ej. "1,5E+07") que rompe la validación KeyPress.
            txtSalario.Text = _currentConfig.SalarioDecreto6132025.ToString("0");
            txtViaticoA.Text = _currentConfig.ViaticoDiarioTarifaA.ToString("0");
            txtViaticoB.Text = _currentConfig.ViaticoDiarioTarifaB.ToString("0");
            txtPorcentajeA.Text = _currentConfig.PorcentajeDecretoA.ToString("0.0000");
            txtPorcentajeB.Text = _currentConfig.PorcentajeDecretoB.ToString("0.0000");
            txtDtNombre.Text = _currentConfig.DtNombre ?? string.Empty;
            txtDtCargo.Text = _currentConfig.DtCargo ?? string.Empty;
            txtCorpCedula.Text = _currentConfig.CorporativoCedula ?? string.Empty;
            txtCorpNombre.Text = _currentConfig.CorporativoNombre ?? string.Empty;
            txtCorpCargo.Text = _currentConfig.CorporativoCargo ?? string.Empty;

            if (!string.IsNullOrEmpty(_currentConfig.CiudadBase) && cmbCiudadBase.Items.Contains(_currentConfig.CiudadBase))
                cmbCiudadBase.SelectedItem = _currentConfig.CiudadBase;

            txtDeptoBase.Text = _currentConfig.DepartamentoBase ?? string.Empty;

            if (!string.IsNullOrEmpty(_currentConfig.AeropuertoBase) && cmbAeropuertoBase.Items.Contains(_currentConfig.AeropuertoBase))
                cmbAeropuertoBase.SelectedItem = _currentConfig.AeropuertoBase;
        }

        /// <summary>
        /// Mapea los valores de la UI a la entidad de dominio Config.
        /// Usa TryParse para evitar silenciar fallos; asigna 0 si el campo está vacío/inválido.
        /// </summary>
        private void MapToConfig()
        {
            // Fase 6: Evitar falso negativo por boxing. SQLite INTEGER → long en Dapper; TryParse es seguro.
            if (cmbPeriodos.SelectedItem is ConversorPDF.XLL.Academic.Domain.Models.LookupItem selectedPeriodo)
                _currentConfig.IdPeriodo = selectedPeriodo.Id;

            _currentConfig.Subunidad = cmbSubunidad.SelectedItem?.ToString();
            _currentConfig.IdTerritorial = int.TryParse(txtIdTerritorial.Text, out int idTerr) ? idTerr : 0;
            _currentConfig.FechaInicio = txtFechaInicio.Text;
            _currentConfig.FechaFin = txtFechaFin.Text;
            _currentConfig.Cdp = txtCdp.Text;

            _currentConfig.SalarioDecreto6132025 = double.TryParse(txtSalario.Text, out double sal) ? sal : 0;
            _currentConfig.ViaticoDiarioTarifaA = double.TryParse(txtViaticoA.Text, out double viaA) ? viaA : 0;
            _currentConfig.ViaticoDiarioTarifaB = double.TryParse(txtViaticoB.Text, out double viaB) ? viaB : 0;
            _currentConfig.PorcentajeDecretoA = double.TryParse(txtPorcentajeA.Text, out double porA) ? porA : 0;
            _currentConfig.PorcentajeDecretoB = double.TryParse(txtPorcentajeB.Text, out double porB) ? porB : 0;

            _currentConfig.DtNombre = txtDtNombre.Text;
            _currentConfig.DtCargo = txtDtCargo.Text;
            _currentConfig.CorporativoCedula = txtCorpCedula.Text;
            _currentConfig.CorporativoNombre = txtCorpNombre.Text;
            _currentConfig.CorporativoCargo = txtCorpCargo.Text;
            _currentConfig.CiudadBase = cmbCiudadBase.SelectedItem?.ToString();
            _currentConfig.DepartamentoBase = txtDeptoBase.Text;
            _currentConfig.AeropuertoBase = cmbAeropuertoBase.SelectedItem?.ToString();
        }

        private void cmbCiudadBase_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ciudad = cmbCiudadBase.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(ciudad))
            {
                var db = _context.Open();
                var depto = db.QueryFirstOrDefault<string>("SELECT departamentoSiif FROM sedes WHERE ciudadSiif = @Ciudad LIMIT 1", new { Ciudad = ciudad });
                // Cambio Fase 18: Coalescencia nula para evitar ArgumentNullException si no se encuentra la ciudad.
                txtDeptoBase.Text = depto ?? string.Empty;
            }
        }

        private void cmbSubunidad_SelectedIndexChanged(object sender, EventArgs e)
        {
            var sub = cmbSubunidad.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(sub))
            {
                var db = _context.Open();
                var id = db.QueryFirstOrDefault<int>("SELECT id FROM territoriales WHERE subunidad = @Sub LIMIT 1", new { Sub = sub });
                txtIdTerritorial.Text = id > 0 ? id.ToString() : "0";
            }
        }

        private void NumerosEnteros_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void NumerosDecimales_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // Solo permitir una coma (Fase 22: permite reemplazar selección si incluye la coma)
            var txt = sender as TextBox;
            if (e.KeyChar == ',' && txt != null)
            {
                // Cambio Fase 24: Math.Min protege contra ArgumentOutOfRangeException en edge cases de selección
                var selStart = Math.Min(txt.SelectionStart, txt.Text.Length);
                var selLen = Math.Min(txt.SelectionLength, txt.Text.Length - selStart);
                var textOutside = txt.Text.Remove(selStart, selLen);

                if (textOutside.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// Valida el formato decimal al perder el foco. Si es inválido, muestra una advertencia y no actualiza.
        /// Cambio: Añadida advertencia para entradas inválidas, evitando parseos silenciosamente fallidos.
        /// </summary>
        private void Decimal_Leave(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt != null)
            {
                if (double.TryParse(txt.Text, out double val))
                {
                    txt.Text = val.ToString("0.0000");
                }
                else if (!string.IsNullOrWhiteSpace(txt.Text))
                {
                    MessageBox.Show($"\"{txt.Text}\" no es un número decimal válido. Use coma como separador.",
                        "Formato inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    // Cambio Fase 21: BeginInvoke difiere la recuperación del foco al siguiente ciclo de mensajes
                    // previniendo un bucle modal infinito de re-disparos del evento Leave.
                    txt.BeginInvoke(new Action(() => { txt.Focus(); txt.SelectAll(); }));
                }
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (_currentConfig == null) return;

            MapToConfig();

            var result = _repo.UpdateConfig(_currentConfig);
            if (result.IsSuccess)
            {
                MessageBox.Show("Configuración guardada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Fase 12: No disponer _context porque es un singleton compartido (SqliteHelper)
            base.OnFormClosing(e);
        }
    }
}
