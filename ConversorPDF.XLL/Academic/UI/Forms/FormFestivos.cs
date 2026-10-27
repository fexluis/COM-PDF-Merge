using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormFestivos : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private BindingSource _bs = new BindingSource();
        private Festivo _currentFestivo;
        private List<Periodo> _periodos = new List<Periodo>();

        public FormFestivos(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
            InitializeComponent();
            SetupGrid();
            CargarPeriodos();
            CargarFestivos();
        }

        private void SetupGrid()
        {
            dgvFestivos.AutoGenerateColumns = false;
            dgvFestivos.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DetalleFestivo", HeaderText = "Detalle / Descripción", Width = 150 });
            dgvFestivos.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "TipoFestivo", HeaderText = "Tipo", Width = 100 });
            
            dgvFestivos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                DataPropertyName = "FechaFestivoInicio", 
                HeaderText = "Fecha Inicio", 
                Width = 95,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
            });
            
            dgvFestivos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                DataPropertyName = "FechaFestivoFinal", 
                HeaderText = "Fecha Fin", 
                Width = 95,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
            });
            
            // Columna virtual/calculada o lookup para mostrar el periodo legible
            var colPeriodo = new DataGridViewTextBoxColumn { HeaderText = "Periodo", Width = 80 };
            dgvFestivos.Columns.Add(colPeriodo);

            dgvFestivos.DataSource = _bs;
            dgvFestivos.SelectionChanged += DgvFestivos_SelectionChanged;
            dgvFestivos.CellFormatting += DgvFestivos_CellFormatting;
        }

        private void CargarPeriodos()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetPeriodos();
                if (result.IsSuccess)
                {
                    _periodos = result.Value.ToList();
                    cmbPeriodo.DataSource = new BindingSource(_periodos, null);
                    cmbPeriodo.DisplayMember = "PeriodoNombre";
                    cmbPeriodo.ValueMember = "IdPeriodo";
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al cargar períodos en FormFestivos", ex);
            }
        }

        private void CargarFestivos()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetFestivos();
                if (result.IsSuccess)
                {
                    _bs.DataSource = result.Value.ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar festivos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvFestivos_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvFestivos.CurrentRow?.DataBoundItem is Festivo f)
            {
                _currentFestivo = f;
                cmbPeriodo.SelectedValue = f.IdPeriodo;
                dtpInicio.Value = f.FechaFestivoInicio;
                dtpFin.Value = f.FechaFestivoFinal;
                txtDetalle.Text = f.DetalleFestivo;
                cmbTipo.SelectedItem = f.TipoFestivo ?? "FESTIVO";
                btnEliminar.Enabled = true;
            }
            else
            {
                LimpiarCampos();
            }
        }

        private void DgvFestivos_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 4 && e.RowIndex >= 0) // Columna Periodo
            {
                if (dgvFestivos.Rows[e.RowIndex].DataBoundItem is Festivo f)
                {
                    var per = _periodos.FirstOrDefault(p => p.IdPeriodo == f.IdPeriodo);
                    e.Value = per != null ? per.PeriodoNombre : f.IdPeriodo.ToString();
                    e.FormattingApplied = true;
                }
            }
        }

        private void LimpiarCampos()
        {
            _currentFestivo = null;
            if (cmbPeriodo.Items.Count > 0) cmbPeriodo.SelectedIndex = 0;
            if (cmbTipo.Items.Count > 0) cmbTipo.SelectedIndex = 0;
            dtpInicio.Value = DateTime.Today;
            dtpFin.Value = DateTime.Today;
            txtDetalle.Clear();
            btnEliminar.Enabled = false;
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
            txtDetalle.Focus();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDetalle.Text))
            {
                MessageBox.Show("Detalle/Descripción es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbPeriodo.SelectedValue == null)
            {
                MessageBox.Show("Debe seleccionar un período válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dtpInicio.Value.Date > dtpFin.Value.Date)
            {
                MessageBox.Show("La fecha de inicio no puede ser posterior a la fecha de fin.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var f = _currentFestivo ?? new Festivo();
            f.IdPeriodo = (int)cmbPeriodo.SelectedValue;
            f.FechaFestivoInicio = dtpInicio.Value.Date;
            f.FechaFestivoFinal = dtpFin.Value.Date;
            f.DetalleFestivo = txtDetalle.Text.Trim();
            f.TipoFestivo = cmbTipo.SelectedItem?.ToString() ?? "FESTIVO";

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                
                if (f.IdFestivo == 0) 
                {
                    var result = repo.InsertFestivo(f);
                    if (result.IsFailure) { MessageBox.Show($"Error al insertar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    f.IdFestivo = result.Value;
                }
                else 
                {
                    var result = repo.UpdateFestivo(f);
                    if (result.IsFailure) { MessageBox.Show($"Error al actualizar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                }

                // Nota: La persistencia al Vault ahora la gestiona el SyncScheduler de forma asíncrona.
                
                MessageBox.Show("Festivo guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarFestivos();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentFestivo == null) return;

            if (MessageBox.Show($"¿Está seguro de eliminar el festivo de '{_currentFestivo.DetalleFestivo}'?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var db = _connectionFactory();
                    var repo = new AcademicRepository(db);
                    var result = repo.DeleteFestivo(_currentFestivo.IdFestivo);
                    if (result.IsFailure) { MessageBox.Show($"Error al eliminar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                    CargarFestivos();
                    LimpiarCampos();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al eliminar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
