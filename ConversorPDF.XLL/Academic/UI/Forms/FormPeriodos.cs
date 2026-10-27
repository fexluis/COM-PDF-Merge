using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormPeriodos : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly BindingSource _bs = new BindingSource();
        private PeriodoGridDto _currentPeriodo;

        public FormPeriodos(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            InitializeComponent();
            SetupGrid();
            CargarPeriodos();
        }

        private void SetupGrid()
        {
            dgvPeriodos.AutoGenerateColumns = false;
            dgvPeriodos.Columns.Clear();

            dgvPeriodos.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IdPeriodo",
                HeaderText = "ID",
                Width = 50,
                Visible = false
            });

            dgvPeriodos.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "PeriodoNombre",
                HeaderText = "Código Período",
                Width = 110
            });

            dgvPeriodos.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "PeriodoInicio",
                HeaderText = "Inicio (ISO)",
                Width = 100
            });

            dgvPeriodos.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "PeriodoFin",
                HeaderText = "Fin (ISO)",
                Width = 100
            });

            dgvPeriodos.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "EstadoLegible",
                HeaderText = "Estado",
                Width = 80
            });

            dgvPeriodos.DataSource = _bs;
            dgvPeriodos.SelectionChanged += DgvPeriodos_SelectionChanged;
        }

        private void CargarPeriodos()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var service = new PeriodoCrudService(repo);
                var result = service.ObtenerTodos();

                if (result.IsSuccess)
                {
                    _bs.DataSource = result.Value.ToList();
                    LimpiarFormulario();
                }
                else
                {
                    MessageBox.Show(result.Error, "Error al cargar períodos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormPeriodos] Error al cargar períodos", ex);
                MessageBox.Show($"Error al cargar períodos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvPeriodos_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPeriodos.SelectedRows.Count > 0)
            {
                _currentPeriodo = dgvPeriodos.SelectedRows[0].DataBoundItem as PeriodoGridDto;
                CargarDetalle(_currentPeriodo);
            }
        }

        private void CargarDetalle(PeriodoGridDto dto)
        {
            if (dto == null) return;

            txtCodigo.Text = dto.PeriodoNombre ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(dto.PeriodoInicio) &&
                DateTime.TryParseExact(dto.PeriodoInicio, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dIni))
            {
                dtpInicio.Value = dIni;
            }
            else
            {
                dtpInicio.Value = DateTime.Today;
            }

            if (!string.IsNullOrWhiteSpace(dto.PeriodoFin) &&
                DateTime.TryParseExact(dto.PeriodoFin, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dFin))
            {
                dtpFin.Value = dFin;
            }
            else
            {
                dtpFin.Value = DateTime.Today.AddMonths(5);
            }

            chkActivo.Checked = dto.Activo == 1;
            btnEliminar.Enabled = true;
        }

        private void LimpiarFormulario()
        {
            _currentPeriodo = null;
            txtCodigo.Text = string.Empty;
            dtpInicio.Value = DateTime.Today;
            dtpFin.Value = DateTime.Today.AddMonths(5);
            chkActivo.Checked = true;
            btnEliminar.Enabled = false;
            txtCodigo.Focus();
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            LimpiarFormulario();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string codigo = txtCodigo.Text.Trim();
            if (string.IsNullOrWhiteSpace(codigo))
            {
                MessageBox.Show("El código de período es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCodigo.Focus();
                return;
            }

            if (dtpFin.Value.Date < dtpInicio.Value.Date)
            {
                MessageBox.Show("La fecha de fin no puede ser anterior a la fecha de inicio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dtpFin.Focus();
                return;
            }

            var dto = new PeriodoGridDto
            {
                IdPeriodo = _currentPeriodo?.IdPeriodo ?? 0,
                PeriodoNombre = codigo,
                PeriodoInicio = dtpInicio.Value.ToString("yyyy-MM-dd"),
                PeriodoFin = dtpFin.Value.ToString("yyyy-MM-dd"),
                Activo = chkActivo.Checked ? 1 : 0
            };

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var service = new PeriodoCrudService(repo);
                var result = service.Guardar(dto);

                if (result.IsSuccess)
                {
                    MessageBox.Show("Período guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CargarPeriodos();
                }
                else
                {
                    MessageBox.Show(result.Error, "Error al guardar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormPeriodos] Error al guardar período", ex);
                MessageBox.Show($"Error al guardar período: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentPeriodo == null || _currentPeriodo.IdPeriodo <= 0) return;

            var confirm = MessageBox.Show(
                $"¿Está seguro de eliminar el período '{_currentPeriodo.PeriodoNombre}'?",
                "Confirmar Eliminación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var service = new PeriodoCrudService(repo);
                var result = service.Eliminar(_currentPeriodo.IdPeriodo);

                if (result.IsSuccess)
                {
                    MessageBox.Show("Período eliminado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CargarPeriodos();
                }
                else
                {
                    MessageBox.Show(result.Error, "No se pudo eliminar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormPeriodos] Error al eliminar período", ex);
                MessageBox.Show($"Error al eliminar período: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
