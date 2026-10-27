using System;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormDivipolaCRUD : Form
    {
        private readonly IAcademicRepository _repository;
        private BindingSource _bs = new BindingSource();
        private Divipola _currentDivipola;

        public FormDivipolaCRUD(IAcademicRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            InitializeComponent();
            SetupGrid();
            CargarDatos();
        }

        private void SetupGrid()
        {
            dgvDivipola.AutoGenerateColumns = false;
            dgvDivipola.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Departamento", HeaderText = "Departamento", Width = 120 });
            dgvDivipola.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Municipio", HeaderText = "Municipio", Width = 150 });
            dgvDivipola.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MunicipioCategoria", HeaderText = "Categoría", Width = 80 });
            dgvDivipola.DataSource = _bs;

            dgvDivipola.SelectionChanged += DgvDivipola_SelectionChanged;
        }

        private void CargarDatos()
        {
            var result = _repository.GetDivipolas();
            if (result.IsSuccess)
            {
                _bs.DataSource = result.Value.ToList();
            }
            else
            {
                MessageBox.Show($"Error al cargar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvDivipola_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvDivipola.CurrentRow?.DataBoundItem is Divipola d)
            {
                _currentDivipola = d;
                txtDepartamento.Text = d.Departamento;
                txtDepCodigo.Text = d.DepartamentoCodigo;
                txtDepNorm.Text = d.DepartamentoNormalizado;
                txtMunicipio.Text = d.Municipio;
                txtMunCodigo.Text = d.MunicipioCodigo;
                txtMunNorm.Text = d.MunicipioNormalizado;
                numCategoria.Value = d.MunicipioCategoria;
                btnEliminar.Enabled = true;
            }
            else
            {
                LimpiarCampos();
            }
        }

        private void LimpiarCampos()
        {
            _currentDivipola = null;
            txtDepartamento.Clear();
            txtDepCodigo.Clear();
            txtDepNorm.Clear();
            txtMunicipio.Clear();
            txtMunCodigo.Clear();
            txtMunNorm.Clear();
            numCategoria.Value = 2; // Default
            btnEliminar.Enabled = false;
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
            txtDepartamento.Focus();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDepartamento.Text) || string.IsNullOrWhiteSpace(txtMunicipio.Text))
            {
                MessageBox.Show("Departamento y Municipio son obligatorios.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var d = _currentDivipola ?? new Divipola();
            d.Departamento = txtDepartamento.Text;
            d.DepartamentoCodigo = txtDepCodigo.Text;
            d.DepartamentoNormalizado = txtDepNorm.Text;
            d.Municipio = txtMunicipio.Text;
            d.MunicipioCodigo = txtMunCodigo.Text;
            d.MunicipioNormalizado = txtMunNorm.Text;
            d.MunicipioCategoria = (int)numCategoria.Value;

            var result = d.Id == 0 ? _repository.InsertDivipola(d) : _repository.UpdateDivipola(d);

            if (result.IsSuccess)
            {
                MessageBox.Show("Registro guardado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarDatos();
            }
            else
            {
                MessageBox.Show($"Error al guardar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentDivipola == null) return;

            if (MessageBox.Show($"¿Eliminar {_currentDivipola.Municipio}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var result = _repository.DeleteDivipola(_currentDivipola.Id);
                if (result.IsSuccess)
                {
                    MessageBox.Show("Eliminado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CargarDatos();
                    LimpiarCampos();
                }
                else
                {
                    MessageBox.Show($"Error: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
