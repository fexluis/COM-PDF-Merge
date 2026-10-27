using System;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormDivipolaBuscador : Form
    {
        private readonly IAcademicRepository _repository;
        private readonly BindingSource _bs = new BindingSource();
        public Divipola SelectedDivipola { get; private set; }

        public FormDivipolaBuscador(IAcademicRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            InitializeComponent();
            SetupGrid();
            CargarDatos("");
        }

        private void SetupGrid()
        {
            dgvResultados.AutoGenerateColumns = false;
            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Departamento", HeaderText = "Departamento", Width = 150 });
            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Municipio", HeaderText = "Municipio", Width = 200 });
            dgvResultados.DataSource = _bs;
        }

        private void CargarDatos(string searchTerm)
        {
            var result = string.IsNullOrWhiteSpace(searchTerm) 
                ? _repository.GetDivipolas() 
                : _repository.SearchDivipola(searchTerm);

            if (result.IsSuccess)
            {
                _bs.DataSource = result.Value.ToList();
            }
            else
            {
                MessageBox.Show($"Error al buscar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtBusqueda_KeyUp(object sender, KeyEventArgs e)
        {
            CargarDatos(txtBusqueda.Text);
        }

        private void dgvResultados_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvResultados.Rows[e.RowIndex].DataBoundItem is Divipola divipola)
            {
                SelectedDivipola = divipola;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            if (dgvResultados.CurrentRow?.DataBoundItem is Divipola divipola)
            {
                SelectedDivipola = divipola;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Seleccione un registro.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
