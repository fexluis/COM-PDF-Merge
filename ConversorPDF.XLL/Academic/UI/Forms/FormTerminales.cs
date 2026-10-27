using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;


namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormTerminales : Form
    {
        private readonly EncryptedSQLiteContext _context;
        private readonly TerminalRepository _repo;
        private BindingList<Terminal> _terminalesList;

        /// <summary>
        /// Inicializa el formulario. 
        /// Cambio (Fase 8): Creación del contexto local que debe ser dispuesto al cerrar el form, en vez de recibir Func de IDbConnection.
        /// </summary>
        public FormTerminales()
        {
            InitializeComponent();
            _context = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext();
            _repo = new TerminalRepository(_context);
        }

        private void FormTerminales_Load(object sender, EventArgs e)
        {
            SetupGrid();
            CargarTerminales();
        }

        private void SetupGrid()
        {
            dgvTerminales.AutoGenerateColumns = false;
            dgvTerminales.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Id", HeaderText = "Id", Width = 50 });
            dgvTerminales.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "TerminalName", HeaderText = "Terminal Aérea", Width = 250 });
            dgvTerminales.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MaxValor", HeaderText = "MaxValor", Width = 100 });
        }

        private void CargarTerminales()
        {
            var result = _repo.GetAll();
            if (result.IsSuccess)
            {
                _terminalesList = new BindingList<Terminal>(result.Value.ToList());
                dgvTerminales.DataSource = _terminalesList;
            }
            else
            {
                MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtNombre.Text))
                {
                    MessageBox.Show("El nombre de la terminal es requerido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!double.TryParse(txtMaxValor.Text, out double maxValor))
                {
                    MessageBox.Show("El valor máximo no es válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var terminal = new Terminal
                {
                    TerminalName = txtNombre.Text.Trim(),
                    MaxValor = maxValor
                };

                if (dgvTerminales.SelectedRows.Count > 0 && btnGuardar.Text == "Actualizar")
                {
                    var selected = (Terminal)dgvTerminales.SelectedRows[0].DataBoundItem;
                    terminal.Id = selected.Id;
                    
                    var result = _repo.Update(terminal);
                    if (result.IsSuccess)
                    {
                        MessageBox.Show("Terminal actualizada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    var result = _repo.Add(terminal);
                    if (result.IsSuccess)
                    {
                        MessageBox.Show("Terminal agregada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                btnLimpiar_Click(null, null);
                CargarTerminales();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar terminal: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            txtNombre.Text = "";
            txtMaxValor.Text = "0";
            btnGuardar.Text = "Agregar";
            dgvTerminales.ClearSelection();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dgvTerminales.SelectedRows.Count == 0) return;

            var selected = (Terminal)dgvTerminales.SelectedRows[0].DataBoundItem;
            if (MessageBox.Show($"¿Desea eliminar la terminal {selected.TerminalName}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var result = _repo.Delete(selected.Id);
                if (result.IsSuccess)
                {
                    MessageBox.Show("Terminal eliminada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnLimpiar_Click(null, null);
                    CargarTerminales();
                }
                else
                {
                    MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void dgvTerminales_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvTerminales.SelectedRows.Count > 0)
            {
                var selected = (Terminal)dgvTerminales.SelectedRows[0].DataBoundItem;
                txtNombre.Text = selected.TerminalName;
                // Cambio Fase 23: Especificar formato "0" para evitar notación científica
                txtMaxValor.Text = selected.MaxValor.ToString("0");
                btnGuardar.Text = "Actualizar";
            }
        }

        private void txtMaxValor_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir solo números y la tecla de borrado (Backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// Cambio (Fase 8): Asegurar que el contexto se dispone (re-encriptando el Vault) al cerrar el formulario.
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _context?.Dispose();
            base.OnFormClosing(e);
        }
    }
}
