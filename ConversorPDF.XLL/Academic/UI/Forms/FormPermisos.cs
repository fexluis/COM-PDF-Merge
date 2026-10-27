using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormPermisos : Form
    {
        private readonly IAcademicRepository _repository;
        private BindingSource _bsPerfiles;

        public FormPermisos()
        {
            InitializeComponent();
            
            // Initialization via pure SQLiteHelper context since this form operates locally
            // We resolve the connection from SqliteHelper as DatabaseBootstrapper does not expose it
            var context = SqliteHelper.GetContext();
            var conn = context.Open();
            _repository = new AcademicRepository(conn);
            _bsPerfiles = new BindingSource();
        }

        private void FormPermisos_Load(object sender, EventArgs e)
        {
            try
            {
                // Cargar enumeración de botones en el CheckedListBox
                var allButtons = Enum.GetValues(typeof(RibbonButtonId)).Cast<RibbonButtonId>().ToList();
                foreach (var btn in allButtons)
                {
                    clbPermisos.Items.Add(btn, false);
                }

                LoadPerfiles();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al cargar FormPermisos", ex);
                MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadPerfiles()
        {
            var result = _repository.GetTodosLosPermisos();
            if (result.IsSuccess)
            {
                _bsPerfiles.DataSource = result.Value.ToList();
                dgvPerfiles.DataSource = _bsPerfiles;
                
                dgvPerfiles.Columns["Botones"].Visible = false; // Hide the raw string column
                
                if (dgvPerfiles.Rows.Count > 0)
                {
                    dgvPerfiles.Rows[0].Selected = true;
                    LoadPermisosForPerfil((Permiso)dgvPerfiles.Rows[0].DataBoundItem);
                }
            }
            else
            {
                MessageBox.Show($"Error al cargar perfiles: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadPermisosForPerfil(Permiso permiso)
        {
            if (permiso == null) return;
            txtPerfil.Text = permiso.Perfil;
            txtPerfil.ReadOnly = true; // No se puede cambiar el nombre a uno existente

            // Limpiar checkmarks
            for (int i = 0; i < clbPermisos.Items.Count; i++)
            {
                clbPermisos.SetItemChecked(i, false);
            }

            if (string.IsNullOrEmpty(permiso.Botones)) return;

            var botonesPermitidos = permiso.Botones.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                   .Select(b => b.Trim())
                                                   .ToList();

            for (int i = 0; i < clbPermisos.Items.Count; i++)
            {
                var itemStr = clbPermisos.Items[i].ToString();
                if (botonesPermitidos.Contains(itemStr))
                {
                    clbPermisos.SetItemChecked(i, true);
                }
            }
        }

        private void dgvPerfiles_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPerfiles.CurrentRow != null)
            {
                var permiso = (Permiso)dgvPerfiles.CurrentRow.DataBoundItem;
                LoadPermisosForPerfil(permiso);
            }
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            txtPerfil.ReadOnly = false;
            txtPerfil.Text = string.Empty;
            txtPerfil.Focus();
            for (int i = 0; i < clbPermisos.Items.Count; i++)
            {
                clbPermisos.SetItemChecked(i, false);
            }
            dgvPerfiles.ClearSelection();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPerfil.Text))
            {
                MessageBox.Show("Debe ingresar el nombre del Perfil.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 16: Validar duplicado solo si es perfil nuevo (txtPerfil editable)
            if (!txtPerfil.ReadOnly)
            {
                // Fase 20: Bloquear nombre reservado Admin
                if (txtPerfil.Text.Trim().Equals("Admin", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("No se puede crear un perfil con el nombre reservado 'Admin'.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var existentes = _repository.GetTodosLosPermisos();
                if (existentes.IsSuccess &&
                    existentes.Value.Any(p => p.Perfil.Equals(txtPerfil.Text.Trim(), StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show("Ya existe un perfil con ese nombre.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            var selectedButtons = new List<string>();
            foreach (var item in clbPermisos.CheckedItems)
            {
                selectedButtons.Add(item.ToString());
            }

            var permiso = new Permiso
            {
                Perfil = txtPerfil.Text.Trim(),
                Botones = string.Join(",", selectedButtons)
            };

            var result = _repository.SavePermiso(permiso);
            if (result.IsSuccess)
            {
                MessageBox.Show("Perfil guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadPerfiles();
                ConversorPDF.XLL.RibbonController.InvalidateRibbon();
            }
            else
            {
                MessageBox.Show($"Error al guardar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dgvPerfiles.CurrentRow != null)
            {
                // Fase 161 y Fase 162: Guarda de nulo única tras el cast para prevenir NRE en todo el método
                var permiso = dgvPerfiles.CurrentRow.DataBoundItem as Permiso;
                if (permiso == null) return;

                if (string.Equals(permiso.Perfil, "Admin", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("No se puede eliminar el perfil Administrador base.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 151 y Fase 155: Validar que no existan usuarios asignados al perfil antes de eliminarlo (fail-closed)
                var usuariosEnUso = _repository.GetUsuarios();
                if (!usuariosEnUso.IsSuccess)
                {
                    MessageBox.Show($"No se pudo verificar el uso del perfil: {usuariosEnUso.Error}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (usuariosEnUso.Value.Any(u => string.Equals(u.Perfil, permiso.Perfil, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show($"No se puede eliminar el perfil '{permiso.Perfil}': hay usuarios asignados a este perfil.",
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (MessageBox.Show($"¿Desea eliminar el perfil '{permiso.Perfil}'?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    var result = _repository.DeletePermiso(permiso.Perfil);
                    if (result.IsSuccess)
                    {
                        LoadPerfiles();
                        ConversorPDF.XLL.RibbonController.InvalidateRibbon();
                    }
                    else
                    {
                        MessageBox.Show($"Error al eliminar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
