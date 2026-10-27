using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using System.Data;
using Dapper;
using System.Reflection;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormCdp : Form
    {
        private readonly Func<IDbConnection> _dbFactory;
        private readonly IDbConnection _currentConnection;
        private readonly CdpDependenciaRepository _repo;
        private BindingList<CdpDependencia> _cdpList;

        private class ActivoOption
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }

        private class DependenciaOption
        {
            public string Text { get; set; }
            public string Value { get; set; }
        }

        private class PeriodoOption
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }

        /// <summary>
        /// Obtiene el valor del atributo Description de un valor de Enum.
        /// Retorna el ToString() si el atributo no está presente.
        /// Implementado en la Fase 6 para lectura resiliente.
        /// </summary>
        private string GetDescription<T>(T value) where T : Enum
        {
            var fieldInfo = value.GetType().GetField(value.ToString());
            if (fieldInfo != null)
            {
                var attrs = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                {
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }
            return value.ToString();
        }

        /// <summary>
        /// Constructor que inicializa el formulario y sus dependencias.
        /// </summary>
        /// <param name="dbFactory">Fábrica para obtener la conexión a la base de datos.</param>
        public FormCdp(Func<IDbConnection> dbFactory)
        {
            InitializeComponent();
            _dbFactory = dbFactory ?? throw new ArgumentNullException(nameof(dbFactory));
            
            // Fase 5: Evita NullReferenceException si _dbFactory() retorna null
            _currentConnection = _dbFactory();
            if (_currentConnection == null)
            {
                throw new InvalidOperationException("La fábrica de conexiones (dbFactory) devolvió una conexión nula.");
            }

            _repo = new CdpDependenciaRepository(_currentConnection);

            cmbActivo.Items.Add(new ActivoOption { Text = "Inactivo", Value = 0 });
            cmbActivo.Items.Add(new ActivoOption { Text = "Activo", Value = 1 });
            cmbActivo.DisplayMember = "Text";
            cmbActivo.ValueMember = "Value";
            cmbActivo.SelectedIndex = 1;

            foreach (Dependencias dep in Enum.GetValues(typeof(Dependencias)))
            {
                var text = GetDescription(dep);
                cmbDependencia.Items.Add(new DependenciaOption { Text = text, Value = text });
            }
            cmbDependencia.DisplayMember = "Text";
            cmbDependencia.ValueMember = "Value";
            if (cmbDependencia.Items.Count > 0) cmbDependencia.SelectedIndex = 0;
        }

        private void FormCdp_Load(object sender, EventArgs e)
        {
            SetupGrid();
            CargarPeriodos();
            CargarCdps();
        }

        /// <summary>
        /// Carga los periodos desde la base de datos al ComboBox.
        /// Se corrigió el nombre de la columna a 'idperiodo' y 'periodo' para coincidir con la base de datos real.
        /// Tampoco utiliza 'using' para no desechar la conexión compartida '_currentConnection'.
        /// </summary>
        private void CargarPeriodos()
        {
            try
            {
                var periodos = _currentConnection.Query<PeriodoOption>("SELECT idperiodo AS Value, periodo AS Text FROM periodos ORDER BY idperiodo DESC").ToList();

                cmbPeriodos.DataSource = new System.Collections.Generic.List<PeriodoOption>(periodos);
                cmbPeriodos.DisplayMember = "Text";
                cmbPeriodos.ValueMember = "Value";

                var filtroPeriodos = new System.Collections.Generic.List<PeriodoOption> { new PeriodoOption { Text = "Todos", Value = 0 } };
                filtroPeriodos.AddRange(periodos);
                cmbPeriodosFiltro.DataSource = filtroPeriodos;
                cmbPeriodosFiltro.DisplayMember = "Text";
                cmbPeriodosFiltro.ValueMember = "Value";
                cmbPeriodosFiltro.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar periodos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el DataGridView.
        /// Fase 7 / Instrucción Principal: Se limita la vista exclusivamente a CDP Num y Dependencia.
        /// </summary>
        private void SetupGrid()
        {
            dgvCdp.AutoGenerateColumns = false;
            dgvCdp.Columns.Clear();
            dgvCdp.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "CdpNum", HeaderText = "CDP Num", Width = 120 });
            dgvCdp.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Dependencia", HeaderText = "Dependencia", Width = 120 });
            dgvCdp.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DependenciaFull", HeaderText = "Dependencia Completa", Width = 200 });
        }

        private void CargarCdps()
        {
            var result = _repo.GetAll();
            if (result.IsSuccess)
            {
                var list = result.Value.ToList();
                if (cmbPeriodosFiltro.SelectedItem is PeriodoOption opt && opt.Value > 0)
                {
                    list = list.Where(x => x.IdPeriodo == opt.Value).ToList();
                }

                _cdpList = new BindingList<CdpDependencia>(list);
                dgvCdp.DataSource = _cdpList;
            }
            else
            {
                MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvCdp_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvCdp.SelectedRows.Count == 0)
            {
                btnGuardar.Text = "Agregar";
                return;
            }

            var selected = (CdpDependencia)dgvCdp.SelectedRows[0].DataBoundItem;
            if (cmbPeriodos.DataSource is System.Collections.Generic.List<PeriodoOption> periodos)
            {
                if (periodos.Any(p => p.Value == selected.IdPeriodo))
                {
                    cmbPeriodos.SelectedValue = selected.IdPeriodo;
                }
            }
            txtCdpNum.Text = selected.CdpNum;
            txtFechaCdp.Text = selected.FechaCdp;
            txtRubro.Text = selected.Rubro;
            txtResponsable.Text = selected.Responsable;
            txtCargo.Text = selected.Cargo;
            txtElaboro.Text = selected.Elaboro;
            txtReviso.Text = selected.Reviso;
            txtDependenciaFull.Text = selected.DependenciaFull ?? "";

            int depIdx = -1;
            for (int i = 0; i < cmbDependencia.Items.Count; i++)
            {
                if (cmbDependencia.Items[i] is DependenciaOption depOpt && depOpt.Value == selected.Dependencia)
                {
                    depIdx = i;
                    break;
                }
            }
            if (depIdx >= 0) cmbDependencia.SelectedIndex = depIdx;

            int idx = -1;
            for (int i = 0; i < cmbActivo.Items.Count; i++)
            {
                if (cmbActivo.Items[i] is ActivoOption opt && opt.Value == selected.Activo)
                {
                    idx = i;
                    break;
                }
            }
            if (idx >= 0) cmbActivo.SelectedIndex = idx;
            
            btnGuardar.Text = "Actualizar";
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            btnLimpiar_Click(null, null);
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            if (cmbPeriodos.Items.Count > 0) cmbPeriodos.SelectedIndex = 0;
            txtCdpNum.Text = "";
            txtFechaCdp.Text = "";
            txtRubro.Text = "";
            txtResponsable.Text = "";
            txtCargo.Text = "";
            txtElaboro.Text = "";
            txtReviso.Text = "";
            txtDependenciaFull.Text = "";
            if (cmbDependencia.Items.Count > 0) cmbDependencia.SelectedIndex = 0;
            if (cmbActivo.Items.Count > 1) cmbActivo.SelectedIndex = 1;

            btnGuardar.Text = "Agregar";
            dgvCdp.ClearSelection();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtCdpNum.Text))
                {
                    MessageBox.Show("El número de CDP es requerido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedDep = cmbDependencia.SelectedItem as DependenciaOption;
                if (selectedDep == null)
                {
                    MessageBox.Show("Debe seleccionar una Dependencia válida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var cdp = new CdpDependencia
                {
                    IdPeriodo = cmbPeriodos.SelectedValue is int pVal ? pVal : 0,
                    CdpNum = txtCdpNum.Text.Trim(),
                    FechaCdp = txtFechaCdp.Text.Trim(),
                    Rubro = txtRubro.Text.Trim(),
                    Responsable = txtResponsable.Text.Trim(),
                    Cargo = txtCargo.Text.Trim(),
                    Elaboro = txtElaboro.Text.Trim(),
                    Reviso = txtReviso.Text.Trim(),
                    Dependencia = selectedDep.Value,
                    DependenciaFull = string.IsNullOrWhiteSpace(txtDependenciaFull.Text) ? null : txtDependenciaFull.Text.Trim(),
                    Activo = cmbActivo.SelectedItem is ActivoOption opt ? opt.Value : 1,
                    Subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad
                };

                if (dgvCdp.SelectedRows.Count > 0 && btnGuardar.Text == "Actualizar")
                {
                    var selected = (CdpDependencia)dgvCdp.SelectedRows[0].DataBoundItem;
                    cdp.Id = selected.Id;

                    var result = _repo.Update(cdp);
                    if (result.IsSuccess)
                        MessageBox.Show("CDP actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    var result = _repo.Add(cdp);
                    if (result.IsSuccess)
                        MessageBox.Show("CDP agregado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                btnLimpiar_Click(null, null);
                CargarCdps();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar CDP: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dgvCdp.SelectedRows.Count == 0) return;

            var selected = (CdpDependencia)dgvCdp.SelectedRows[0].DataBoundItem;
            if (MessageBox.Show($"¿Desea eliminar el CDP {selected.CdpNum}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var result = _repo.Delete(selected.Id);
                if (result.IsSuccess)
                {
                    MessageBox.Show("CDP eliminado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnLimpiar_Click(null, null);
                    CargarCdps();
                }
                else
                {
                    MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void cmbPeriodosFiltro_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarCdps();
        }
    }
}
