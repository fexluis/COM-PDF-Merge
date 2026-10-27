using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormGastosTerrestres : Form
    {
        private readonly IGastosTerrestresService _service;
        private readonly IDbConnection _dbConnection;
        private List<GastoTerrestreGridDto> _gastos;
        private BindingSource _bindingSource;

        public FormGastosTerrestres(Func<IDbConnection> dbFactory, IConfigRepository configRepo = null)
        {
            InitializeComponent();
            _dbConnection = dbFactory();
            
            var repo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.GastosTerrestresRepository(_dbConnection);
            var finalConfigRepo = configRepo ?? new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.ConfigRepository(SqliteHelper.GetContext());
            var encoder = new ConversorPDF.XLL.Academic.Application.Services.TrayectoEncoderService(repo, finalConfigRepo);
            _service = new ConversorPDF.XLL.Academic.Application.Services.GastosTerrestresService(repo, encoder);

            _gastos = new List<GastoTerrestreGridDto>();
            _bindingSource = new BindingSource();
            dgvGastos.DataSource = _bindingSource;
        }

        /// <summary>
        /// Libera los recursos del formulario, incluyendo la conexión a base de datos instanciada.
        /// Resolviendo conflicto de override (CS0111) moviendo esta lógica exclusivamente al code-behind.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados deben ser eliminados.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                _dbConnection?.Dispose();
            }
            base.Dispose(disposing);
        }

        private async void FormGastosTerrestres_Load(object sender, EventArgs e)
        {
            await CargarSubunidadesAsync();
            await CargarGrillaAsync();
        }

        /// <summary>
        /// Carga asincrónicamente el listado de subunidades directamente de la BD sin bloquear la interfaz.
        /// (Fase 12: Dapper asíncrono implementado).
        /// </summary>
        private async Task CargarSubunidadesAsync()
        {
            try
            {
                var subunidades = (await _dbConnection.QueryAsync<string>("SELECT DISTINCT subunidad FROM territoriales WHERE subunidad IS NOT NULL AND subunidad != '' ORDER BY subunidad")).ToList();
                cmbSubunidad.Items.Clear();
                cmbSubunidad.Items.Add("");
                cmbSubunidad.Items.AddRange(subunidades.ToArray());
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormGastosTerrestres] Error cargando subunidades", ex);
                MessageBox.Show("Error al cargar las subunidades.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task CargarGrillaAsync()
        {
            try
            {
                var result = await _service.GetAllAsync();
                if (this.IsDisposed) return;
                
                this.Invoke((MethodInvoker)delegate
                {
                    if (this.IsDisposed) return;
                    if (result.IsSuccess)
                    {
                        _gastos = new List<GastoTerrestreGridDto>(result.Value);
                        _bindingSource.DataSource = _gastos;
                        dgvGastos.Refresh();
                    }
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormGastosTerrestres] Error al cargar gastos terrestres.", ex);
                if (!this.IsDisposed)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show("Error al cargar los datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void dgvGastos_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvGastos.SelectedRows.Count == 0)
            {
                cmbSubunidad.Enabled = true;
                return;
            }

            var selectedDto = dgvGastos.SelectedRows[0].DataBoundItem as GastoTerrestreGridDto;
            if (selectedDto == null) return;
            
            cmbSubunidad.Enabled = false; // Disable to prevent changing territorial for existing trayectos

            txtCodigoTrayecto.Text = selectedDto.CodigoTrayecto ?? "";
            txtTrayecto.Text = selectedDto.Trayecto ?? "";
            
            // Fase 22: Abrazadera (clamp) para prevenir ArgumentOutOfRangeException si el valor BD excede UI
            numValor.Value = Math.Min(Math.Max(selectedDto.Valor ?? 0, numValor.Minimum), numValor.Maximum);
            
            // Fase 46: Fallback a Territorial para registros antiguos donde Subunidad es NULL
            string searchValue = !string.IsNullOrEmpty(selectedDto.Subunidad) ? selectedDto.Subunidad : selectedDto.Territorial;

            if (!string.IsNullOrEmpty(searchValue) && cmbSubunidad.Items.Contains(searchValue))
            {
                cmbSubunidad.SelectedItem = searchValue;
            }
            else
            {
                // Fase 24: Guard preventivo si items está vacío
                if (cmbSubunidad.Items.Count > 0)
                {
                    cmbSubunidad.SelectedIndex = 0;
                }
            }
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            dgvGastos.ClearSelection();
            txtCodigoTrayecto.Text = "";
            txtTrayecto.Text = "";
            numValor.Value = 0;
            
            // Fase 23: Prevenir ArgumentOutOfRange si el combo está vacío
            if (cmbSubunidad.Items.Count > 0)
            {
                cmbSubunidad.SelectedIndex = 0;
            }
            
            txtTrayecto.Focus();
        }

        private async void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                string trayecto = txtTrayecto.Text.Trim();
                string subunidadSeleccionada = cmbSubunidad.SelectedItem?.ToString() ?? "";
                decimal? valor = numValor.Value > 0 ? (decimal?)numValor.Value : null;

                if (string.IsNullOrWhiteSpace(trayecto) || string.IsNullOrWhiteSpace(subunidadSeleccionada))
                {
                    MessageBox.Show("El trayecto y la territorial son obligatorios.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (dgvGastos.SelectedRows.Count > 0)
                {
                    // Update
                    var selectedDto = dgvGastos.SelectedRows[0].DataBoundItem as GastoTerrestreGridDto;
                    if (selectedDto == null) return;
                    
                    selectedDto.Trayecto = trayecto;
                    selectedDto.Valor = valor;
                    // Territorial NO se modifica - proviene del registro original

                    var result = await _service.UpdateAsync(selectedDto);
                    if (result.IsSuccess)
                    {
                        MessageBox.Show("Registro actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    // Create new
                    var result = await _service.CrearTrayectoAsync(subunidadSeleccionada, trayecto, valor);
                    if (result.IsSuccess)
                    {
                        MessageBox.Show("Registro creado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                await CargarGrillaAsync();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormGastosTerrestres] Error al guardar.", ex);
                MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dgvGastos.SelectedRows.Count == 0) return;
            
            var resultConfirm = MessageBox.Show("¿Está seguro de eliminar este registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultConfirm == DialogResult.Yes)
            {
                var selectedDto = dgvGastos.SelectedRows[0].DataBoundItem as GastoTerrestreGridDto;
                if (selectedDto == null) return;
                
                var result = await _service.DeleteAsync(selectedDto.Id);
                
                if (result.IsSuccess)
                {
                    MessageBox.Show("Registro eliminado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    await CargarGrillaAsync();
                }
                else
                {
                    MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
