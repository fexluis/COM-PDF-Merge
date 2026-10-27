using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;
using ExcelDna.Integration;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormTerceros : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly ITerceroAsignaturaRepository _tasRepo;
        private readonly ITerceroAsignaturaSyncService _syncService;

        private BindingSource _bs = new BindingSource();
        private BindingSource _bsAsignaturas = new BindingSource();
        private Tercero _currentTercero;
        private int? _selectedResidenciaDivipolaId;
        private List<Tercero> _allTerceros = new List<Tercero>();
        private List<TerceroAsignaturaDto> _asignaturasDocente = new List<TerceroAsignaturaDto>();
        private bool _cargandoCombo = false;

        public FormTerceros(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _tasRepo = new TerceroAsignaturaRepository(connectionFactory);
            _syncService = new TerceroAsignaturaSyncService(connectionFactory, _tasRepo);

            InitializeComponent();
            SetupGrid();
            SetupAsignaturasGrid();
            CargarTerceros();

            tabControlTerceros.SelectedIndexChanged += TabControlTerceros_SelectedIndexChanged;
        }

        private void SetupGrid()
        {
            dgvTerceros.AutoGenerateColumns = false;
            dgvTerceros.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Nombre", HeaderText = "Nombre", Width = 200 });
            dgvTerceros.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Cedula", HeaderText = "Cédula", Width = 100 });
            dgvTerceros.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Tipo", HeaderText = "Tipo", Width = 120 });
            dgvTerceros.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Categoria", HeaderText = "Rol", Width = 120 });
            dgvTerceros.DataSource = _bs;
            
            dgvTerceros.SelectionChanged += DgvTerceros_SelectionChanged;
            txtBuscar.TextChanged += (s, e) => AplicarFiltroTerceros();
        }

        private void SetupAsignaturasGrid()
        {
            dgvAsignaturasHabilitadas.AutoGenerateColumns = false;
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "NombrePrograma", HeaderText = "Programa", Width = 180 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Semestre", HeaderText = "Semestre", Width = 70 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "NombreAsignatura", HeaderText = "Asignatura", Width = 220 });
            
            var colActivo = new DataGridViewCheckBoxColumn 
            { 
                DataPropertyName = "Activo", 
                HeaderText = "Activo", 
                Width = 60,
                TrueValue = 1,
                FalseValue = 0
            };
            dgvAsignaturasHabilitadas.Columns.Add(colActivo);

            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Prioridad", HeaderText = "Prioridad", Width = 70 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "VecesAsignada", HeaderText = "Veces Asignada", Width = 110 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "FechaUltimaAsignacion", HeaderText = "Última Asignación", Width = 120 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Origen", HeaderText = "Origen", Width = 90 });
            dgvAsignaturasHabilitadas.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Observaciones", HeaderText = "Observaciones", Width = 200 });

            dgvAsignaturasHabilitadas.DataSource = _bsAsignaturas;
            tabAsignaturas.Enabled = false;
        }

        private void TabControlTerceros_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlTerceros.SelectedTab == tabAsignaturas)
            {
                CargarAsignaturasDocenteActual();
            }
        }

        private void CargarTerritorialesEnCheckedListBox()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetSubunidades();
                if (result.IsSuccess)
                {
                    chkListDepartamentos.Items.Clear();
                    foreach (var sub in result.Value)
                    {
                        chkListDepartamentos.Items.Add(sub);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[FormTerceros] No se pudieron cargar territoriales: {ex.Message}");
            }
        }

        private void CargarTerceros()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                bool soloActivos = chkMostrarInactivos == null || !chkMostrarInactivos.Checked;
                
                var result = repo.GetTerceros(soloActivos);
                if (result.IsSuccess)
                {
                    _allTerceros = result.Value.ToList();
                    CargarTerritorialesEnCheckedListBox();
                    ActualizarComboTerritoriales();
                    AplicarFiltroTerceros();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar terceros: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ActualizarComboTerritoriales()
        {
            if (_cargandoCombo) return;
            _cargandoCombo = true;
            try
            {
                string selectedValue = cmbTerritorial.SelectedItem as string ?? "- Todas las Territoriales";
                
                var departamentos = _allTerceros
                    .SelectMany(t => (t.Departamento ?? "")
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Trim().ToUpper()))
                    .Where(d => !string.IsNullOrWhiteSpace(d))
                    .Distinct()
                    .OrderBy(d => d)
                    .ToList();
                    
                cmbTerritorial.Items.Clear();
                cmbTerritorial.Items.Add("- Todas las Territoriales");
                foreach (var dep in departamentos)
                {
                    cmbTerritorial.Items.Add(dep);
                }
                
                if (cmbTerritorial.Items.Contains(selectedValue))
                {
                    cmbTerritorial.SelectedItem = selectedValue;
                }
                else
                {
                    cmbTerritorial.SelectedIndex = 0;
                }
            }
            finally
            {
                _cargandoCombo = false;
            }
        }

        private void AplicarFiltroTerceros()
        {
            string selectedValue = cmbTerritorial.SelectedItem as string ?? "- Todas las Territoriales";
            string searchText = txtBuscar?.Text?.Trim().ToLowerInvariant() ?? string.Empty;
            
            var filtered = _allTerceros.AsEnumerable();
            
            if (selectedValue != "- Todas las Territoriales")
            {
                filtered = filtered
                    .Where(t => (t.Departamento ?? "")
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Trim())
                        .Any(s => string.Equals(s, selectedValue, StringComparison.OrdinalIgnoreCase)));
            }

            if (!string.IsNullOrEmpty(searchText))
            {
                filtered = filtered.Where(t => 
                    (t.Nombre != null && t.Nombre.ToLowerInvariant().Contains(searchText)) ||
                    (t.Cedula != null && t.Cedula.ToLowerInvariant().Contains(searchText)));
            }
            
            _bs.DataSource = filtered.ToList();
        }

        private void cmbTerritorial_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cargandoCombo) return;
            AplicarFiltroTerceros();
        }

        private void DgvTerceros_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvTerceros.CurrentRow?.DataBoundItem is Tercero t)
            {
                CargarDatosTercero(t);
            }
            else
            {
                LimpiarCampos();
            }
        }

        private void CargarDatosTercero(Tercero t)
        {
            _currentTercero = t;
            
            for (int i = 0; i < chkListDepartamentos.Items.Count; i++)
            {
                chkListDepartamentos.SetItemChecked(i, false);
            }

            txtNombre.Text = t.Nombre;
            txtCedula.Text = t.Cedula;
            if (t.Subunidad != null || t.Departamento != null)
            {
                var subunidades = (t.Subunidad ?? t.Departamento ?? "")
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
                
                for (int i = 0; i < chkListDepartamentos.Items.Count; i++)
                {
                    string item = chkListDepartamentos.Items[i].ToString();
                    chkListDepartamentos.SetItemChecked(i, subunidades.Any(s => string.Equals(s, item, StringComparison.OrdinalIgnoreCase)));
                }
            }
            cmbTipo.Text = t.Tipo;
            numSalario.Value = t.Salario;
            cmbCategoria.Text = t.Categoria;
            txtEntidad.Text = t.EntidadBancaria;
            cmbCuenta.Text = t.CuentaBancaria;
            cmbTipoCuenta.Text = t.TipoCuentaBancaria;
            txtTelefono.Text = t.Telefono;
            txtCorreo.Text = t.Correo;
            txtCorreoPersonal.Text = t.CorreoPersonal;
            DateTime fecha;
            if (DateTime.TryParseExact(t.FechaDeNacimiento, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha)
                || DateTime.TryParseExact(t.FechaDeNacimiento, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
            {
                dtpFechaNacimiento.Value = fecha;
            }
            else
            {
                dtpFechaNacimiento.Value = DateTime.Today;
            }
            txtContrato.Text = t.Contrato;
            txtResidencia.Text = t.Residencia;
            _selectedResidenciaDivipolaId = t.IdResidenciaDivipola;
            chkActivo.Checked = t.Activo == 1;
            chkPensionado.Checked = t.Pensionado == 1;
            numPuntaje.Value = Math.Min(t.PuntajeOcasionales, numPuntaje.Maximum);
            btnEliminar.Enabled = true;

            // Habilitación de pestaña de asignaturas
            tabAsignaturas.Enabled = _currentTercero != null && _currentTercero.Id > 0;
            lblDocenteAsignaturasTitulo.Text = _currentTercero != null 
                ? $"Docente: {_currentTercero.Nombre} (C.C. {_currentTercero.Cedula})"
                : "Docente: (ninguno seleccionado)";

            if (tabControlTerceros.SelectedTab == tabAsignaturas)
            {
                CargarAsignaturasDocenteActual();
            }
        }

        private async void CargarAsignaturasDocenteActual()
        {
            if (_currentTercero == null || _currentTercero.Id <= 0)
            {
                _asignaturasDocente = new List<TerceroAsignaturaDto>();
                _bsAsignaturas.DataSource = _asignaturasDocente;
                return;
            }

            try
            {
                var result = await _tasRepo.GetByTerceroAsync(_currentTercero.Id, CancellationToken.None);
                if (result.IsFailure)
                {
                    MessageBox.Show($"Error al cargar asignaturas habilitadas: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Dictionary<int, string> programasDict;
                Dictionary<int, string> asignaturasDict;
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    programasDict = db.Query<Programa>("SELECT id as Id, nombreprograma as NombrePrograma FROM programas").ToDictionary(p => p.Id, p => p.NombrePrograma);
                    asignaturasDict = db.Query<Asignatura>("SELECT id as Id, nombre as Nombre FROM asignaturas").ToDictionary(a => a.Id, a => a.Nombre);
                }

                _asignaturasDocente = result.Value.Select(entity =>
                {
                    programasDict.TryGetValue(entity.IdPrograma, out string nomProg);
                    asignaturasDict.TryGetValue(entity.IdAsignatura, out string nomAsig);
                    return TerceroAsignaturaMapper.ToDto(entity, _currentTercero.Nombre, nomProg ?? $"Prog {entity.IdPrograma}", nomAsig ?? $"Asig {entity.IdAsignatura}");
                }).OrderBy(d => d.NombrePrograma).ThenBy(d => d.Semestre).ThenBy(d => d.NombreAsignatura).ToList();

                _bsAsignaturas.DataSource = _asignaturasDocente;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormTerceros] Error al cargar asignaturas del docente", ex);
                MessageBox.Show($"Error al cargar asignaturas del docente: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LimpiarCampos()
        {
            _currentTercero = null;
            _selectedResidenciaDivipolaId = null;
            txtNombre.Clear();
            txtCedula.Clear();
            for (int i = 0; i < chkListDepartamentos.Items.Count; i++)
            {
                chkListDepartamentos.SetItemChecked(i, false);
            }
            cmbTipo.SelectedIndex = -1;
            numSalario.Value = 0;
            cmbCategoria.SelectedIndex = -1;
            txtEntidad.Clear();
            cmbCuenta.SelectedIndex = -1;
            cmbTipoCuenta.SelectedIndex = -1;
            txtTelefono.Clear();
            txtCorreo.Clear();
            txtCorreoPersonal.Clear();
            dtpFechaNacimiento.Value = DateTime.Today;
            txtContrato.Clear();
            txtResidencia.Clear();
            chkActivo.Checked = true;
            chkPensionado.Checked = false;
            numPuntaje.Value = 0;
            btnEliminar.Enabled = false;

            tabAsignaturas.Enabled = false;
            lblDocenteAsignaturasTitulo.Text = "Docente: (ninguno seleccionado)";
            _asignaturasDocente = new List<TerceroAsignaturaDto>();
            _bsAsignaturas.DataSource = _asignaturasDocente;
        }

        private void BtnAgregarAsignatura_Click(object sender, EventArgs e)
        {
            if (_currentTercero == null || _currentTercero.Id <= 0) return;

            using (var form = new FormTerceroAsignaturas(_currentTercero.Id, null, _connectionFactory))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    CargarAsignaturasDocenteActual();
                }
            }
        }

        private async void BtnEditarAsignatura_Click(object sender, EventArgs e)
        {
            if (_currentTercero == null || _currentTercero.Id <= 0) return;
            if (dgvAsignaturasHabilitadas.CurrentRow?.DataBoundItem is TerceroAsignaturaDto dto)
            {
                var result = await _tasRepo.GetByIdAsync(dto.Id, CancellationToken.None);
                if (result.IsFailure || result.Value == null)
                {
                    MessageBox.Show("No se pudo cargar la habilitación seleccionada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (var form = new FormTerceroAsignaturas(_currentTercero.Id, result.Value, _connectionFactory))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        CargarAsignaturasDocenteActual();
                    }
                }
            }
            else
            {
                MessageBox.Show("Seleccione una asignatura habilitada para editar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void BtnEliminarAsignatura_Click(object sender, EventArgs e)
        {
            if (_currentTercero == null || _currentTercero.Id <= 0) return;
            if (dgvAsignaturasHabilitadas.CurrentRow?.DataBoundItem is TerceroAsignaturaDto dto)
            {
                if (MessageBox.Show($"¿Desea deshabilitar la asignatura '{dto.NombreAsignatura}' para el docente '{_currentTercero.Nombre}'?", "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    var result = await _tasRepo.DeleteAsync(dto.Id, CancellationToken.None);
                    if (result.IsFailure)
                    {
                        MessageBox.Show($"Error al eliminar habilitación: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    CargarAsignaturasDocenteActual();
                }
            }
            else
            {
                MessageBox.Show("Seleccione una asignatura habilitada para eliminar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void BtnImportarHistorico_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea sincronizar y derivar automáticamente el pool de habilitaciones a partir del histórico de vinculaciones?", "Sincronizar Histórico", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            string subunidadActual = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
            var result = await _syncService.SincronizarTodoAsync(subunidadActual, CancellationToken.None);
            if (result.IsFailure)
            {
                MessageBox.Show($"Error en sincronización: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show($"Sincronización completada. Se procesaron {result.Value} registros del histórico.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            CargarAsignaturasDocenteActual();
        }

        private void BtnRefrescarAsignaturas_Click(object sender, EventArgs e)
        {
            CargarAsignaturasDocenteActual();
        }

        private string SolicitarCedula(string mensaje)
        {
            using (Form prompt = new Form())
            {
                prompt.Width = 350;
                prompt.Height = 150;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.Text = "Cédula";
                prompt.StartPosition = FormStartPosition.CenterParent;
                prompt.MaximizeBox = false;
                prompt.MinimizeBox = false;

                Label textLabel = new Label() { Left = 20, Top = 20, Text = mensaje, Width = 300 };
                TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 300 };
                Button confirmation = new Button() { Text = "Aceptar", Left = 220, Width = 100, Top = 80, DialogResult = DialogResult.OK };

                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog(this) == DialogResult.OK ? textBox.Text.Trim() : null;
            }
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            string cedula = SolicitarCedula("Ingrese la cédula para el nuevo tercero:");
            if (string.IsNullOrWhiteSpace(cedula)) return;

            var db = _connectionFactory();
            var repo = new AcademicRepository(db);
            var result = repo.GetTerceroByCedula(cedula);

            if (result.IsSuccess && result.Value != null)
            {
                MessageBox.Show("El tercero ya existe. Se cargarán sus datos para edición.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarDatosTercero(result.Value);
            }
            else
            {
                LimpiarCampos();
                txtCedula.Text = cedula;
                txtNombre.Focus();
            }
        }

        private void btnEditCedula_Click(object sender, EventArgs e)
        {
            string nuevaCedula = SolicitarCedula("Ingrese la nueva cédula:");
            if (string.IsNullOrWhiteSpace(nuevaCedula) || nuevaCedula == txtCedula.Text) return;

            var db = _connectionFactory();
            var repo = new AcademicRepository(db);
            var result = repo.GetTerceroByCedula(nuevaCedula);

            if (result.IsSuccess && result.Value != null)
            {
                MessageBox.Show("Ya existe un tercero con esa cédula.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                txtCedula.Text = nuevaCedula;
            }
        }

        private void btnSelectCiudad_Click(object sender, EventArgs e)
        {
            var db = _connectionFactory();
            if (db == null)
            {
                MessageBox.Show("Error de conexión: no se pudo conectar a la base de datos.", 
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            var repo = new AcademicRepository(db);
            using (var form = new FormDivipolaBuscador(repo))
            {
                if (form.ShowDialog(this) == DialogResult.OK && form.SelectedDivipola != null)
                {
                    txtResidencia.Text = form.SelectedDivipola.Municipio;
                    _selectedResidenciaDivipolaId = form.SelectedDivipola.Id;
                }
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNombre.Text) || string.IsNullOrWhiteSpace(txtCedula.Text))
            {
                MessageBox.Show("Nombre y Cédula son obligatorios.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var t = _currentTercero ?? new Tercero();
            t.Nombre = txtNombre.Text;
            t.Cedula = txtCedula.Text;
            var subunidadesOriginales = (t.Subunidad ?? t.Departamento ?? "")
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
            
            string currentSubunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
            subunidadesOriginales.RemoveAll(s => string.Equals(s, currentSubunidad, StringComparison.OrdinalIgnoreCase));
            
            var subunidadesSeleccionadasUI = new List<string>();
            for (int i = 0; i < chkListDepartamentos.Items.Count; i++)
            {
                if (chkListDepartamentos.GetItemChecked(i))
                {
                    subunidadesSeleccionadasUI.Add(chkListDepartamentos.Items[i].ToString());
                }
            }
            
            string subunidadesJoined = string.Join(",",
                subunidadesOriginales.Concat(subunidadesSeleccionadasUI).Distinct(StringComparer.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(subunidadesJoined))
            {
                MessageBox.Show("Debe seleccionar al menos una subunidad.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            t.Departamento = subunidadesJoined;
            t.Subunidad = subunidadesJoined;
            t.Tipo = cmbTipo.Text;
            t.Salario = numSalario.Value;
            t.Categoria = cmbCategoria.Text;
            t.EntidadBancaria = txtEntidad.Text;
            t.CuentaBancaria = cmbCuenta.Text;
            t.TipoCuentaBancaria = cmbTipoCuenta.Text;
            t.Telefono = txtTelefono.Text;
            t.Correo = txtCorreo.Text;
            t.CorreoPersonal = txtCorreoPersonal.Text;
            t.FechaDeNacimiento = dtpFechaNacimiento.Value.ToString("dd/MM/yyyy");
            t.Contrato = txtContrato.Text;
            t.Residencia = txtResidencia.Text;
            t.IdResidenciaDivipola = _selectedResidenciaDivipolaId;
            t.Activo = chkActivo.Checked ? 1 : 0;
            t.Pensionado = chkPensionado.Checked ? 1 : 0;
            t.PuntajeOcasionales = numPuntaje.Value;

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = _currentTercero == null || _currentTercero.Id == 0 ? repo.InsertTercero(t) : repo.UpdateTercero(t);
                if (result.IsFailure) { MessageBox.Show($"Error al guardar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                if (_currentTercero == null || _currentTercero.Id == 0) t.Id = result.Value;
                
                CatalogCacheService.Invalidate();
                
                MessageBox.Show("Tercero guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarTerceros();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentTercero == null) return;

            if (MessageBox.Show($"¿Remover {_currentTercero.Nombre} de la subunidad actual?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var db = _connectionFactory();
                    var repo = new AcademicRepository(db);
                    var result = repo.DeleteTercero(_currentTercero.Id);
                    
                    if (result.IsFailure) 
                    { 
                        MessageBox.Show($"Error al remover: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                        return; 
                    }
                    
                    CatalogCacheService.Invalidate();
                    MessageBox.Show("Tercero removido correctamente de la subunidad.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CargarTerceros();
                    LimpiarCampos();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al remover: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkMostrarInactivos_CheckedChanged(object sender, EventArgs e)
        {
            CargarTerceros();
        }

        private void btnExportarBase_Click(object sender, EventArgs e)
        {
            ExportarAExcel("Base docente");
        }

        private void btnExportarTodos_Click(object sender, EventArgs e)
        {
            ExportarAExcel("Datos");
        }

        private void ExportarAExcel(string nombreHoja)
        {
            string territorial = cmbTerritorial.Text;
            string mensajeConfirmacion = territorial == "- Todas las Territoriales"
                ? "¿Desea exportar todas las territoriales?"
                : $"¿Desea exportar {territorial}?";

            if (MessageBox.Show(mensajeConfirmacion, "Confirmar Exportación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            dynamic excelApp = null;
            dynamic wb = null;
            dynamic ws = null;
            try
            {
                excelApp = ExcelDnaUtil.Application;
                if (excelApp == null)
                {
                    MessageBox.Show("No se pudo obtener la instancia activa de Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                wb = excelApp.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("No hay un libro de Excel activo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool existeHoja = false;
                int sheetCount = wb.Sheets.Count;
                for (int i = 1; i <= sheetCount; i++)
                {
                    dynamic tempSheet = wb.Sheets[i];
                    if (string.Equals(tempSheet.Name, nombreHoja, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = tempSheet;
                        existeHoja = true;
                        break;
                    }
                    Marshal.ReleaseComObject(tempSheet);
                }

                if (!existeHoja)
                {
                    ws = wb.Sheets.Add(Type.Missing, wb.Sheets[wb.Sheets.Count]);
                    ws.Name = nombreHoja;
                }

                ws.Cells.Clear();

                var terceros = new List<Tercero>();
                for (int i = 0; i < dgvTerceros.Rows.Count; i++)
                {
                    if (dgvTerceros.Rows[i].DataBoundItem is Tercero t)
                    {
                        terceros.Add(t);
                    }
                }

                using (ExcelOptimizeHelper.Optimize(excelApp))
                {
                    if (nombreHoja == "Base docente")
                    {
                        string[] headers = { "IDENTIFICACIÓN", "NOMBRE", "CATEGORIA" };
                        for (int c = 0; c < headers.Length; c++)
                        {
                            dynamic cell = ws.Cells[1, c + 1];
                            cell.Value = headers[c];
                            cell.Font.Bold = true;
                            Marshal.ReleaseComObject(cell);
                        }

                        for (int r = 0; r < terceros.Count; r++)
                        {
                            var t = terceros[r];
                            EscribirCelda(ws, r + 2, 1, t.Cedula);
                            EscribirCelda(ws, r + 2, 2, t.Nombre);
                            EscribirCelda(ws, r + 2, 3, t.Categoria);
                        }

                        dynamic range = ws.Range["A1:C1"].EntireColumn;
                        range.AutoFit();
                        Marshal.ReleaseComObject(range);
                    }
                    else if (nombreHoja == "Datos")
                    {
                        string[] headers = {
                            "NOMBRE", "CEDULA", "TIPO", "SALARIO", "CATEGORIA", 
                            "ENTIDAD BANCARIA", "CUENTA BANCARIA", "AHORRO / CORRIENTE", 
                            "TELEFONO", "CORREO INSTITUCIONAL", "CORREO PERSONAL", 
                            "FECHA DE NACIMIENTO", "CONTRATO", "RESIDENCIA", "ID DIVIPOLA RESIDENCIA", "PENSIONADO"
                        };
                        for (int c = 0; c < headers.Length; c++)
                        {
                            dynamic cell = ws.Cells[1, c + 1];
                            cell.Value = headers[c];
                            cell.Font.Bold = true;
                            Marshal.ReleaseComObject(cell);
                        }

                        for (int r = 0; r < terceros.Count; r++)
                        {
                            var t = terceros[r];
                            string pensionadoText = (t.Pensionado == 1) ? "PENSIONADO" : "NO PENSIONADO";

                            EscribirCelda(ws, r + 2, 1, t.Nombre);
                            EscribirCelda(ws, r + 2, 2, t.Cedula);
                            EscribirCelda(ws, r + 2, 3, t.Tipo);
                            EscribirCelda(ws, r + 2, 4, t.Salario);
                            EscribirCelda(ws, r + 2, 5, t.Categoria);
                            EscribirCelda(ws, r + 2, 6, t.EntidadBancaria);
                            EscribirCelda(ws, r + 2, 7, t.CuentaBancaria);
                            EscribirCelda(ws, r + 2, 8, t.TipoCuentaBancaria);
                            EscribirCelda(ws, r + 2, 9, t.Telefono);
                            EscribirCelda(ws, r + 2, 10, t.Correo);
                            EscribirCelda(ws, r + 2, 11, t.CorreoPersonal);
                            EscribirCelda(ws, r + 2, 12, t.FechaDeNacimiento);
                            EscribirCelda(ws, r + 2, 13, t.Contrato);
                            EscribirCelda(ws, r + 2, 14, t.Residencia);
                            EscribirCelda(ws, r + 2, 15, t.IdResidenciaDivipola?.ToString() ?? "");
                            EscribirCelda(ws, r + 2, 16, pensionadoText);
                        }

                        dynamic range = ws.Range["A1:P1"].EntireColumn;
                        range.AutoFit();
                        Marshal.ReleaseComObject(range);
                    }
                }

                MessageBox.Show($"Exportación a la hoja '{nombreHoja}' completada con éxito ({terceros.Count} registros).", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar a Excel: {ex.Message}", "Error de Exportación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        }

        private void EscribirCelda(dynamic ws, int fila, int col, object valor)
        {
            dynamic cell = null;
            try
            {
                cell = ws.Cells[fila, col];
                cell.Value = valor;
            }
            finally
            {
                if (cell != null) Marshal.ReleaseComObject(cell);
            }
        }

        private void chkListDepartamentos_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string currentSubunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
            string item = chkListDepartamentos.Items[e.Index].ToString();

            if (!string.Equals(item, currentSubunidad, StringComparison.OrdinalIgnoreCase))
            {
                e.NewValue = e.CurrentValue;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            foreach (Form form in System.Windows.Forms.Application.OpenForms)
            {
                if (form is FormVinculaciones fv)
                {
                    fv.RecargarCatalogosUI();
                }
            }
        }
    }
}
