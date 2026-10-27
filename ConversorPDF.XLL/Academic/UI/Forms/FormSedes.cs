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
    public partial class FormSedes : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private BindingSource _bs = new BindingSource();
        private Sede _currentSede;
        private bool _cargandoSubunidades = false;
        private readonly decimal _numTransporteMaxOriginal = 1000000m;
        // Fase 18: Prevención de doble clic en botón de guardado
        private bool _guardando = false;
        // Fase 23: Prevención de reentrancia en renderizado de Grid
        private bool _cargandoSedes = false;
        private bool _mapeandoDivipola = false;

        private class ActivoOption
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }

        public FormSedes(Func<IDbConnection> connectionFactory)
        {
            // Fase 3: Guarda para prevenir NullReferenceException si la factoría inyectada es null
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            InitializeComponent();
            
            // Poblar cmbActivo
            cmbActivo.Items.Add(new ActivoOption { Text = "Inactivo", Value = 0 });
            cmbActivo.Items.Add(new ActivoOption { Text = "Activo", Value = 1 });
            cmbActivo.DisplayMember = "Text";
            cmbActivo.ValueMember = "Value";
            cmbActivo.SelectedIndex = 1;

            // Fase 32: Capturar excepciones críticas durante la inicialización
            try
            {
                SetupGrid();
                CargarSedes();
                // Fase 19: Trasladar CargarSubunidades al evento Shown para evitar modal bloqueante antes del renderizado
                this.Shown += (s, ev) => CargarSubunidades();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al inicializar FormSedes", ex);
                // Fase 36: Inicializar fuente de datos en caso de fallo para evitar efecto cascada en eventos de renderizado
                _bs.DataSource = new List<Sede>();
                MessageBox.Show($"Error al inicializar el formulario de sedes: {ex.Message}",
                    "Error crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarSubunidades()
        {
            // Fase 4: Guarda de reentrancia para evitar múltiples llamadas concurrentes y acumulación de MessageBox
            if (_cargandoSubunidades) return;
            _cargandoSubunidades = true;

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetSubunidades();
                // Fase 2: Control adicional para evitar NullReferenceException si result.Value llega nulo pese al IsSuccess
                if (result.IsSuccess && result.Value != null)
                {
                    cmbTerritoriales.Items.Clear();
                    foreach (var s in result.Value)
                    {
                        cmbTerritoriales.Items.Add(s);
                    }
                    // Fase 25: No sobrescribir la selección si ya hay una válida (aunque tras el Clear suele ser -1, protege la intención semántica)
                    if (cmbTerritoriales.Items.Count > 0 && cmbTerritoriales.SelectedIndex < 0)
                        cmbTerritoriales.SelectedIndex = 0;
                }
                else
                {
                    LoggerHelper.Warning($"Error al cargar subunidades: {result.Error}");
                    cmbTerritoriales.Items.Clear();
                    cmbTerritoriales.Items.Add("(Error - click Reintentar)");
                    // Fase 20: Prevención de OutOfRange si el listado quedó vacío
                    if (cmbTerritoriales.Items.Count > 0) cmbTerritoriales.SelectedIndex = 0;
                    MessageBox.Show("No se pudieron cargar las territoriales. Use el botón Reintentar.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al cargar subunidades en FormSedes", ex);
                cmbTerritoriales.Items.Clear();
                cmbTerritoriales.Items.Add("(Error - click Reintentar)");
                // Fase 20: Prevención de OutOfRange si el listado quedó vacío
                if (cmbTerritoriales.Items.Count > 0) cmbTerritoriales.SelectedIndex = 0;
                MessageBox.Show("No se pudieron cargar las territoriales. Use el botón Reintentar.", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _cargandoSubunidades = false;
            }
        }

        private void btnReintentarSubunidades_Click(object sender, EventArgs e)
        {
            CargarSubunidades();
        }

        private void SetupGrid()
        {
            dgvSedes.AutoGenerateColumns = false;
            dgvSedes.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Ciudad", HeaderText = "Ciudad", Width = 150 });
            dgvSedes.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Departamento", HeaderText = "Departamento", Width = 150 });
            // Fase 14: Establecer ValueType numérico explícitamente para evitar excepciones de casteo implícito en la UI de WinForms
            dgvSedes.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Categoria", HeaderText = "Categoría", Width = 80, ValueType = typeof(int) });
            dgvSedes.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Activo", HeaderText = "Activo", Width = 60, ValueType = typeof(int) });
            dgvSedes.DataSource = _bs;
            
            dgvSedes.SelectionChanged += DgvSedes_SelectionChanged;
        }

        private void CargarSedes()
        {
            // Fase 23: Desestimar carga concurrente
            if (_cargandoSedes) return;
            _cargandoSedes = true;
            
            try
            {
                // Fase 16: Limpiar estado local antes de recargar para evitar entidades huérfanas
                _currentSede = null;
                btnEliminar.Enabled = false;
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetSedes();
                // Fase 6: Control adicional para evitar NullReferenceException si result.Value llega nulo
                if (result.IsSuccess && result.Value != null)
                {
                    _bs.DataSource = result.Value.ToList();
                }
                else
                {
                    LoggerHelper.Warning($"[FormSedes] Error al cargar sedes: {result.Error}");
                    MessageBox.Show($"Error al cargar sedes: {result.Error}", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar sedes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cargandoSedes = false;
            }
        }

        private void DgvSedes_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvSedes.CurrentRow?.DataBoundItem is Sede s)
            {
                _currentSede = s;
                txtCiudad.Text = s.Ciudad;
                txtDepartamento.Text = s.Departamento;
                txtDepartamentoSiif.Text = s.DepartamentoSiif;
                txtCiudadSiif.Text = s.CiudadSiif;
                txtCategoria.Text = s.Categoria.ToString();
                txtSubsede.Text = s.Subsede;
                
                // Fase 24: Retirado el MessageBox para evitar reentrancias silenciosas ocasionadas por el message pump modal dentro del evento de cambio de fila
                if ((decimal)s.Transporte > numTransporte.Maximum)
                {
                    numTransporte.Maximum = (decimal)s.Transporte;
                }
                numTransporte.Value = (decimal)s.Transporte;
                
                // Fase 13: Evitar ArgumentOutOfRangeException si el item no existe
                int idx = -1;
                for (int i = 0; i < cmbActivo.Items.Count; i++)
                {
                    // Fase 33: Comprobación segura de tipos con Pattern Matching evitando InvalidCastException
                    if (cmbActivo.Items[i] is ActivoOption opt && opt.Value == s.Activo)
                    {
                        idx = i;
                        break;
                    }
                }
                if (idx < 0) idx = cmbActivo.Items.Count > 0 ? 0 : -1;
                if (idx >= 0) cmbActivo.SelectedIndex = idx;
                
                // Fase 11: Corrección type-mismatch al verificar elementos del ComboBox
                if (!string.IsNullOrEmpty(s.Subunidad) && cmbTerritoriales.Items.Cast<string>().Any(i => i == s.Subunidad))
                {
                    cmbTerritoriales.SelectedItem = s.Subunidad;
                }
                else
                {
                    cmbTerritoriales.SelectedIndex = -1;
                }
                
                btnEliminar.Enabled = true;
            }
            else
            {
                LimpiarCampos();
            }
        }

        private void LimpiarCampos()
        {
            _currentSede = null;
            txtCiudad.Clear();
            txtDepartamento.Clear();
            txtDepartamentoSiif.Clear();
            txtCiudadSiif.Clear();
            txtCategoria.Clear();
            txtSubsede.Clear();
            
            // Fase 15: Restaurar el máximo original antes de aplicar el 0
            numTransporte.Maximum = _numTransporteMaxOriginal;
            numTransporte.Value = 0;
            
            // Fase 37: Prevenir ArgumentOutOfRangeException si los comboboxes están vacíos
            if (cmbActivo.Items.Count > 1) cmbActivo.SelectedIndex = 1;
            if (cmbTerritoriales.Items.Count > 0) cmbTerritoriales.SelectedIndex = -1;
            btnEliminar.Enabled = false;
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
            // Fase 9: txtCiudad es ReadOnly, enfocar el botón correcto
            btnSelectDivipola.Focus();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            // Fase 18: Retornar inmediatamente si el proceso de guardado ya se encuentra en ejecución
            if (_guardando) return;

            if (string.IsNullOrWhiteSpace(txtCiudad.Text))
            {
                MessageBox.Show("Ciudad es obligatoria.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var s = _currentSede ?? new Sede();
            s.Ciudad = txtCiudad.Text;
            s.Departamento = txtDepartamento.Text;
            s.DepartamentoSiif = txtDepartamentoSiif.Text;
            s.CiudadSiif = txtCiudadSiif.Text;
            s.Subsede = txtSubsede.Text;
            if (int.TryParse(txtCategoria.Text, out int cat))
            {
                s.Categoria = cat;
            }
            s.Transporte = numTransporte.Value;
            s.Activo = cmbActivo.SelectedItem is ActivoOption opt ? opt.Value : 1;
            s.Subunidad = cmbTerritoriales.SelectedItem?.ToString();
            
            if (s.Subunidad == "(Error - click Reintentar)")
            {
                MessageBox.Show("No se pudo determinar la territorial. Use el botón Reintentar.", 
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _guardando = true;
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                
                if (s.Id == 0)
                {
                    var result = repo.InsertSede(s);
                    if (result.IsFailure) { MessageBox.Show($"Error al insertar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    s.Id = result.Value;
                }
                else
                {
                    var result = repo.UpdateSede(s);
                    if (result.IsFailure) { MessageBox.Show($"Error al actualizar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                }

                // Nota: La persistencia al Vault ahora la gestiona el SyncScheduler de forma asíncrona.
                
                MessageBox.Show("Sede guardada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarSedes();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _guardando = false;
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentSede == null) return;

            if (MessageBox.Show($"¿Está seguro de eliminar la sede de {_currentSede.Ciudad}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var db = _connectionFactory();
                    var repo = new AcademicRepository(db);
                    var result = repo.DeleteSede(_currentSede.Id);
                    if (result.IsFailure) { MessageBox.Show($"Error al eliminar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                    CargarSedes();
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

        /// <summary>
        /// Mapea los datos de Divipola a una Sede.
        /// Nota intencional (Fase 10): No pre-selecciona automáticamente cmbTerritoriales
        /// basándose en el municipio, para permitir que el sistema asuma por defecto la territorial 
        /// del usuario autenticado (comportamiento legacy para caja/presupuesto).
        /// </summary>
        private void MapearDivipolaASede(Divipola divipola)
        {
            // Fase 31: Prevenir reentrancia originada por confirmaciones modales o selecciones asíncronas
            if (_mapeandoDivipola) return;
            _mapeandoDivipola = true;
            try
            {
                // Fase 12: Clarificar lógica para no descartar silenciosamente el Id ni duplicar filas al mapear sobre sede existente
                bool esNueva = (_currentSede == null);
                if (esNueva) _currentSede = new Sede();
                
                // Fase 22: Advertir sobre la sobrescritura del mapeo geográfico si la sede no es nueva y ya tenía otra Divipola asignada
                if (!esNueva && _currentSede.IdDivipola.HasValue && _currentSede.IdDivipola != divipola.Id)
                {
                    var confirm = MessageBox.Show(
                        $"La sede ya está asociada a otra división política. ¿Desea reemplazarla?",
                        "Confirmar cambio", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirm != DialogResult.Yes) return;
                }
                
                _currentSede.IdDivipola = divipola.Id;
                _currentSede.Ciudad = divipola.Municipio;
                _currentSede.Departamento = divipola.Departamento;
                _currentSede.Categoria = divipola.MunicipioCategoria;
                _currentSede.Activo = 1;
                
                txtCiudad.Text = divipola.Municipio;
                txtDepartamento.Text = divipola.Departamento;
                txtCategoria.Text = divipola.MunicipioCategoria.ToString();
                
                // Fase 17: Limpiar residuos visuales de la sede previa
                txtSubsede.Text = string.Empty;
                _currentSede.Subsede = string.Empty;
                
                txtDepartamentoSiif.Text = string.Empty;
                _currentSede.DepartamentoSiif = string.Empty;
                txtCiudadSiif.Text = string.Empty;
                _currentSede.CiudadSiif = string.Empty;
                
                numTransporte.Maximum = _numTransporteMaxOriginal;
                numTransporte.Value = 0;
                _currentSede.Transporte = 0;
                
                cmbActivo.SelectedIndex = 1;
                // No blanqueamos subunidad aquí para que el usuario pueda escogerla, 
                // a menos que ya esté configurada.
            }
            finally { _mapeandoDivipola = false; }
        }

        private void btnSelectDivipola_Click(object sender, EventArgs e)
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                using (var frm = new FormDivipolaBuscador(repo))
                {
                    if (frm.ShowDialog() == DialogResult.OK && frm.SelectedDivipola != null)
                    {
                        MapearDivipolaASede(frm.SelectedDivipola);
                        // Fase 38: Evitar desreferencia de nulos si el mapeo retornó anticipadamente por reentrancia
                        if (_currentSede == null) return;
                        // Fase 27: Enlazar entidad nueva a la lista e iterar Grid para autoselección tras mapear
                        bool esNueva = _currentSede.Id == 0;
                        if (esNueva && _bs.DataSource is IList<Sede> list && !list.Contains(_currentSede))
                        {
                            list.Add(_currentSede);
                            _bs.ResetBindings(false);
                            // Fase 29: Proteger asignación de celda actual contra grillas vacías
                            if (dgvSedes.Rows.Count > 0 && dgvSedes.Columns.Count > 0)
                            {
                                for (int i = 0; i < dgvSedes.Rows.Count; i++)
                                {
                                    if (dgvSedes.Rows[i].DataBoundItem == _currentSede)
                                    { dgvSedes.CurrentCell = dgvSedes.Rows[i].Cells[0]; break; }
                                }
                            }
                        }
                        // Fase 28: Defensa de referencia nula por si el Data Source fue vaciado por un error de BD previo
                        else if (_bs.DataSource != null)
                        {
                            _bs.ResetBindings(false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar buscador de divipola: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDivipolaCRUD_Click(object sender, EventArgs e)
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                using (var frm = new FormDivipolaCRUD(repo))
                {
                    frm.ShowDialog();
                }
                // Fase 21: Refrescar el maestro de Sedes al regresar del CRUD, ya que una sede puede haber modificado indirectamente la categoría de su Divipola
                CargarSedes();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar CRUD de divipola: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
