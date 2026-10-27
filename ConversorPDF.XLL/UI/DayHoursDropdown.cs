using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.UI
{
    /// <summary>
    /// Control personalizado que actúa como un ComboBox multiselect para días de la semana y horas.
    /// Basado en DataGridView hospedado en un ToolStripDropDown.
    /// </summary>
    public partial class DayHoursDropdown : UserControl
    {
        private TextBox txtDisplay;
        private Button btnDropDown;
        private ToolStripDropDown _dropDown;
        private DataGridView _grid;
        private bool _isClosingFromCode = false;

        private readonly string[] _diasSemana = { "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO" };
        private Dictionary<string, DaySessionInfo> _selectedDays = new Dictionary<string, DaySessionInfo>();

        public event EventHandler SelectedDaysChanged;

        [Browsable(true)]
        [Category("Data")]
        public Dictionary<string, DaySessionInfo> SelectedDays
        {
            get => _selectedDays;
            set
            {
                _selectedDays = value ?? new Dictionary<string, DaySessionInfo>();
                UpdateDisplayText();
            }
        }

        public string PlaceholderText { get; set; } = "Seleccionar días y horas...";

        public DayHoursDropdown()
        {
            InitializeComponents();
            SetupDropDown();
            UpdateDisplayText();
        }

        private void InitializeComponents()
        {
            this.Size = new Size(220, 23);
            this.Padding = new Padding(0);
            this.BackColor = Color.FromArgb(171, 171, 171); // Borde simulado

            txtDisplay = new TextBox
            {
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Location = new Point(1, 1),
                Size = new Size(this.Width - 19, this.Height - 2),
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            txtDisplay.Click += (s, e) => ShowDropDown();

            btnDropDown = new Button
            {
                Text = "▼",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                FlatAppearance = { BorderSize = 0 },
                Location = new Point(this.Width - 18, 1),
                Size = new Size(17, this.Height - 2),
                Font = new Font("Segoe UI", 7f),
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(0)
            };
            btnDropDown.Click += (s, e) => ShowDropDown();

            this.Controls.Add(txtDisplay);
            this.Controls.Add(btnDropDown);
        }

        private void SetupDropDown()
        {
            _grid = new DataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = false,
                ColumnHeadersVisible = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                ScrollBars = ScrollBars.None,
                GridColor = Color.FromArgb(224, 224, 224),
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 25 },
                Width = 250,
                Height = 175,
                Font = new Font("Segoe UI", 9f)
            };

            _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ColCheck", Width = 30 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDay", Width = 90, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColHours", Width = 50, MaxInputLength = 2 });
            
            var colSchedule = new DataGridViewComboBoxColumn { Name = "ColSchedule", Width = 150, FlatStyle = FlatStyle.Flat };
            colSchedule.Items.AddRange(
                "08:00 A 12:00",
                "13:00 A 17:00",
                "18:00 A 22:00",
                "08:00 A 10:00",
                "10:00 A 12:00",
                "13:00 A 15:00",
                "15:00 A 17:00",
                "18:00 A 20:00",
                "20:00 A 22:00",
                "08:00 A 17:00",
                "14:00 A 22:00"
            );
            _grid.Columns.Add(colSchedule);
            _grid.Width = 340; // Expand width to fit new column

            foreach (var dia in _diasSemana)
            {
                _grid.Rows.Add(false, dia, "", "");
            }

            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;
            _grid.EditingControlShowing += Grid_EditingControlShowing;
            _grid.KeyDown += Grid_KeyDown;
            _grid.CellClick += Grid_CellClick;
            _grid.DataError += Grid_DataError;

            // Corrección Fase 19: Forzar el tamaño del host para respetar el diseño
            var host = new ToolStripControlHost(_grid) { Margin = Padding.Empty, Padding = Padding.Empty, AutoSize = false, Size = _grid.Size };
            _dropDown = new ToolStripDropDown { Padding = Padding.Empty };
            _dropDown.Items.Add(host);
            _dropDown.Closed += DropDown_Closed;
        }

        private void ShowDropDown()
        {
            // Corrección Fase 62: Prevenir invocación múltiple si ya está visible
            if (_dropDown != null && _dropDown.Visible) return;

            // Corrección Fase 15: Limpiar bandera por si hubo una interrupción en un cierre previo
            _isClosingFromCode = false;

            // Corrección Fase 11, 26 y 50: Suspender eventos recursivos y DataError durante la carga
            _grid.CellValueChanged -= Grid_CellValueChanged;
            _grid.CurrentCellDirtyStateChanged -= Grid_CurrentCellDirtyStateChanged;
            _grid.DataError -= Grid_DataError;
            try
            {
                // Cargar datos actuales en el grid
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    // Corrección: Usar null-conditional y coalescing para prevenir NullReferenceException (Error 6)
                    string day = row.Cells["ColDay"].Value?.ToString() ?? "";
                    if (_selectedDays.TryGetValue(day, out DaySessionInfo info))
                    {
                        row.Cells["ColCheck"].Value = true;
                        row.Cells["ColHours"].Value = info.Horas.ToString();
                        
                        // Actualizar las opciones en cascada antes de asignar el valor
                        ActualizarOpcionesHorario(row, info.Horas.ToString());
                        
                        // Corrección Fase 6 & F651 / F660: Validar si el horario pertenece a la colección Items (o su equivalente militar) antes de asignar
                        var combo = row.Cells["ColSchedule"] as DataGridViewComboBoxCell;
                        string horNorm = ConversorPDF.XLL.Academic.Application.Services.HorarioMilitarizador.NormalizarHorario(info.Horario);
                        if (combo != null && combo.Items.Cast<string>().Contains(horNorm))
                            row.Cells["ColSchedule"].Value = horNorm;
                        else if (combo != null && combo.Items.Cast<string>().Contains(info.Horario))
                            row.Cells["ColSchedule"].Value = info.Horario;
                        else
                        {
                            // Fase 1001: no perder franjas en silencio si el catálogo no contiene el horario
                            LoggerHelper.Warning(
                                $"[DayHoursDropdown] Horario '{info.Horario}' ({day}) no existe en opciones ({string.Join(", ", combo?.Items.Cast<string>() ?? new string[0])}); se descartará al guardar.");
                            row.Cells["ColSchedule"].Value = DBNull.Value;
                        }

                        ApplyRowHighlight(row, true);
                    }
                    else
                    {
                        row.Cells["ColCheck"].Value = false;
                        row.Cells["ColHours"].Value = "";
                        
                        ActualizarOpcionesHorario(row, ""); // Opciones por defecto
                        
                        row.Cells["ColSchedule"].Value = DBNull.Value;
                        ApplyRowHighlight(row, false);
                    }
                }
            }
            finally
            {
                _grid.DataError += Grid_DataError;
                _grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;
                _grid.CellValueChanged += Grid_CellValueChanged;
            }

            // Corrección Fase 31: Limpiar selección previa
            _grid.ClearSelection();
            _grid.CurrentCell = null;

            _dropDown.Show(this, new Point(0, this.Height));
            _grid.Focus();
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Si se hace click en Check o en el nombre del Día, alternamos
            if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
            {
                // Corrección Fase 40: Desactivar DataError al commitear
                _grid.DataError -= Grid_DataError;
                try
                {
                    // Corrección Fase 9: Usar CommitEdit en lugar de EndEdit para guardar cambios en celdas previamente activas
                    if (_grid.IsCurrentCellInEditMode)
                        _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);

                    // Corrección Fase 13: Terminar edición explícitamente si la celda activa era la de horas
                    if (_grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == 2)
                        _grid.EndEdit();
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }

                var cell = _grid.Rows[e.RowIndex].Cells["ColCheck"];
                // Corrección: Evitar InvalidCastException y NullReferenceException al leer el estado actual de la celda
                bool currentValue = cell.Value is bool b ? b : false;
                
                // Corrección Fase 78: Desactivar DataError al asignar programáticamente a DataGridViewCheckBoxCell
                _grid.DataError -= Grid_DataError;
                try
                {
                    cell.Value = !currentValue;
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }
            }
        }

        private void Grid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Omitir errores de DataGridView generados al limpiar o cambiar el Items de un ComboBoxCell dinámico
            e.ThrowException = false;
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (e.ColumnIndex == 0) // Checkbox
            {
                var row = _grid.Rows[e.RowIndex];
                // Corrección: Prevenir InvalidCastException usando pattern matching (la celda puede ser DBNull/null)
                bool isChecked = row.Cells["ColCheck"].Value is bool cb && cb;
                ApplyRowHighlight(row, isChecked);

                // Auto-llenar horas si está checked y vacío
                if (isChecked && string.IsNullOrWhiteSpace(row.Cells["ColHours"].Value?.ToString()))
                {
                    // Corrección Fase 71: Desactivar DataError al auto-llenar
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColHours"].Value = "4"; 
                        // Corrección: Sincronizar explícitamente las opciones del ComboBox para asegurar consistencia (Error 12)
                        ActualizarOpcionesHorario(row, "4");
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                }
                
                // Limpiar si se desmarca
                if (!isChecked)
                {
                    // Corrección Fase 17: Forzar EndEdit si hay edición pendiente al desmarcar
                    if (_grid.CurrentCell != null && _grid.CurrentCell.RowIndex == row.Index && _grid.IsCurrentCellInEditMode)
                        _grid.EndEdit();

                    // Corrección Fase 42: Desactivar DataError al limpiar
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColHours"].Value = "";
                        row.Cells["ColSchedule"].Value = DBNull.Value;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    // Corrección Fase 24: Asegurar highlight correcto
                    ApplyRowHighlight(row, false);
                }
            }
            else if (e.ColumnIndex == 2) // Horas
            {
                var row = _grid.Rows[e.RowIndex];
                string val = row.Cells["ColHours"].Value?.ToString();
                
                // Corrección Fase 38 y 28: Desactivar DataError aquí en vez de dentro de ActualizarOpcionesHorario
                _grid.DataError -= Grid_DataError;
                try
                {
                    // Aplicar lógica en cascada para filtrar franjas horarias
                    ActualizarOpcionesHorario(row, val);
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }

                // Corrección Fase 4: Leer el ColSchedule DESPUÉS de aplicar el filtro, pues ActualizarOpcionesHorario pudo haberlo borrado
                string sched = row.Cells["ColSchedule"].Value?.ToString();

                // Corrección Fase 64: Desactivar DataError al asignar programáticamente el CheckBox
                if (int.TryParse(val, out int hr) && hr >= 1 && hr <= 8 && !string.IsNullOrEmpty(sched))
                {
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = true;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, true);
                }
                // Corrección Fase 20: Desmarcar solo si el valor es explícitamente inválido, no cuando está en blanco mientras digita
                else if (!string.IsNullOrEmpty(val) && (!int.TryParse(val, out hr) || hr < 1 || hr > 8))
                {
                    // Corrección Fase 4 y 72: Desmarcar explícitamente desactivando DataError
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = false;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, false);
                }
                // Corrección Fase 35 y 72: Desmarcar si el usuario borró las horas manualmente
                else if (string.IsNullOrEmpty(val))
                {
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = false;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, false);
                }
            }
            else if (e.ColumnIndex == 3) // Horario (Schedule)
            {
                var row = _grid.Rows[e.RowIndex];
                string sched = row.Cells["ColSchedule"].Value?.ToString();

                // Corrección Fase 45 y 73: Desmarcar si se borró el horario suprimiendo DataError
                if (string.IsNullOrEmpty(sched))
                {
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = false;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, false);
                    return;
                }

                // Corrección Fase 10: Auto-completar horas si está vacío y se escogió un horario
                if (string.IsNullOrEmpty(row.Cells["ColHours"].Value?.ToString()) && !string.IsNullOrEmpty(sched))
                {
                    row.Cells["ColHours"].Value = "4";
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        ActualizarOpcionesHorario(row, "4");
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                }

                string val = row.Cells["ColHours"].Value?.ToString();

                // Corrección Fase 65: Desactivar DataError al asignar programáticamente el CheckBox
                if (int.TryParse(val, out int hr) && hr >= 1 && hr <= 8 && !string.IsNullOrEmpty(sched))
                {
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = true;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, true);
                }
                else
                {
                    // Corrección Fase 4 y 74: Desmarcar explícitamente si dejó de ser válido tras el filtrado
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        row.Cells["ColCheck"].Value = false;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    ApplyRowHighlight(row, false);
                }
            }
        }

        /// <summary>
        /// Filtra las opciones del ComboBox de horarios en cascada según las horas digitadas.
        /// </summary>
        private void ActualizarOpcionesHorario(DataGridViewRow row, string horasStr)
        {
            if (!(row.Cells["ColSchedule"] is DataGridViewComboBoxCell comboCell)) return;

            // Corrección Fase 14: Proteger contra edición activa en la celda combobox antes de modificar Items
            if (_grid.CurrentCell != null && _grid.CurrentCell.RowIndex == row.Index && _grid.CurrentCell.ColumnIndex == 3 && _grid.IsCurrentCellInEditMode)
                _grid.EndEdit();

            var opciones2Horas = new[] { "08:00 A 10:00", "10:00 A 12:00", "13:00 A 15:00", "15:00 A 17:00", "18:00 A 20:00", "20:00 A 22:00" };
            var opciones4Horas = new[] { "08:00 A 12:00", "13:00 A 17:00", "18:00 A 22:00" };
            var opciones8Horas = new[] { "08:00 A 17:00", "14:00 A 22:00" };

            string[] opcionesPermitidas;
            if (int.TryParse(horasStr, out int h) && h >= 1 && h <= 8)
            {
                if (h <= 2) opcionesPermitidas = opciones2Horas;
                else if (h <= 4) opcionesPermitidas = opciones4Horas;
                else opcionesPermitidas = opciones8Horas;
            }
            else
            {
                // Corrección Fase 51: Retornar todas las opciones concatenadas en lugar de un array vacío como fallback para horas atípicas o vacías
                opcionesPermitidas = opciones2Horas.Concat(opciones4Horas).Concat(opciones8Horas).ToArray();
            }

            string valorActual = comboCell.Value?.ToString();

            comboCell.Items.Clear();
            foreach (var opt in opcionesPermitidas)
            {
                comboCell.Items.Add(opt);
            }

            // Corrección Fase 5 y 37: Revertir a case-sensitive exacta (Contains normal) para evitar DataError al asignar valores no idénticos
            if (!string.IsNullOrEmpty(valorActual) && !opcionesPermitidas.Contains(valorActual))
            {
                comboCell.Value = DBNull.Value; // Fase 37 y 43: Usar DBNull en vez de null
            }
        }

        private void ApplyRowHighlight(DataGridViewRow row, bool highlight)
        {
            row.DefaultCellStyle.BackColor = highlight ? SystemColors.Highlight : Color.White;
            row.DefaultCellStyle.ForeColor = highlight ? SystemColors.HighlightText : SystemColors.ControlText;
            row.DefaultCellStyle.SelectionBackColor = highlight ? SystemColors.Highlight : Color.White;
            row.DefaultCellStyle.SelectionForeColor = highlight ? SystemColors.HighlightText : SystemColors.ControlText;
        }

        private void Grid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            // Forzamos el commit en celdas CheckBox
            if (_grid.IsCurrentCellDirty && _grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex != 3)
            {
                // Corrección Fase 47: Desactivar DataError al commitear
                _grid.DataError -= Grid_DataError;
                try
                {
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }
            }
        }

        private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // Corrección Fase 18: Desuscribir siempre de otras columnas porque el Control de edición se reutiliza
            if (e.Control is TextBox tb)
            {
                tb.KeyPress -= ValidarNumeros;
                tb.TextChanged -= Horas_TextChanged;
                
                // Corrección: Verificar _grid.CurrentCell != null para evitar NullReferenceException si el grid pierde el foco
                if (_grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == 2)
                {
                    tb.KeyPress += ValidarNumeros;
                    tb.TextChanged += Horas_TextChanged;
                }
            }
        }

        private void ValidarNumeros(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
        }

        // Corrección Fase 30 y 36: Validar pegado de texto (Ctrl+V) y prevenir bucle infinito
        private void Horas_TextChanged(object sender, EventArgs e)
        {
            if (sender is TextBox tb && !string.IsNullOrEmpty(tb.Text))
            {
                if (!int.TryParse(tb.Text, out int hr) || hr < 1 || hr > 8)
                {
                    tb.TextChanged -= Horas_TextChanged;
                    try
                    {
                        string safeText = new string(tb.Text.Where(char.IsDigit).ToArray());
                        if (int.TryParse(safeText, out int safeHr) && safeHr >= 1 && safeHr <= 8)
                            tb.Text = safeText;
                        else
                            tb.Text = "";
                    }
                    finally
                    {
                        tb.TextChanged += Horas_TextChanged;
                    }
                }
            }
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                // Corrección Fase 7 y 39: Commitear explícitamente y finalizar edición antes de cerrar, desactivando DataError
                _grid.DataError -= Grid_DataError;
                try
                {
                    if (_grid.CurrentCell != null && _grid.IsCurrentCellDirty)
                        _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    if (_grid.IsCurrentCellInEditMode)
                        _grid.EndEdit();
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }

                _dropDown.Close(ToolStripDropDownCloseReason.ItemClicked);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                _isClosingFromCode = true;
                // Corrección Fase 54: Usar CloseCalled para indicar cierre programático cancelado
                _dropDown.Close(ToolStripDropDownCloseReason.CloseCalled);
            }
            // Corrección: Prevenir activación si CurrentCell es null, controlando el shortcut (Error 10)
            else if (e.KeyCode == Keys.Space && (_grid.CurrentCell == null || _grid.CurrentCell.ColumnIndex != 2))
            {
                // Corrección Fase 41: Desactivar DataError al commitear con teclado
                _grid.DataError -= Grid_DataError;
                try
                {
                    // Corrección Fase 12: Commitear edición activa antes de alternar checkbox
                    if (_grid.IsCurrentCellInEditMode)
                        _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
                finally
                {
                    _grid.DataError += Grid_DataError;
                }

                // Corrección: Validar _grid.CurrentRow != null y usar casteo seguro
                if (_grid.CurrentRow != null)
                {
                    var cell = _grid.CurrentRow.Cells["ColCheck"];
                    bool currentValue = cell.Value is bool b ? b : false;
                    // Corrección Fase 79: Desactivar DataError por si el click espacial choca contra celda en edición
                    _grid.DataError -= Grid_DataError;
                    try
                    {
                        cell.Value = !currentValue;
                    }
                    finally
                    {
                        _grid.DataError += Grid_DataError;
                    }
                    // Corrección Fase 23: Removido el _grid.EndEdit() innecesario
                }
                e.Handled = true;
            }
        }

        private void DropDown_Closed(object sender, ToolStripDropDownClosedEventArgs e)
        {
            if (_isClosingFromCode)
            {
                _isClosingFromCode = false;
                return;
            }

            // FORZAR COMMIT de la celda activa antes de leer
            // Corrección Fase 27 y 46: Desactivar DataError durante commit masivo y lectura final
            _grid.DataError -= Grid_DataError;
            var newSelections = new Dictionary<string, DaySessionInfo>();
            try
            {
                // Corrección Fase 8: Invertir el orden (CommitEdit primero, EndEdit después)
                if (_grid.CurrentCell != null && _grid.IsCurrentCellDirty)
                {
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
                // Corrección Fase 16 & 22: Verificar CurrentCell != null antes de llamar a EndEdit para evitar excepciones de WinForms
                if (_grid.CurrentCell != null && _grid.IsCurrentCellInEditMode)
                {
                    _grid.EndEdit();
                }

                // Guardar cambios
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    // Corrección: Prevenir InvalidCastException si el valor de la celda es null o DBNull
                    bool isChecked = row.Cells["ColCheck"].Value is bool cb && cb;
                    string val = row.Cells["ColHours"].Value?.ToString();
                    string sched = row.Cells["ColSchedule"].Value?.ToString();
                    // Corrección: Asegurar que el día no sea null antes de usarlo como llave en el diccionario (Error 7)
                    string day = row.Cells["ColDay"].Value?.ToString() ?? "";

                    // Corrección: Validar que el día no esté vacío y que las horas estén entre 1 y 8 (Error 7 y 9)
                    if (isChecked && int.TryParse(val, out int hr) && hr >= 1 && hr <= 8 && !string.IsNullOrEmpty(sched) && !string.IsNullOrEmpty(day))
                    {
                        newSelections[day] = new DaySessionInfo(hr, sched);
                    }
                }
            }
            finally
            {
                _grid.DataError += Grid_DataError;
            }

            _selectedDays = newSelections;
            UpdateDisplayText();
            SelectedDaysChanged?.Invoke(this, EventArgs.Empty);
        }

        private void UpdateDisplayText()
        {
            if (_selectedDays == null || _selectedDays.Count == 0)
            {
                txtDisplay.Text = PlaceholderText;
                txtDisplay.ForeColor = Color.Gray;
                return;
            }

            // Orden cronológico basado en _diasSemana
            var sorted = _selectedDays
                .OrderBy(x => Array.IndexOf(_diasSemana, x.Key))
                .Select(x => $"{x.Key}|{x.Value.Horas}|{x.Value.Horario}");

            txtDisplay.Text = string.Join(", ", sorted);
            txtDisplay.ForeColor = SystemColors.WindowText;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (txtDisplay != null)
            {
                txtDisplay.Size = new Size(this.Width - 19, this.Height - 2);
            }
            if (btnDropDown != null)
            {
                btnDropDown.Location = new Point(this.Width - 18, 1);
                btnDropDown.Height = this.Height - 2;
            }
        }
    }
}
