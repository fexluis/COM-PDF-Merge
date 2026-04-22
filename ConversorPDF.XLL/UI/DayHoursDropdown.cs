using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Windows.Forms;

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
        private Dictionary<string, int> _selectedDays = new Dictionary<string, int>();

        public event EventHandler SelectedDaysChanged;

        [Browsable(true)]
        [Category("Data")]
        public Dictionary<string, int> SelectedDays
        {
            get => _selectedDays;
            set
            {
                _selectedDays = value ?? new Dictionary<string, int>();
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
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDay", Width = 130, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColHours", Width = 70, MaxInputLength = 2 });

            foreach (var dia in _diasSemana)
            {
                _grid.Rows.Add(false, dia, "");
            }

            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;
            _grid.EditingControlShowing += Grid_EditingControlShowing;
            _grid.KeyDown += Grid_KeyDown;
            _grid.CellClick += Grid_CellClick;

            var host = new ToolStripControlHost(_grid) { Margin = Padding.Empty, Padding = Padding.Empty };
            _dropDown = new ToolStripDropDown { Padding = Padding.Empty };
            _dropDown.Items.Add(host);
            _dropDown.Closed += DropDown_Closed;
        }

        private void ShowDropDown()
        {
            // Cargar datos actuales en el grid
            foreach (DataGridViewRow row in _grid.Rows)
            {
                string day = row.Cells["ColDay"].Value.ToString();
                if (_selectedDays.TryGetValue(day, out int hours))
                {
                    row.Cells["ColCheck"].Value = true;
                    row.Cells["ColHours"].Value = hours.ToString();
                    ApplyRowHighlight(row, true);
                }
                else
                {
                    row.Cells["ColCheck"].Value = false;
                    row.Cells["ColHours"].Value = "";
                    ApplyRowHighlight(row, false);
                }
            }

            _dropDown.Show(this, new Point(0, this.Height));
            _grid.Focus();
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Si se hace click en Check o en el nombre del Día, alternamos
            if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
            {
                var cell = _grid.Rows[e.RowIndex].Cells["ColCheck"];
                cell.Value = !(bool)cell.Value;
                _grid.EndEdit();
            }
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (e.ColumnIndex == 0) // Check
            {
                bool isChecked = (bool)_grid.Rows[e.RowIndex].Cells["ColCheck"].Value;
                ApplyRowHighlight(_grid.Rows[e.RowIndex], isChecked);
                
                // Si se desmarca, limpiamos horas
                if (!isChecked) _grid.Rows[e.RowIndex].Cells["ColHours"].Value = "";
            }
            else if (e.ColumnIndex == 2) // Hours
            {
                var row = _grid.Rows[e.RowIndex];
                string val = row.Cells["ColHours"].Value?.ToString();
                
                if (int.TryParse(val, out int hr) && hr > 0)
                {
                    row.Cells["ColCheck"].Value = true;
                }
                else
                {
                    // Si deja vacío o 0, desmarcamos
                    row.Cells["ColCheck"].Value = false;
                    row.Cells["ColHours"].Value = "";
                }
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
            if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (_grid.CurrentCell.ColumnIndex == 2 && e.Control is TextBox tb)
            {
                tb.KeyPress -= ValidarNumeros;
                tb.KeyPress += ValidarNumeros;
            }
        }

        private void ValidarNumeros(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                _dropDown.Close(ToolStripDropDownCloseReason.ItemClicked);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                _isClosingFromCode = true; // Indicar que no queremos guardar
                _dropDown.Close(ToolStripDropDownCloseReason.CloseCalled);
            }
            else if (e.KeyCode == Keys.Space && _grid.CurrentCell?.ColumnIndex != 2)
            {
                var cell = _grid.CurrentRow.Cells["ColCheck"];
                cell.Value = !(bool)cell.Value;
                _grid.EndEdit();
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

            // Guardar cambios
            var newSelections = new Dictionary<string, int>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                bool isChecked = (bool)row.Cells["ColCheck"].Value;
                string val = row.Cells["ColHours"].Value?.ToString();

                if (isChecked && int.TryParse(val, out int hr) && hr > 0)
                {
                    newSelections[row.Cells["ColDay"].Value.ToString()] = hr;
                }
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
                .Select(x => $"{x.Key}|{x.Value}");

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
