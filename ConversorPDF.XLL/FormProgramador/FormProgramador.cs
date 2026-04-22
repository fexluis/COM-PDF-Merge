using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ConversorPDF.XLL;
using ConversorPDF.XLL.Core.Services;
namespace ConversorPDF.XLL.FormProgramador
{
    public partial class FormProgramador : Form
    {
        private TipoFormMode currentMode;
        private dynamic activeWorkbook;
        private IViaticosService _viaticosService;

        private Button[] dayButtons = new Button[42];
        private int currentMonth;
        private int currentYear;
        private TextBox targetTextBox = null;

        private const int WIDE_DEPENDENCIAS = 1200;
        private const int WIDE_DESTINOS = 800;
        private const int WIDE_CATEGORIA1 = 500;

        public DateTime SelectedFechaInicio { get; private set; }
        public DateTime SelectedFechaFin { get; private set; }
        public string SelectedCDP { get; private set; }
        public string SelectedDependenciaDescripcion { get; private set; }


        public FormProgramador(TipoFormMode mode, dynamic wb, IViaticosService viaticosService)
        {
            this.currentMode = mode;
            this.activeWorkbook = wb;
            this._viaticosService = viaticosService;

            InitializeComponent();

            InitializeCalendarGrid();
            AttachEvents();


            this.Load += FormProgramador_Load;
        }

        private void InitializeCalendarGrid()
        {
            int btnWidth = 49;
            int btnHeight = 31;


            for (int row = 0; row < 6; row++)
            {
                for (int col = 0; col < 7; col++)
                {
                    int index = row * 7 + col;
                    Button btnDay = new Button
                    {
                        Size = new Size(btnWidth, btnHeight),
                        Location = new Point(col * btnWidth, row * btnHeight),
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font("Segoe UI", 10F),
                        BackColor = Color.White,
                        Cursor = Cursors.Hand,
                        Margin = new Padding(0),
                        Padding = new Padding(0)
                    };
                    btnDay.FlatAppearance.BorderSize = 0;
                    btnDay.Click += BtnDay_Click;


                    dayButtons[index] = btnDay;
                    pnlDaysGrid.Controls.Add(btnDay);
                }
            }
        }

        private void AttachEvents()
        {
            btnPrevMonth.Click += delegate (object s, EventArgs e)
            {
                currentMonth--;
                if (currentMonth < 1) { currentMonth = 12; currentYear--; }
                BuildCalendar(currentYear, currentMonth);
            };

            btnNextMonth.Click += delegate (object s, EventArgs e)
            {
                currentMonth++;
                if (currentMonth > 12) { currentMonth = 1; currentYear++; }
                BuildCalendar(currentYear, currentMonth);
            };

            // Toggle Calendar Buttons
            btnCalInicio.Click += delegate (object s, EventArgs e)
            {

                txtFechaInicio.Focus();

                ToggleCalendar(txtFechaInicio);

            };
            btnCalFinal.Click += delegate (object s, EventArgs e)
            {

                txtFechaFinal.Focus();
                ToggleCalendar(txtFechaFinal);
            };

            // Focus and Visual Highlights

            txtFechaInicio.Enter += TextBox_Enter;
            txtFechaInicio.Leave += TextBox_Leave;
            txtFechaFinal.Enter += TextBox_Enter;
            txtFechaFinal.Leave += TextBox_Leave;

            // Date validation (KeyPress / Validating)
            txtFechaInicio.KeyPress += TextBox_KeyPress;
            txtFechaFinal.KeyPress += TextBox_KeyPress;

            txtFechaInicio.Validating += TextBox_ValidatingDate;
            txtFechaFinal.Validating += TextBox_ValidatingDate;

            cmbCDP.SelectedIndexChanged += CmbCDP_SelectedIndexChanged;


            btnImprimirPDF.Click += BtnImprimirPDF_Click;
            btnImprimirMasivo.Click += BtnImprimirMasivo_Click;
            btnImprimirCumplido.Click += BtnImprimirCumplido_Click;
            btnEditarDependencias.Click += BtnEditarDependencias_Click;
            btnEditarConfig.Click += BtnEditarConfig_Click;
            btnEditarDestinos.Click += BtnEditarDestinos_Click;
            btnEditarCategoria1.Click += BtnEditarCategoria1_Click;
        }

        private void TextBox_Enter(object sender, EventArgs e)
        {
            if (sender is TextBox tb)
            {
                tb.BackColor = Color.LightYellow; // Highlight color
                ToggleCalendar(tb);
            }
        }

        private void TextBox_Leave(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb != null)
            {
                tb.BackColor = SystemColors.Window; // Default color
            }
        }

        private void TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Only allow numbers, slash, and control characters (backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '/')
            {
                e.Handled = true;
            }
        }

        private void TextBox_ValidatingDate(object sender, System.ComponentModel.CancelEventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb != null && !string.IsNullOrWhiteSpace(tb.Text))
            {
                string[] validFormats = { "dd/MM/yyyy", "d/M/yyyy", "d/MM/yyyy", "dd/M/yyyy" };


                if (DateTime.TryParseExact(tb.Text, validFormats, null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                {
                    // Forzar el formato correcto
                    tb.Text = dt.ToString("dd/MM/yyyy");

                    // Validar la secuencia (Inicio vs Final)

                    ValidarSecuenciaFechas();
                }
                else
                {
                    MessageBox.Show("El formato de la fecha debe ser DD/MM/YYYY.", "Formato Inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tb.Text = ""; // Limpiar el campo en caso de error
                }
            }
        }

        private void ValidarSecuenciaFechas()
        {
            if (!string.IsNullOrWhiteSpace(txtFechaInicio.Text) && !string.IsNullOrWhiteSpace(txtFechaFinal.Text))
            {
                if (DateTime.TryParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dInicio) &&
                    DateTime.TryParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dFin))
                {
                    if (dInicio > dFin)
                    {
                        MessageBox.Show("La fecha de inicio no puede ser mayor que la final.", "Secuencia Inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtFechaFinal.Text = ""; // Limpiar la final para obligar a poner una válida
                    }
                }
            }
        }

        private void FormProgramador_Load(object sender, EventArgs e)
        {
            try
            {
                DateTime now = DateTime.Now;
                currentMonth = now.Month;
                currentYear = now.Year;


                BuildCalendar(currentYear, currentMonth);

                // Cargar datos desde excel

                var datos = _viaticosService.GetSelectionConfig();
                if (!string.IsNullOrEmpty(datos[0]) && datos[0] != "01/01/0001") txtFechaInicio.Text = datos[0];
                if (!string.IsNullOrEmpty(datos[1]) && datos[1] != "01/01/0001") txtFechaFinal.Text = datos[1];

                List<string> cdpList = _viaticosService.GetCdpList();

                cmbCDP.Items.Clear();
                foreach (var c in cdpList) cmbCDP.Items.Add(c);


                if (!string.IsNullOrEmpty(datos[2]) && cmbCDP.Items.Contains(datos[2]))
                {
                    cmbCDP.SelectedItem = datos[2];
                    lblCDPValue.Text = datos[2];
                }
                else if (cmbCDP.Items.Count > 0)
                {
                    cmbCDP.SelectedIndex = 0;
                }

                // Configurar visibilidad según el modo
                if (currentMode == TipoFormMode.verViaIndividual)
                {
                    btnImprimirPDF.Visible = true;
                    btnImprimirMasivo.Visible = false;
                    btnImprimirCumplido.Visible = false;
                }
                else if (currentMode == TipoFormMode.verViaMasivo)
                {
                    btnImprimirPDF.Visible = false;
                    btnImprimirMasivo.Visible = true;
                    btnImprimirCumplido.Visible = false;
                    this.Text = "PROCESAR VIATICOS MASIVOS";
                }
                else if (currentMode == TipoFormMode.verViaCumplidos)
                {
                    btnImprimirPDF.Visible = false;
                    btnImprimirMasivo.Visible = false;
                    btnImprimirCumplido.Visible = true;
                    this.Text = "GENERAR CUMPLIDOS E INFORMES";
                }

                // Verificar modo recurrente variable en excel
                string varRec = _viaticosService.GetGlobalVariable("Property_CrearREC_Id");

                if (!string.IsNullOrEmpty(varRec))
                {
                    // Modo recurrencia
                    btnImprimirPDF.Visible = false;
                    btnImprimirMasivo.Visible = false;
                    // btnEditarConfig solo visible
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error en la carga del FormProgramador", ex);
            }
        }

        private void BuildCalendar(int year, int month)
        {
            DateTime firstDayOfMonth = new DateTime(year, month, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // DOM=0, ..., SAB=6

            int startDayOfWeek = (int)firstDayOfMonth.DayOfWeek;


            string monthName = firstDayOfMonth.ToString("MMMM").ToUpper();
            lblMonthYear.Text = monthName + " " + year.ToString();

            DateTime prevMonthDate = firstDayOfMonth.AddMonths(-1);
            int daysInPrevMonth = DateTime.DaysInMonth(prevMonthDate.Year, prevMonthDate.Month);

            for (int i = 0; i < 42; i++)
            {
                Button btn = dayButtons[i];
                btn.Tag = null;

                if (i < startDayOfWeek)
                {
                    // Prev month days
                    int day = daysInPrevMonth - (startDayOfWeek - 1 - i);
                    btn.Text = day.ToString();
                    btn.ForeColor = Color.Gray;
                    btn.BackColor = Color.White;
                }
                else if (i >= startDayOfWeek + daysInMonth)
                {
                    // Next month days
                    int day = i - (startDayOfWeek + daysInMonth) + 1;
                    btn.Text = day.ToString();
                    btn.ForeColor = Color.Gray;
                    btn.BackColor = Color.White;
                }
                else
                {
                    // Current month days
                    int day = i - startDayOfWeek + 1;
                    btn.Text = day.ToString();
                    btn.ForeColor = Color.Black;
                    btn.BackColor = Color.White;
                    btn.Tag = new DateTime(year, month, day);


                    if (year == DateTime.Now.Year && month == DateTime.Now.Month && day == DateTime.Now.Day)
                    {
                        btn.BackColor = Color.LightBlue; // Highlight today
                    }
                }
            }
        }

        private void ToggleCalendar(TextBox target)
        {
            this.targetTextBox = target;
            if (DateTime.TryParse(target.Text, out DateTime dt))
            {
                currentMonth = dt.Month;
                currentYear = dt.Year;
            }
            else
            {
                currentMonth = DateTime.Now.Month;
                currentYear = DateTime.Now.Year;
            }


            BuildCalendar(currentYear, currentMonth);
        }

        private void BtnDay_Click(object sender, EventArgs e)
        {
            try
            {
                Button btn = sender as Button;
                if (btn != null && btn.Tag != null && targetTextBox != null)
                {
                    DateTime dt = (DateTime)btn.Tag;
                    targetTextBox.Text = dt.ToString("dd/MM/yyyy");

                    // Guardar auto a Excel

                    string cdps = cmbCDP.SelectedItem != null ? cmbCDP.SelectedItem.ToString() : "";
                    
                    DateTime? d1 = null;
                    DateTime? d2 = null;
                    if (DateTime.TryParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var t1)) d1 = t1;
                    if (DateTime.TryParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var t2)) d2 = t2;
                    
                    if (targetTextBox == txtFechaInicio) d1 = dt;
                    if (targetTextBox == txtFechaFinal) d2 = dt;

                    _viaticosService.FilaSeleccionadaSaveConfig(d1, d2, cdps);


                    targetTextBox = null; // deselect target
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al seleccionar día en el calendario", ex);
            }
        }

        private void CmbCDP_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbCDP.SelectedItem == null) return;
                string cdp = cmbCDP.SelectedItem.ToString();


                lblCDPValue.Text = cdp;


                var config = _viaticosService.GetConfig(null, null, cdp);
                lblCDPDependencia.Text = string.IsNullOrEmpty(config.CortaDependencia) ? "DEPENDENCIA" : config.CortaDependencia;


                string fullDesc = string.IsNullOrEmpty(config.FullDependencia) ? "DESCRIPCION" : config.FullDependencia;

                SelectedDependenciaDescripcion = fullDesc;

                // Truncate for UI display only

                lblCDPDependenciaDescripcion.Text = fullDesc.Length > 38 ? fullDesc.Substring(0, 38) : fullDesc;

                // Persistir en excel
                DateTime? d1 = null;
                DateTime? d2 = null;
                if (DateTime.TryParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var t1)) d1 = t1;
                if (DateTime.TryParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var t2)) d2 = t2;

                _viaticosService.FilaSeleccionadaSaveConfig(d1, d2, cdp);

            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al cambiar CDP seleccionado", ex);
            }
        }

        private void BtnImprimirPDF_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidarEntradas())
                {
                    SelectedFechaInicio = DateTime.ParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null);
                    SelectedFechaFin = DateTime.ParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null);
                    SelectedCDP = cmbCDP.SelectedItem?.ToString() ?? "";
                    // SelectedDependenciaDescripcion is already updated in CmbCDP_SelectedIndexChanged

                    // Guardar B1, B2 y B3 en Config
                    _viaticosService.FilaSeleccionadaSaveConfig(SelectedFechaInicio, SelectedFechaFin, SelectedCDP);



                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al procesar botón Imprimir PDF", ex);
                MessageBox.Show("Ocurrió un error inesperado. Consulte el registro de errores para más detalles.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImprimirMasivo_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidarEntradas())
                {
                    SelectedFechaInicio = DateTime.ParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null);
                    SelectedFechaFin = DateTime.ParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null);
                    SelectedCDP = cmbCDP.SelectedItem?.ToString() ?? "";

                    // Guardar B1, B2 y B3 en Config
                    _viaticosService.FilaSeleccionadaSaveConfig(SelectedFechaInicio, SelectedFechaFin, SelectedCDP);


                    this.Hide();

                    int vCount = 0;
                    using (var progress = new UI.ProgressManager("Generando Masivo Viáticos", 100))
                    {
                        vCount = _viaticosService.ProcesarMasivo(
                            (Action<int, int>)((curr, total) => progress.AddProgress(100.0 / total)));
                    }


                    if (vCount > 0)
                    {
                        MessageBox.Show(
                            $"Masivo generado exitosamente: {vCount} registros procesados.\nRevise su carpeta de Documentos.",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            "No se encontraron registros en el rango de fechas para generar el masivo.",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al procesar botón Imprimir Masivo", ex);
                MessageBox.Show("Ocurrió un error inesperado al generar el masivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImprimirCumplido_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidarEntradas())
                {
                    SelectedFechaInicio = DateTime.ParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null);
                    SelectedFechaFin = DateTime.ParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null);
                    SelectedCDP = cmbCDP.SelectedItem?.ToString() ?? "";

                    // Guardar B1, B2 y B3 en Config
                    _viaticosService.FilaSeleccionadaSaveConfig(SelectedFechaInicio, SelectedFechaFin, SelectedCDP);


                    this.Hide();

                    int vCount = 0;
                    using (var progress = new UI.ProgressManager("Generando Cumplidos e Informes", 100))
                    {
                        vCount = _viaticosService.ProcesarCumplidos(
                            (Action<int, int>)((curr, total) => progress.AddProgress(100.0 / total)),
                            (Action)(() => progress.Close()));
                    }


                    if (vCount > 0)
                    {
                        MessageBox.Show(
                            $"Cumplidos generados exitosamente: {vCount} registros procesados.\nRevise su carpeta de Documentos.",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            "No se encontraron registros en el rango de fechas para generar los cumplidos.",
                            App.APP_NAME, MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error al procesar botón Imprimir Cumplidos", ex);
                MessageBox.Show("Ocurrió un error inesperado al generar los cumplidos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Show();
            }
        }

        private void BtnEditarConfig_Click(object sender, EventArgs e)
        {
            using (var frm = new FormConfig(_viaticosService))
            {
                frm.ShowDialog();
            }
        }


        private void BtnEditarDependencias_Click(object sender, EventArgs e)
        {
            using (var frm = new FormTableEditor("tblDependencias", "GESTION DE DEPENDENCIAS", WIDE_DEPENDENCIAS, _viaticosService))
            {
                frm.ShowDialog();
            }
        }


        private void BtnEditarDestinos_Click(object sender, EventArgs e)
        {
            using (var frm = new FormTableEditor("tblDestinos", "GESTION DE DESTINOS", WIDE_DESTINOS, _viaticosService))
            {
                frm.ShowDialog();
            }
        }


        private void BtnEditarCategoria1_Click(object sender, EventArgs e)
        {
            using (var frm = new FormTableEditor("tblCategoria1", "GESTION DE MUNICIPIOS CATEGORIA UNO", WIDE_CATEGORIA1, _viaticosService))
            {
                frm.ShowDialog();
            }
        }


        private bool ValidarEntradas()
        {
            if (!DateTime.TryParseExact(txtFechaInicio.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime d1))
            {
                MessageBox.Show("Fecha de inicio no es válida. Use DD/MM/YYYY");
                return false;
            }
            if (!DateTime.TryParseExact(txtFechaFinal.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime d2))
            {
                MessageBox.Show("Fecha final no es válida. Use DD/MM/YYYY");
                return false;
            }
            if (d1 > d2)
            {
                MessageBox.Show("La fecha de inicio no puede ser mayor que la final.");
                return false;
            }
            string cdp = cmbCDP.SelectedItem != null ? cmbCDP.SelectedItem.ToString() : "";
            if (!long.TryParse(cdp, out _))
            {
                MessageBox.Show("El CDP debe ser numérico.");
                return false;
            }
            return true;
        }
    }
}