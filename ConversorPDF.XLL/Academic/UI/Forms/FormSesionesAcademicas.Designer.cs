namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormSesionesAcademicas
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            if (disposing)
            {
                _sharedConn?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.cmbPeriodos = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cmbSedes = new System.Windows.Forms.ComboBox();
            this.cmbProgramas = new System.Windows.Forms.ComboBox();
            this.cmbSemestres = new System.Windows.Forms.ComboBox();
            this.cmbAsignaturas = new System.Windows.Forms.ComboBox();
            this.cmbTipoTutoria = new System.Windows.Forms.ComboBox();
            this.cmbTerceros = new System.Windows.Forms.ComboBox();
            this.cmbGrupos = new System.Windows.Forms.ComboBox();
            this.txtTotalHoras = new System.Windows.Forms.TextBox();
            this.lblIntHoraria = new System.Windows.Forms.Label();
            this.txtDiasHoras = new ConversorPDF.XLL.UI.DayHoursDropdown();
            this.dtpInicio = new System.Windows.Forms.DateTimePicker();
            this.dtpFinal = new System.Windows.Forms.DateTimePicker();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnToggleViatica = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.dgvSesiones = new System.Windows.Forms.DataGridView();
            this.lblSelectionSummary = new System.Windows.Forms.Label();
            this.pgbGuardar = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblGrupo = new System.Windows.Forms.Label();
            this.lblTipoTutoria = new System.Windows.Forms.Label();

            // Controles Terminal Aéreo y Pernocta
            this.chkTerminalAereo = new System.Windows.Forms.CheckBox();
            this.chkPernoctaUltimoDia = new System.Windows.Forms.CheckBox();
            this.pnlTransporteAereo = new System.Windows.Forms.Panel();
            this.lblTerminalAereo = new System.Windows.Forms.Label();
            this.cmbTerminalesAereas = new System.Windows.Forms.ComboBox();
            this.lblGastoTerminalAerea = new System.Windows.Forms.Label();
            this.txtGastoTerminalAerea = new System.Windows.Forms.TextBox();
            this.lblMaxValor = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(this.dgvSesiones)).BeginInit();
            this.pnlTransporteAereo.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbPeriodos
            // 
            this.cmbPeriodos.Location = new System.Drawing.Point(20, 40);
            this.cmbPeriodos.Size = new System.Drawing.Size(80, 23);
            this.cmbPeriodos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbSedes
            // 
            this.cmbSedes.Location = new System.Drawing.Point(110, 40);
            this.cmbSedes.Size = new System.Drawing.Size(200, 23);
            this.cmbSedes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbProgramas
            // 
            this.cmbProgramas.Location = new System.Drawing.Point(320, 40);
            this.cmbProgramas.Size = new System.Drawing.Size(130, 23);
            this.cmbProgramas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbProgramas.SelectedIndexChanged += new System.EventHandler(this.cmbProgramas_SelectedIndexChanged);
            // 
            // cmbSemestres
            // 
            this.cmbSemestres.Location = new System.Drawing.Point(460, 40);
            this.cmbSemestres.Size = new System.Drawing.Size(60, 23);
            this.cmbSemestres.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestres.SelectedIndexChanged += new System.EventHandler(this.cmbSemestres_SelectedIndexChanged);
            // 
            // cmbGrupos
            // 
            this.cmbGrupos.Location = new System.Drawing.Point(530, 40);
            this.cmbGrupos.Size = new System.Drawing.Size(60, 23);
            this.cmbGrupos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbAsignaturas
            // 
            this.cmbAsignaturas.Location = new System.Drawing.Point(600, 40);
            this.cmbAsignaturas.Size = new System.Drawing.Size(170, 23);
            this.cmbAsignaturas.SelectedIndexChanged += new System.EventHandler(this.cmbAsignaturas_SelectedIndexChanged);
            this.cmbAsignaturas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbTipoTutoria
            // 
            this.cmbTipoTutoria.Location = new System.Drawing.Point(780, 40);
            this.cmbTipoTutoria.Size = new System.Drawing.Size(120, 23);
            this.cmbTipoTutoria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbTerceros
            // 
            this.cmbTerceros.Location = new System.Drawing.Point(20, 90);
            this.cmbTerceros.Size = new System.Drawing.Size(400, 23);
            this.cmbTerceros.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            this.cmbTerceros.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbTerceros.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            // 
            // txtTotalHoras
            // 
            this.txtTotalHoras.Location = new System.Drawing.Point(430, 90);
            this.txtTotalHoras.Size = new System.Drawing.Size(50, 23);

            // lblIntHoraria
            this.lblIntHoraria.AutoSize = true;
            this.lblIntHoraria.Location = new System.Drawing.Point(470, 70);
            this.lblIntHoraria.Name = "lblIntHoraria";
            this.lblIntHoraria.Size = new System.Drawing.Size(60, 15);
            this.lblIntHoraria.Text = "";
            this.lblIntHoraria.ForeColor = System.Drawing.Color.DimGray;
            // 
            // txtDiasHoras
            // 
            this.txtDiasHoras.Location = new System.Drawing.Point(490, 90);
            this.txtDiasHoras.Size = new System.Drawing.Size(410, 23);
            this.txtDiasHoras.PlaceholderText = "Seleccionar días y horas...";
            // 
            // dtpInicio
            // 
            this.dtpInicio.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpInicio.Location = new System.Drawing.Point(20, 140);
            this.dtpInicio.Size = new System.Drawing.Size(120, 23);
            // 
            // dtpFinal
            // 
            this.dtpFinal.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFinal.Location = new System.Drawing.Point(160, 140);
            this.dtpFinal.Size = new System.Drawing.Size(120, 23);
            // 
            // chkTerminalAereo
            // 
            this.chkTerminalAereo.AutoSize = true;
            this.chkTerminalAereo.Location = new System.Drawing.Point(290, 143);
            this.chkTerminalAereo.Name = "chkTerminalAereo";
            this.chkTerminalAereo.Size = new System.Drawing.Size(60, 19);
            this.chkTerminalAereo.TabIndex = 22;
            this.chkTerminalAereo.Text = "Aéreo";
            this.chkTerminalAereo.UseVisualStyleBackColor = true;
            this.chkTerminalAereo.CheckedChanged += new System.EventHandler(this.chkTerminalAereo_CheckedChanged);
            // 
            // chkPernoctaUltimoDia
            // 
            this.chkPernoctaUltimoDia.AutoSize = true;
            this.chkPernoctaUltimoDia.Location = new System.Drawing.Point(356, 143);
            this.chkPernoctaUltimoDia.Name = "chkPernoctaUltimoDia";
            this.chkPernoctaUltimoDia.Size = new System.Drawing.Size(135, 19);
            this.chkPernoctaUltimoDia.TabIndex = 23;
            this.chkPernoctaUltimoDia.Text = "Pernocta último día";
            this.chkPernoctaUltimoDia.UseVisualStyleBackColor = true;
            // 
            // btnGenerar
            // 
            this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            this.btnGenerar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerar.FlatAppearance.BorderSize = 0;
            this.btnGenerar.ForeColor = System.Drawing.Color.White;
            this.btnGenerar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenerar.Location = new System.Drawing.Point(500, 137);
            this.btnGenerar.Size = new System.Drawing.Size(140, 30);
            this.btnGenerar.Text = "GENERAR SESIONES";
            this.btnGenerar.UseVisualStyleBackColor = false;
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.FlatAppearance.BorderSize = 0;
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGuardar.Location = new System.Drawing.Point(650, 137);
            this.btnGuardar.Size = new System.Drawing.Size(150, 30);
            this.btnGuardar.Text = "GUARDAR SESIONES";
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnToggleViatica
            // 
            this.btnToggleViatica.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            this.btnToggleViatica.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToggleViatica.FlatAppearance.BorderSize = 0;
            this.btnToggleViatica.ForeColor = System.Drawing.Color.White;
            this.btnToggleViatica.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnToggleViatica.Location = new System.Drawing.Point(810, 137);
            this.btnToggleViatica.Size = new System.Drawing.Size(40, 30);
            this.btnToggleViatica.Text = "V";
            this.btnToggleViatica.UseVisualStyleBackColor = false;
            this.btnToggleViatica.Click += new System.EventHandler(this.btnToggleViatica_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.BackColor = System.Drawing.Color.FromArgb(220, 53, 69);
            this.btnEliminar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEliminar.FlatAppearance.BorderSize = 0;
            this.btnEliminar.ForeColor = System.Drawing.Color.White;
            this.btnEliminar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnEliminar.Location = new System.Drawing.Point(860, 137);
            this.btnEliminar.Size = new System.Drawing.Size(40, 30);
            this.btnEliminar.Text = "E";
            this.btnEliminar.UseVisualStyleBackColor = false;
            this.btnEliminar.Visible = false;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // pnlTransporteAereo
            // 
            this.pnlTransporteAereo.Controls.Add(this.lblTerminalAereo);
            this.pnlTransporteAereo.Controls.Add(this.cmbTerminalesAereas);
            this.pnlTransporteAereo.Controls.Add(this.lblGastoTerminalAerea);
            this.pnlTransporteAereo.Controls.Add(this.txtGastoTerminalAerea);
            this.pnlTransporteAereo.Controls.Add(this.lblMaxValor);
            this.pnlTransporteAereo.Location = new System.Drawing.Point(20, 168);
            this.pnlTransporteAereo.Name = "pnlTransporteAereo";
            this.pnlTransporteAereo.Size = new System.Drawing.Size(880, 30);
            this.pnlTransporteAereo.TabIndex = 23;
            this.pnlTransporteAereo.Visible = false;
            // 
            // lblTerminalAereo
            // 
            this.lblTerminalAereo.AutoSize = true;
            this.lblTerminalAereo.Location = new System.Drawing.Point(0, 5);
            this.lblTerminalAereo.Name = "lblTerminalAereo";
            this.lblTerminalAereo.Size = new System.Drawing.Size(91, 15);
            this.lblTerminalAereo.Text = "Terminal Aéreo:";
            // 
            // cmbTerminalesAereas
            // 
            this.cmbTerminalesAereas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTerminalesAereas.Location = new System.Drawing.Point(100, 2);
            this.cmbTerminalesAereas.Name = "cmbTerminalesAereas";
            this.cmbTerminalesAereas.Size = new System.Drawing.Size(200, 23);
            this.cmbTerminalesAereas.SelectedIndexChanged += new System.EventHandler(this.cmbTerminalesAereas_SelectedIndexChanged);
            // 
            // lblGastoTerminalAerea
            // 
            this.lblGastoTerminalAerea.AutoSize = true;
            this.lblGastoTerminalAerea.Location = new System.Drawing.Point(310, 5);
            this.lblGastoTerminalAerea.Name = "lblGastoTerminalAerea";
            this.lblGastoTerminalAerea.Size = new System.Drawing.Size(126, 15);
            this.lblGastoTerminalAerea.Text = "Gasto Terminal aérea:";
            // 
            // txtGastoTerminalAerea
            // 
            this.txtGastoTerminalAerea.Location = new System.Drawing.Point(440, 2);
            this.txtGastoTerminalAerea.Name = "txtGastoTerminalAerea";
            this.txtGastoTerminalAerea.Size = new System.Drawing.Size(100, 23);
            // 
            // lblMaxValor
            // 
            this.lblMaxValor.AutoSize = true;
            this.lblMaxValor.ForeColor = System.Drawing.Color.DimGray;
            this.lblMaxValor.Location = new System.Drawing.Point(550, 5);
            this.lblMaxValor.Name = "lblMaxValor";
            this.lblMaxValor.Size = new System.Drawing.Size(200, 15);
            this.lblMaxValor.Text = "";
            // 
            // dgvSesiones
            // 
            this.dgvSesiones.Location = new System.Drawing.Point(20, 195);
            this.dgvSesiones.Size = new System.Drawing.Size(880, 235);
            this.dgvSesiones.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSesiones_CellValueChanged);
            this.dgvSesiones.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgvSesiones_CurrentCellDirtyStateChanged);
            // 
            // lblSelectionSummary
            // 
            this.lblSelectionSummary.Location = new System.Drawing.Point(20, 438);
            this.lblSelectionSummary.Size = new System.Drawing.Size(400, 20);
            this.lblSelectionSummary.Text = "0 sesiones seleccionadas...";
            // 
            // pgbGuardar
            // 
            this.pgbGuardar.Location = new System.Drawing.Point(20, 462);
            this.pgbGuardar.Size = new System.Drawing.Size(880, 8);
            this.pgbGuardar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.pgbGuardar.MarqueeAnimationSpeed = 30;
            this.pgbGuardar.Visible = false;
            // 
            // Labels...
            this.label10.Text = "Período:"; this.label10.Location = new System.Drawing.Point(20, 20); this.label10.AutoSize = true;
            this.label1.Text = "Sede:"; this.label1.Location = new System.Drawing.Point(110, 20); this.label1.AutoSize = true;
            this.label2.Text = "Programa:"; this.label2.Location = new System.Drawing.Point(320, 20); this.label2.AutoSize = true;
            this.label3.Text = "Semestre:"; this.label3.Location = new System.Drawing.Point(460, 20); this.label3.AutoSize = true;
            this.label4.Text = "Asignatura:"; this.label4.Location = new System.Drawing.Point(600, 20); this.label4.AutoSize = true;
            this.label5.Text = "Tercero:"; this.label5.Location = new System.Drawing.Point(20, 70); this.label5.AutoSize = true;
            this.label6.Text = "Horas:"; this.label6.Location = new System.Drawing.Point(430, 70); this.label6.AutoSize = true;
            this.label8.Text = "Fecha Inicio:"; this.label8.Location = new System.Drawing.Point(20, 120); this.label8.AutoSize = true;
            this.label9.Text = "Fecha Final:"; this.label9.Location = new System.Drawing.Point(160, 120); this.label9.AutoSize = true;
            this.lblGrupo.Text = "Grupo:"; this.lblGrupo.Location = new System.Drawing.Point(530, 20); this.lblGrupo.AutoSize = true;
            this.lblTipoTutoria.Text = "Tipo Tutoría:"; this.lblTipoTutoria.Location = new System.Drawing.Point(780, 20); this.lblTipoTutoria.AutoSize = true;

            // 
            // pnlLoading
            // 
            this.pnlLoading = new System.Windows.Forms.Panel();
            this.lblLoading = new System.Windows.Forms.Label();
            this.pgbLoading = new System.Windows.Forms.ProgressBar();
            this.pnlLoading.SuspendLayout();
            this.pnlLoading.BackColor = System.Drawing.Color.White;
            this.pnlLoading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlLoading.Location = new System.Drawing.Point(0, 0);
            this.pnlLoading.Name = "pnlLoading";
            this.pnlLoading.Size = new System.Drawing.Size(920, 480);
            this.pnlLoading.TabIndex = 22;
            // 
            // lblLoading
            // 
            this.lblLoading.AutoSize = false;
            this.lblLoading.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold);
            this.lblLoading.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(212)))));
            this.lblLoading.Location = new System.Drawing.Point(0, 200);
            this.lblLoading.Name = "lblLoading";
            this.lblLoading.Size = new System.Drawing.Size(920, 30);
            this.lblLoading.TabIndex = 0;
            this.lblLoading.Text = "Cargando componentes y datos seguros...";
            this.lblLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pgbLoading
            // 
            this.pgbLoading.Location = new System.Drawing.Point(310, 240);
            this.pgbLoading.MarqueeAnimationSpeed = 30;
            this.pgbLoading.Name = "pgbLoading";
            this.pgbLoading.Size = new System.Drawing.Size(300, 15);
            this.pgbLoading.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.pgbLoading.TabIndex = 1;

            this.pnlLoading.Controls.Add(this.lblLoading);
            this.pnlLoading.Controls.Add(this.pgbLoading);

            // Form
            this.ClientSize = new System.Drawing.Size(920, 480);
            this.Controls.Add(this.pnlLoading);
            this.Controls.Add(this.cmbPeriodos); this.Controls.Add(this.label10);
            this.Controls.Add(this.cmbSedes); this.Controls.Add(this.cmbProgramas); this.Controls.Add(this.cmbSemestres);
            this.Controls.Add(this.cmbGrupos); this.Controls.Add(this.lblGrupo);
            this.Controls.Add(this.cmbAsignaturas); this.Controls.Add(this.cmbTipoTutoria); this.Controls.Add(this.lblTipoTutoria); this.Controls.Add(this.cmbTerceros); this.Controls.Add(this.txtTotalHoras);
            this.Controls.Add(this.txtDiasHoras); this.Controls.Add(this.lblIntHoraria); this.Controls.Add(this.dtpInicio); this.Controls.Add(this.dtpFinal);
            this.Controls.Add(this.btnGenerar); this.Controls.Add(this.btnGuardar); this.Controls.Add(this.btnToggleViatica); this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.chkTerminalAereo);
            this.Controls.Add(this.chkPernoctaUltimoDia);
            this.Controls.Add(this.pnlTransporteAereo);
            this.Controls.Add(this.dgvSesiones);
            this.Controls.Add(this.lblSelectionSummary); this.Controls.Add(this.pgbGuardar);
            this.Controls.Add(this.label1); this.Controls.Add(this.label2); this.Controls.Add(this.label3);
            this.Controls.Add(this.label4); this.Controls.Add(this.label5); this.Controls.Add(this.label6);
            this.Controls.Add(this.label8); this.Controls.Add(this.label9);
            this.Text = "Sesiones Académicas";
            this.BackColor = System.Drawing.Color.White;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MaximizeBox = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;

            this.pnlTransporteAereo.ResumeLayout(false);
            this.pnlTransporteAereo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSesiones)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
            this.pnlTransporteAereo.BringToFront();
        }

        private System.Windows.Forms.ComboBox cmbPeriodos;
        private System.Windows.Forms.ComboBox cmbSedes;
        private System.Windows.Forms.ComboBox cmbProgramas;
        private System.Windows.Forms.ComboBox cmbSemestres;
        private System.Windows.Forms.ComboBox cmbAsignaturas;
        private System.Windows.Forms.ComboBox cmbTipoTutoria;
        private System.Windows.Forms.ComboBox cmbTerceros;
        private System.Windows.Forms.TextBox txtTotalHoras;
        private System.Windows.Forms.Label lblIntHoraria;
        private ConversorPDF.XLL.UI.DayHoursDropdown txtDiasHoras;
        private System.Windows.Forms.DateTimePicker dtpInicio;
        private System.Windows.Forms.DateTimePicker dtpFinal;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnToggleViatica;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.DataGridView dgvSesiones;
        private System.Windows.Forms.Label lblSelectionSummary;
        private System.Windows.Forms.ProgressBar pgbGuardar;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbGrupos;
        private System.Windows.Forms.Label lblGrupo;
        private System.Windows.Forms.Label lblTipoTutoria;
        private System.Windows.Forms.Panel pnlLoading;
        private System.Windows.Forms.Label lblLoading;
        private System.Windows.Forms.ProgressBar pgbLoading;

        // Terminal Aéreo y Pernocta
        private System.Windows.Forms.CheckBox chkTerminalAereo;
        private System.Windows.Forms.CheckBox chkPernoctaUltimoDia;
        private System.Windows.Forms.Panel pnlTransporteAereo;
        private System.Windows.Forms.Label lblTerminalAereo;
        private System.Windows.Forms.ComboBox cmbTerminalesAereas;
        private System.Windows.Forms.Label lblGastoTerminalAerea;
        private System.Windows.Forms.TextBox txtGastoTerminalAerea;
        private System.Windows.Forms.Label lblMaxValor;
    }
}
