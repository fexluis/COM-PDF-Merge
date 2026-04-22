namespace ConversorPDF.XLL.FormAcademic
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
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.cmbSedes = new System.Windows.Forms.ComboBox();
            this.cmbProgramas = new System.Windows.Forms.ComboBox();
            this.cmbSemestres = new System.Windows.Forms.ComboBox();
            this.cmbAsignaturas = new System.Windows.Forms.ComboBox();
            this.cmbTerceros = new System.Windows.Forms.ComboBox();
            this.txtTotalHoras = new System.Windows.Forms.TextBox();
            this.txtDiasHoras = new ConversorPDF.XLL.UI.DayHoursDropdown();
            this.dtpInicio = new System.Windows.Forms.DateTimePicker();
            this.dtpFinal = new System.Windows.Forms.DateTimePicker();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
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
            ((System.ComponentModel.ISupportInitialize)(this.dgvSesiones)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbSedes
            // 
            this.cmbSedes.Location = new System.Drawing.Point(20, 40);
            this.cmbSedes.Size = new System.Drawing.Size(200, 23);
            this.cmbSedes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbProgramas
            // 
            this.cmbProgramas.Location = new System.Drawing.Point(230, 40);
            this.cmbProgramas.Size = new System.Drawing.Size(200, 23);
            this.cmbProgramas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbProgramas.SelectedIndexChanged += new System.EventHandler(this.cmbProgramas_SelectedIndexChanged);
            // 
            // cmbSemestres
            // 
            this.cmbSemestres.Location = new System.Drawing.Point(440, 40);
            this.cmbSemestres.Size = new System.Drawing.Size(100, 23);
            this.cmbSemestres.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestres.SelectedIndexChanged += new System.EventHandler(this.cmbSemestres_SelectedIndexChanged);
            // 
            // cmbAsignaturas
            // 
            this.cmbAsignaturas.Location = new System.Drawing.Point(550, 40);
            this.cmbAsignaturas.Size = new System.Drawing.Size(250, 23);
            this.cmbAsignaturas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // cmbTerceros
            // 
            this.cmbTerceros.Location = new System.Drawing.Point(20, 90);
            this.cmbTerceros.Size = new System.Drawing.Size(300, 23);
            this.cmbTerceros.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // 
            // txtTotalHoras
            // 
            this.txtTotalHoras.Location = new System.Drawing.Point(340, 90);
            this.txtTotalHoras.Size = new System.Drawing.Size(50, 23);
            // 
            // txtDiasHoras
            // 
            this.txtDiasHoras.Location = new System.Drawing.Point(410, 90);
            this.txtDiasHoras.Size = new System.Drawing.Size(390, 23);
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
            // btnGenerar
            // 
            this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            this.btnGenerar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerar.FlatAppearance.BorderSize = 0;
            this.btnGenerar.ForeColor = System.Drawing.Color.White;
            this.btnGenerar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenerar.Location = new System.Drawing.Point(300, 137);
            this.btnGenerar.Size = new System.Drawing.Size(240, 30);
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
            this.btnGuardar.Location = new System.Drawing.Point(550, 137);
            this.btnGuardar.Size = new System.Drawing.Size(250, 30);
            this.btnGuardar.Text = "GUARDAR SESIONES";
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // dgvSesiones
            // 
            this.dgvSesiones.Location = new System.Drawing.Point(20, 190);
            this.dgvSesiones.Size = new System.Drawing.Size(780, 240);
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
            this.pgbGuardar.Size = new System.Drawing.Size(780, 8);
            this.pgbGuardar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.pgbGuardar.MarqueeAnimationSpeed = 30;
            this.pgbGuardar.Visible = false;
            // 
            // Labels...
            this.label1.Text = "Sede:"; this.label1.Location = new System.Drawing.Point(20, 20);
            this.label2.Text = "Programa:"; this.label2.Location = new System.Drawing.Point(230, 20);
            this.label3.Text = "Semestre:"; this.label3.Location = new System.Drawing.Point(440, 20);
            this.label4.Text = "Asignatura:"; this.label4.Location = new System.Drawing.Point(550, 20);
            this.label5.Text = "Tercero:"; this.label5.Location = new System.Drawing.Point(20, 70);
            this.label6.Text = "Total Horas:"; this.label6.Location = new System.Drawing.Point(340, 70);
            this.label8.Text = "Fecha Inicio:"; this.label8.Location = new System.Drawing.Point(20, 120);
            this.label9.Text = "Fecha Final:"; this.label9.Location = new System.Drawing.Point(160, 120);

            // Form
            this.ClientSize = new System.Drawing.Size(820, 480);
            this.Controls.Add(this.cmbSedes); this.Controls.Add(this.cmbProgramas); this.Controls.Add(this.cmbSemestres);
            this.Controls.Add(this.cmbAsignaturas); this.Controls.Add(this.cmbTerceros); this.Controls.Add(this.txtTotalHoras);
            this.Controls.Add(this.txtDiasHoras); this.Controls.Add(this.dtpInicio); this.Controls.Add(this.dtpFinal);
            this.Controls.Add(this.btnGenerar); this.Controls.Add(this.btnGuardar);
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

            ((System.ComponentModel.ISupportInitialize)(this.dgvSesiones)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.ComboBox cmbSedes;
        private System.Windows.Forms.ComboBox cmbProgramas;
        private System.Windows.Forms.ComboBox cmbSemestres;
        private System.Windows.Forms.ComboBox cmbAsignaturas;
        private System.Windows.Forms.ComboBox cmbTerceros;
        private System.Windows.Forms.TextBox txtTotalHoras;
        private ConversorPDF.XLL.UI.DayHoursDropdown txtDiasHoras;
        private System.Windows.Forms.DateTimePicker dtpInicio;
        private System.Windows.Forms.DateTimePicker dtpFinal;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.DataGridView dgvSesiones;
        private System.Windows.Forms.Label lblSelectionSummary;
        private System.Windows.Forms.ProgressBar pgbGuardar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
    }
}
