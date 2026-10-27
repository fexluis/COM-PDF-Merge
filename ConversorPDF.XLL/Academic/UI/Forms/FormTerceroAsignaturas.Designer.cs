namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormTerceroAsignaturas
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

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblPrograma = new System.Windows.Forms.Label();
            this.cmbPrograma = new System.Windows.Forms.ComboBox();
            this.lblSemestre = new System.Windows.Forms.Label();
            this.cmbSemestre = new System.Windows.Forms.ComboBox();
            this.lblAsignatura = new System.Windows.Forms.Label();
            this.cmbAsignatura = new System.Windows.Forms.ComboBox();
            this.chkActivo = new System.Windows.Forms.CheckBox();
            this.lblPrioridad = new System.Windows.Forms.Label();
            this.numPrioridad = new System.Windows.Forms.NumericUpDown();
            this.lblVecesAsignadaTitulo = new System.Windows.Forms.Label();
            this.lblVecesAsignadaValor = new System.Windows.Forms.Label();
            this.lblUltimaAsignacionTitulo = new System.Windows.Forms.Label();
            this.lblUltimaAsignacionValor = new System.Windows.Forms.Label();
            this.lblOrigen = new System.Windows.Forms.Label();
            this.txtOrigen = new System.Windows.Forms.TextBox();
            this.lblObservaciones = new System.Windows.Forms.Label();
            this.txtObservaciones = new System.Windows.Forms.TextBox();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numPrioridad)).BeginInit();
            this.SuspendLayout();
            // 
            // lblPrograma
            // 
            this.lblPrograma.AutoSize = true;
            this.lblPrograma.Location = new System.Drawing.Point(20, 20);
            this.lblPrograma.Name = "lblPrograma";
            this.lblPrograma.Size = new System.Drawing.Size(63, 15);
            this.lblPrograma.TabIndex = 0;
            this.lblPrograma.Text = "Programa:";
            // 
            // cmbPrograma
            // 
            this.cmbPrograma.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPrograma.FormattingEnabled = true;
            this.cmbPrograma.Location = new System.Drawing.Point(120, 17);
            this.cmbPrograma.Name = "cmbPrograma";
            this.cmbPrograma.Size = new System.Drawing.Size(340, 23);
            this.cmbPrograma.TabIndex = 1;
            this.cmbPrograma.SelectedIndexChanged += new System.EventHandler(this.CmbPrograma_SelectedIndexChanged);
            // 
            // lblSemestre
            // 
            this.lblSemestre.AutoSize = true;
            this.lblSemestre.Location = new System.Drawing.Point(20, 55);
            this.lblSemestre.Name = "lblSemestre";
            this.lblSemestre.Size = new System.Drawing.Size(61, 15);
            this.lblSemestre.TabIndex = 2;
            this.lblSemestre.Text = "Semestre:";
            // 
            // cmbSemestre
            // 
            this.cmbSemestre.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestre.FormattingEnabled = true;
            this.cmbSemestre.Location = new System.Drawing.Point(120, 52);
            this.cmbSemestre.Name = "cmbSemestre";
            this.cmbSemestre.Size = new System.Drawing.Size(150, 23);
            this.cmbSemestre.TabIndex = 3;
            this.cmbSemestre.SelectedIndexChanged += new System.EventHandler(this.CmbSemestre_SelectedIndexChanged);
            // 
            // lblAsignatura
            // 
            this.lblAsignatura.AutoSize = true;
            this.lblAsignatura.Location = new System.Drawing.Point(20, 90);
            this.lblAsignatura.Name = "lblAsignatura";
            this.lblAsignatura.Size = new System.Drawing.Size(68, 15);
            this.lblAsignatura.TabIndex = 4;
            this.lblAsignatura.Text = "Asignatura:";
            // 
            // cmbAsignatura
            // 
            this.cmbAsignatura.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAsignatura.FormattingEnabled = true;
            this.cmbAsignatura.Location = new System.Drawing.Point(120, 87);
            this.cmbAsignatura.Name = "cmbAsignatura";
            this.cmbAsignatura.Size = new System.Drawing.Size(340, 23);
            this.cmbAsignatura.TabIndex = 5;
            // 
            // chkActivo
            // 
            this.chkActivo.AutoSize = true;
            this.chkActivo.Checked = true;
            this.chkActivo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkActivo.Location = new System.Drawing.Point(120, 125);
            this.chkActivo.Name = "chkActivo";
            this.chkActivo.Size = new System.Drawing.Size(126, 19);
            this.chkActivo.TabIndex = 6;
            this.chkActivo.Text = "Habilitación Activa";
            this.chkActivo.UseVisualStyleBackColor = true;
            // 
            // lblPrioridad
            // 
            this.lblPrioridad.AutoSize = true;
            this.lblPrioridad.Location = new System.Drawing.Point(270, 126);
            this.lblPrioridad.Name = "lblPrioridad";
            this.lblPrioridad.Size = new System.Drawing.Size(61, 15);
            this.lblPrioridad.TabIndex = 7;
            this.lblPrioridad.Text = "Prioridad:";
            // 
            // numPrioridad
            // 
            this.numPrioridad.Location = new System.Drawing.Point(345, 124);
            this.numPrioridad.Maximum = new decimal(new int[] { 100, 0, 0, 0 });
            this.numPrioridad.Name = "numPrioridad";
            this.numPrioridad.Size = new System.Drawing.Size(60, 23);
            this.numPrioridad.TabIndex = 8;
            // 
            // lblVecesAsignadaTitulo
            // 
            this.lblVecesAsignadaTitulo.AutoSize = true;
            this.lblVecesAsignadaTitulo.Location = new System.Drawing.Point(20, 160);
            this.lblVecesAsignadaTitulo.Name = "lblVecesAsignadaTitulo";
            this.lblVecesAsignadaTitulo.Size = new System.Drawing.Size(95, 15);
            this.lblVecesAsignadaTitulo.TabIndex = 9;
            this.lblVecesAsignadaTitulo.Text = "Veces Asignada:";
            // 
            // lblVecesAsignadaValor
            // 
            this.lblVecesAsignadaValor.AutoSize = true;
            this.lblVecesAsignadaValor.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblVecesAsignadaValor.Location = new System.Drawing.Point(120, 160);
            this.lblVecesAsignadaValor.Name = "lblVecesAsignadaValor";
            this.lblVecesAsignadaValor.Size = new System.Drawing.Size(14, 15);
            this.lblVecesAsignadaValor.TabIndex = 10;
            this.lblVecesAsignadaValor.Text = "0";
            // 
            // lblUltimaAsignacionTitulo
            // 
            this.lblUltimaAsignacionTitulo.AutoSize = true;
            this.lblUltimaAsignacionTitulo.Location = new System.Drawing.Point(230, 160);
            this.lblUltimaAsignacionTitulo.Name = "lblUltimaAsignacionTitulo";
            this.lblUltimaAsignacionTitulo.Size = new System.Drawing.Size(110, 15);
            this.lblUltimaAsignacionTitulo.TabIndex = 11;
            this.lblUltimaAsignacionTitulo.Text = "Última Asignación:";
            // 
            // lblUltimaAsignacionValor
            // 
            this.lblUltimaAsignacionValor.AutoSize = true;
            this.lblUltimaAsignacionValor.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblUltimaAsignacionValor.Location = new System.Drawing.Point(345, 160);
            this.lblUltimaAsignacionValor.Name = "lblUltimaAsignacionValor";
            this.lblUltimaAsignacionValor.Size = new System.Drawing.Size(22, 15);
            this.lblUltimaAsignacionValor.TabIndex = 12;
            this.lblUltimaAsignacionValor.Text = "N/A";
            // 
            // lblOrigen
            // 
            this.lblOrigen.AutoSize = true;
            this.lblOrigen.Location = new System.Drawing.Point(20, 195);
            this.lblOrigen.Name = "lblOrigen";
            this.lblOrigen.Size = new System.Drawing.Size(46, 15);
            this.lblOrigen.TabIndex = 13;
            this.lblOrigen.Text = "Origen:";
            // 
            // txtOrigen
            // 
            this.txtOrigen.Location = new System.Drawing.Point(120, 192);
            this.txtOrigen.Name = "txtOrigen";
            this.txtOrigen.ReadOnly = true;
            this.txtOrigen.Size = new System.Drawing.Size(150, 23);
            this.txtOrigen.TabIndex = 14;
            this.txtOrigen.Text = "MANUAL";
            // 
            // lblObservaciones
            // 
            this.lblObservaciones.AutoSize = true;
            this.lblObservaciones.Location = new System.Drawing.Point(20, 230);
            this.lblObservaciones.Name = "lblObservaciones";
            this.lblObservaciones.Size = new System.Drawing.Size(89, 15);
            this.lblObservaciones.TabIndex = 15;
            this.lblObservaciones.Text = "Observaciones:";
            // 
            // txtObservaciones
            // 
            this.txtObservaciones.Location = new System.Drawing.Point(120, 227);
            this.txtObservaciones.Multiline = true;
            this.txtObservaciones.Name = "txtObservaciones";
            this.txtObservaciones.Size = new System.Drawing.Size(340, 70);
            this.txtObservaciones.TabIndex = 16;
            // 
            // btnGuardar
            // 
            this.btnGuardar.Location = new System.Drawing.Point(270, 315);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(90, 30);
            this.btnGuardar.TabIndex = 17;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.BtnGuardar_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancelar.Location = new System.Drawing.Point(370, 315);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(90, 30);
            this.btnCancelar.TabIndex = 18;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            // 
            // FormTerceroAsignaturas
            // 
            this.AcceptButton = this.btnGuardar;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancelar;
            this.ClientSize = new System.Drawing.Size(484, 361);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.txtObservaciones);
            this.Controls.Add(this.lblObservaciones);
            this.Controls.Add(this.txtOrigen);
            this.Controls.Add(this.lblOrigen);
            this.Controls.Add(this.lblUltimaAsignacionValor);
            this.Controls.Add(this.lblUltimaAsignacionTitulo);
            this.Controls.Add(this.lblVecesAsignadaValor);
            this.Controls.Add(this.lblVecesAsignadaTitulo);
            this.Controls.Add(this.numPrioridad);
            this.Controls.Add(this.lblPrioridad);
            this.Controls.Add(this.chkActivo);
            this.Controls.Add(this.cmbAsignatura);
            this.Controls.Add(this.lblAsignatura);
            this.Controls.Add(this.cmbSemestre);
            this.Controls.Add(this.lblSemestre);
            this.Controls.Add(this.cmbPrograma);
            this.Controls.Add(this.lblPrograma);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormTerceroAsignaturas";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Habilitar Asignatura para Docente";
            ((System.ComponentModel.ISupportInitialize)(this.numPrioridad)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblPrograma;
        private System.Windows.Forms.ComboBox cmbPrograma;
        private System.Windows.Forms.Label lblSemestre;
        private System.Windows.Forms.ComboBox cmbSemestre;
        private System.Windows.Forms.Label lblAsignatura;
        private System.Windows.Forms.ComboBox cmbAsignatura;
        private System.Windows.Forms.CheckBox chkActivo;
        private System.Windows.Forms.Label lblPrioridad;
        private System.Windows.Forms.NumericUpDown numPrioridad;
        private System.Windows.Forms.Label lblVecesAsignadaTitulo;
        private System.Windows.Forms.Label lblVecesAsignadaValor;
        private System.Windows.Forms.Label lblUltimaAsignacionTitulo;
        private System.Windows.Forms.Label lblUltimaAsignacionValor;
        private System.Windows.Forms.Label lblOrigen;
        private System.Windows.Forms.TextBox txtOrigen;
        private System.Windows.Forms.Label lblObservaciones;
        private System.Windows.Forms.TextBox txtObservaciones;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnCancelar;
    }
}
