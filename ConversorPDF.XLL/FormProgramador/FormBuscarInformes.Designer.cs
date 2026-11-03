namespace ConversorPDF.XLL.FormProgramador
{
    partial class FormBuscarInformes
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblCarpeta;
        private System.Windows.Forms.TextBox txtCarpeta;
        private System.Windows.Forms.Button btnExaminar;
        private System.Windows.Forms.Label lblInicio;
        private System.Windows.Forms.TextBox txtViaticoInicial;
        private System.Windows.Forms.Label lblFin;
        private System.Windows.Forms.TextBox txtViaticoFinal;
        private System.Windows.Forms.CheckBox chkUnificar;
        private System.Windows.Forms.CheckBox chkAuditar;
        private System.Windows.Forms.CheckBox chkSimular;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Button btnProcesar;
        private System.Windows.Forms.Button btnCerrar;

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
            this.lblCarpeta = new System.Windows.Forms.Label();
            this.txtCarpeta = new System.Windows.Forms.TextBox();
            this.btnExaminar = new System.Windows.Forms.Button();
            this.lblInicio = new System.Windows.Forms.Label();
            this.txtViaticoInicial = new System.Windows.Forms.TextBox();
            this.lblFin = new System.Windows.Forms.Label();
            this.txtViaticoFinal = new System.Windows.Forms.TextBox();
            this.chkUnificar = new System.Windows.Forms.CheckBox();
            this.chkAuditar = new System.Windows.Forms.CheckBox();
            this.chkSimular = new System.Windows.Forms.CheckBox();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnProcesar = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblCarpeta
            // 
            this.lblCarpeta.AutoSize = true;
            this.lblCarpeta.Location = new System.Drawing.Point(12, 18);
            this.lblCarpeta.Name = "lblCarpeta";
            this.lblCarpeta.Size = new System.Drawing.Size(110, 13);
            this.lblCarpeta.TabIndex = 0;
            this.lblCarpeta.Text = "Carpeta de informes:";
            // 
            // txtCarpeta
            // 
            this.txtCarpeta.Location = new System.Drawing.Point(130, 15);
            this.txtCarpeta.Name = "txtCarpeta";
            this.txtCarpeta.ReadOnly = true;
            this.txtCarpeta.Size = new System.Drawing.Size(470, 20);
            this.txtCarpeta.TabIndex = 1;
            // 
            // btnExaminar
            // 
            this.btnExaminar.Location = new System.Drawing.Point(608, 16);
            this.btnExaminar.Name = "btnExaminar";
            this.btnExaminar.Size = new System.Drawing.Size(92, 24);
            this.btnExaminar.TabIndex = 2;
            this.btnExaminar.Text = "Examinar...";
            this.btnExaminar.UseVisualStyleBackColor = true;
            // 
            // lblInicio
            // 
            this.lblInicio.AutoSize = true;
            this.lblInicio.Location = new System.Drawing.Point(12, 52);
            this.lblInicio.Name = "lblInicio";
            this.lblInicio.Size = new System.Drawing.Size(92, 13);
            this.lblInicio.TabIndex = 3;
            this.lblInicio.Text = "Solicitud inicial:";
            // 
            // txtViaticoInicial
            // 
            this.txtViaticoInicial.Location = new System.Drawing.Point(130, 49);
            this.txtViaticoInicial.Name = "txtViaticoInicial";
            this.txtViaticoInicial.Size = new System.Drawing.Size(100, 20);
            this.txtViaticoInicial.TabIndex = 4;
            // 
            // lblFin
            // 
            this.lblFin.AutoSize = true;
            this.lblFin.Location = new System.Drawing.Point(250, 52);
            this.lblFin.Name = "lblFin";
            this.lblFin.Size = new System.Drawing.Size(83, 13);
            this.lblFin.TabIndex = 5;
            this.lblFin.Text = "Solicitud final:";
            // 
            // txtViaticoFinal
            // 
            this.txtViaticoFinal.Location = new System.Drawing.Point(350, 49);
            this.txtViaticoFinal.Name = "txtViaticoFinal";
            this.txtViaticoFinal.Size = new System.Drawing.Size(100, 20);
            this.txtViaticoFinal.TabIndex = 6;
            // 
            // chkUnificar
            // 
            this.chkUnificar.AutoSize = true;
            this.chkUnificar.Checked = true;
            this.chkUnificar.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUnificar.Location = new System.Drawing.Point(130, 80);
            this.chkUnificar.Name = "chkUnificar";
            this.chkUnificar.Size = new System.Drawing.Size(215, 17);
            this.chkUnificar.TabIndex = 7;
            this.chkUnificar.Text = "Unificar parejas INF + CUM (archivo UNI)";
            this.chkUnificar.UseVisualStyleBackColor = true;
            // 
            // chkAuditar
            // 
            this.chkAuditar.AutoSize = true;
            this.chkAuditar.Checked = true;
            this.chkAuditar.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAuditar.Location = new System.Drawing.Point(130, 103);
            this.chkAuditar.Name = "chkAuditar";
            this.chkAuditar.Size = new System.Drawing.Size(250, 17);
            this.chkAuditar.TabIndex = 8;
            this.chkAuditar.Text = "Actualizar columna INFORME de CONSULTAS";
            this.chkAuditar.UseVisualStyleBackColor = true;
            // 
            // chkSimular
            // 
            this.chkSimular.AutoSize = true;
            this.chkSimular.Location = new System.Drawing.Point(130, 126);
            this.chkSimular.Name = "chkSimular";
            this.chkSimular.Size = new System.Drawing.Size(290, 17);
            this.chkSimular.TabIndex = 9;
            this.chkSimular.Text = "Simulación (no modifica archivos ni escribe en Excel)";
            this.chkSimular.UseVisualStyleBackColor = true;
            // 
            // txtLog
            // 
            this.txtLog.BackColor = System.Drawing.SystemColors.Window;
            this.txtLog.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLog.Location = new System.Drawing.Point(12, 155);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLog.Size = new System.Drawing.Size(688, 330);
            this.txtLog.TabIndex = 15;
            this.txtLog.WordWrap = false;
            // 
            // btnProcesar
            // 
            this.btnProcesar.Location = new System.Drawing.Point(518, 495);
            this.btnProcesar.Name = "btnProcesar";
            this.btnProcesar.Size = new System.Drawing.Size(92, 30);
            this.btnProcesar.TabIndex = 16;
            this.btnProcesar.Text = "Procesar";
            this.btnProcesar.UseVisualStyleBackColor = true;
            // 
            // btnCerrar
            // 
            this.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCerrar.Location = new System.Drawing.Point(616, 495);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(84, 30);
            this.btnCerrar.TabIndex = 17;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            // 
            // FormBuscarInformes
            // 
            this.AcceptButton = this.btnProcesar;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCerrar;
            this.ClientSize = new System.Drawing.Size(712, 537);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.btnProcesar);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.chkSimular);
            this.Controls.Add(this.chkAuditar);
            this.Controls.Add(this.chkUnificar);
            this.Controls.Add(this.txtViaticoFinal);
            this.Controls.Add(this.lblFin);
            this.Controls.Add(this.txtViaticoInicial);
            this.Controls.Add(this.lblInicio);
            this.Controls.Add(this.btnExaminar);
            this.Controls.Add(this.txtCarpeta);
            this.Controls.Add(this.lblCarpeta);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormBuscarInformes";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Buscar Informes";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
