namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormCdpToPDF
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblCdp;
        private System.Windows.Forms.ComboBox cmbCdp;
        private System.Windows.Forms.Button btnPdfViaticos;
        private System.Windows.Forms.Button btnMasivo;
        private System.Windows.Forms.PictureBox picLogo;

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
            this.lblCdp = new System.Windows.Forms.Label();
            this.cmbCdp = new System.Windows.Forms.ComboBox();
            this.btnPdfViaticos = new System.Windows.Forms.Button();
            this.btnMasivo = new System.Windows.Forms.Button();
            this.picLogo = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // picLogo
            // 
            this.picLogo.Location = new System.Drawing.Point(280, 10);
            this.picLogo.Name = "picLogo";
            this.picLogo.Size = new System.Drawing.Size(100, 40);
            this.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picLogo.TabIndex = 4;
            this.picLogo.TabStop = false;
            // 
            // lblCdp
            // 
            this.lblCdp.AutoSize = true;
            this.lblCdp.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCdp.ForeColor = System.Drawing.Color.FromArgb(0, 51, 153);
            this.lblCdp.Location = new System.Drawing.Point(21, 23);
            this.lblCdp.Name = "lblCdp";
            this.lblCdp.Size = new System.Drawing.Size(182, 21);
            this.lblCdp.TabIndex = 0;
            this.lblCdp.Text = "Seleccione CDP a usar:";
            // 
            // cmbCdp
            // 
            this.cmbCdp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCdp.FormattingEnabled = true;
            this.cmbCdp.Location = new System.Drawing.Point(24, 60);
            this.cmbCdp.Name = "cmbCdp";
            this.cmbCdp.Size = new System.Drawing.Size(350, 21);
            this.cmbCdp.TabIndex = 1;
            // 
            // btnPdfViaticos
            // 
            this.btnPdfViaticos.BackColor = System.Drawing.Color.FromArgb(0, 51, 153);
            this.btnPdfViaticos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPdfViaticos.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPdfViaticos.ForeColor = System.Drawing.Color.White;
            this.btnPdfViaticos.Location = new System.Drawing.Point(62, 95);
            this.btnPdfViaticos.Name = "btnPdfViaticos";
            this.btnPdfViaticos.Size = new System.Drawing.Size(120, 35);
            this.btnPdfViaticos.TabIndex = 2;
            this.btnPdfViaticos.Text = "PDF Viaticos";
            this.btnPdfViaticos.UseVisualStyleBackColor = false;
            this.btnPdfViaticos.Click += new System.EventHandler(this.btnPdfViaticos_Click);
            // 
            // btnMasivo
            // 
            this.btnMasivo.BackColor = System.Drawing.Color.FromArgb(0, 51, 153);
            this.btnMasivo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMasivo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMasivo.ForeColor = System.Drawing.Color.White;
            this.btnMasivo.Location = new System.Drawing.Point(217, 95);
            this.btnMasivo.Name = "btnMasivo";
            this.btnMasivo.Size = new System.Drawing.Size(120, 35);
            this.btnMasivo.TabIndex = 3;
            this.btnMasivo.Text = "Masivo";
            this.btnMasivo.UseVisualStyleBackColor = false;
            this.btnMasivo.Click += new System.EventHandler(this.btnMasivo_Click);
            // 
            // FormCdpToPDF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(400, 150);
            this.Controls.Add(this.picLogo);
            this.Controls.Add(this.btnMasivo);
            this.Controls.Add(this.btnPdfViaticos);
            this.Controls.Add(this.cmbCdp);
            this.Controls.Add(this.lblCdp);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormCdpToPDF";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Exportar a PDF";
            this.Load += new System.EventHandler(this.FormCdpToPDF_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
