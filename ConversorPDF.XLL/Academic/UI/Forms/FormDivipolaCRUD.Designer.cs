namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormDivipolaCRUD
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
            this.dgvDivipola = new System.Windows.Forms.DataGridView();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.lblDepartamento = new System.Windows.Forms.Label();
            this.txtDepartamento = new System.Windows.Forms.TextBox();
            this.lblDepCodigo = new System.Windows.Forms.Label();
            this.txtDepCodigo = new System.Windows.Forms.TextBox();
            this.lblDepNorm = new System.Windows.Forms.Label();
            this.txtDepNorm = new System.Windows.Forms.TextBox();
            this.lblMunicipio = new System.Windows.Forms.Label();
            this.txtMunicipio = new System.Windows.Forms.TextBox();
            this.lblMunCodigo = new System.Windows.Forms.Label();
            this.txtMunCodigo = new System.Windows.Forms.TextBox();
            this.lblMunNorm = new System.Windows.Forms.Label();
            this.txtMunNorm = new System.Windows.Forms.TextBox();
            this.lblCategoria = new System.Windows.Forms.Label();
            this.numCategoria = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDivipola)).BeginInit();
            this.pnlDetails.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numCategoria)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvDivipola
            // 
            this.dgvDivipola.AllowUserToAddRows = false;
            this.dgvDivipola.AllowUserToDeleteRows = false;
            this.dgvDivipola.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDivipola.Dock = System.Windows.Forms.DockStyle.Left;
            this.dgvDivipola.Location = new System.Drawing.Point(0, 0);
            this.dgvDivipola.Name = "dgvDivipola";
            this.dgvDivipola.ReadOnly = true;
            this.dgvDivipola.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvDivipola.Size = new System.Drawing.Size(400, 361);
            this.dgvDivipola.TabIndex = 0;
            // 
            // pnlDetails
            // 
            this.pnlDetails.Controls.Add(this.btnCerrar);
            this.pnlDetails.Controls.Add(this.btnEliminar);
            this.pnlDetails.Controls.Add(this.btnGuardar);
            this.pnlDetails.Controls.Add(this.btnNuevo);
            this.pnlDetails.Controls.Add(this.numCategoria);
            this.pnlDetails.Controls.Add(this.lblCategoria);
            this.pnlDetails.Controls.Add(this.txtMunNorm);
            this.pnlDetails.Controls.Add(this.lblMunNorm);
            this.pnlDetails.Controls.Add(this.txtMunCodigo);
            this.pnlDetails.Controls.Add(this.lblMunCodigo);
            this.pnlDetails.Controls.Add(this.txtMunicipio);
            this.pnlDetails.Controls.Add(this.lblMunicipio);
            this.pnlDetails.Controls.Add(this.txtDepNorm);
            this.pnlDetails.Controls.Add(this.lblDepNorm);
            this.pnlDetails.Controls.Add(this.txtDepCodigo);
            this.pnlDetails.Controls.Add(this.lblDepCodigo);
            this.pnlDetails.Controls.Add(this.txtDepartamento);
            this.pnlDetails.Controls.Add(this.lblDepartamento);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDetails.Location = new System.Drawing.Point(400, 0);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(334, 361);
            this.pnlDetails.TabIndex = 1;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(240, 310);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 16;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Enabled = false;
            this.btnEliminar.ForeColor = System.Drawing.Color.Red;
            this.btnEliminar.Location = new System.Drawing.Point(159, 310);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(75, 23);
            this.btnEliminar.TabIndex = 15;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.Location = new System.Drawing.Point(78, 310);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 14;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnNuevo
            // 
            this.btnNuevo.Location = new System.Drawing.Point(17, 310);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.Size = new System.Drawing.Size(55, 23);
            this.btnNuevo.TabIndex = 13;
            this.btnNuevo.Text = "Nuevo";
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            // 
            // lblDepartamento
            // 
            this.lblDepartamento.AutoSize = true;
            this.lblDepartamento.Location = new System.Drawing.Point(14, 20);
            this.lblDepartamento.Name = "lblDepartamento";
            this.lblDepartamento.Size = new System.Drawing.Size(77, 13);
            this.lblDepartamento.TabIndex = 0;
            this.lblDepartamento.Text = "Departamento:";
            // 
            // txtDepartamento
            // 
            this.txtDepartamento.Location = new System.Drawing.Point(140, 17);
            this.txtDepartamento.Name = "txtDepartamento";
            this.txtDepartamento.Size = new System.Drawing.Size(160, 20);
            this.txtDepartamento.TabIndex = 1;
            // 
            // lblDepCodigo
            // 
            this.lblDepCodigo.AutoSize = true;
            this.lblDepCodigo.Location = new System.Drawing.Point(14, 50);
            this.lblDepCodigo.Name = "lblDepCodigo";
            this.lblDepCodigo.Size = new System.Drawing.Size(111, 13);
            this.lblDepCodigo.TabIndex = 2;
            this.lblDepCodigo.Text = "Código Departamento:";
            // 
            // txtDepCodigo
            // 
            this.txtDepCodigo.Location = new System.Drawing.Point(140, 47);
            this.txtDepCodigo.Name = "txtDepCodigo";
            this.txtDepCodigo.Size = new System.Drawing.Size(160, 20);
            this.txtDepCodigo.TabIndex = 3;
            // 
            // lblDepNorm
            // 
            this.lblDepNorm.AutoSize = true;
            this.lblDepNorm.Location = new System.Drawing.Point(14, 80);
            this.lblDepNorm.Name = "lblDepNorm";
            this.lblDepNorm.Size = new System.Drawing.Size(102, 13);
            this.lblDepNorm.TabIndex = 4;
            this.lblDepNorm.Text = "Dpto. Normalizado:";
            // 
            // txtDepNorm
            // 
            this.txtDepNorm.Location = new System.Drawing.Point(140, 77);
            this.txtDepNorm.Name = "txtDepNorm";
            this.txtDepNorm.Size = new System.Drawing.Size(160, 20);
            this.txtDepNorm.TabIndex = 5;
            // 
            // lblMunicipio
            // 
            this.lblMunicipio.AutoSize = true;
            this.lblMunicipio.Location = new System.Drawing.Point(14, 120);
            this.lblMunicipio.Name = "lblMunicipio";
            this.lblMunicipio.Size = new System.Drawing.Size(55, 13);
            this.lblMunicipio.TabIndex = 6;
            this.lblMunicipio.Text = "Municipio:";
            // 
            // txtMunicipio
            // 
            this.txtMunicipio.Location = new System.Drawing.Point(140, 117);
            this.txtMunicipio.Name = "txtMunicipio";
            this.txtMunicipio.Size = new System.Drawing.Size(160, 20);
            this.txtMunicipio.TabIndex = 7;
            // 
            // lblMunCodigo
            // 
            this.lblMunCodigo.AutoSize = true;
            this.lblMunCodigo.Location = new System.Drawing.Point(14, 150);
            this.lblMunCodigo.Name = "lblMunCodigo";
            this.lblMunCodigo.Size = new System.Drawing.Size(89, 13);
            this.lblMunCodigo.TabIndex = 8;
            this.lblMunCodigo.Text = "Código Municipio:";
            // 
            // txtMunCodigo
            // 
            this.txtMunCodigo.Location = new System.Drawing.Point(140, 147);
            this.txtMunCodigo.Name = "txtMunCodigo";
            this.txtMunCodigo.Size = new System.Drawing.Size(160, 20);
            this.txtMunCodigo.TabIndex = 9;
            // 
            // lblMunNorm
            // 
            this.lblMunNorm.AutoSize = true;
            this.lblMunNorm.Location = new System.Drawing.Point(14, 180);
            this.lblMunNorm.Name = "lblMunNorm";
            this.lblMunNorm.Size = new System.Drawing.Size(117, 13);
            this.lblMunNorm.TabIndex = 10;
            this.lblMunNorm.Text = "Municipio Normalizado:";
            // 
            // txtMunNorm
            // 
            this.txtMunNorm.Location = new System.Drawing.Point(140, 177);
            this.txtMunNorm.Name = "txtMunNorm";
            this.txtMunNorm.Size = new System.Drawing.Size(160, 20);
            this.txtMunNorm.TabIndex = 11;
            // 
            // lblCategoria
            // 
            this.lblCategoria.AutoSize = true;
            this.lblCategoria.Location = new System.Drawing.Point(14, 220);
            this.lblCategoria.Name = "lblCategoria";
            this.lblCategoria.Size = new System.Drawing.Size(57, 13);
            this.lblCategoria.TabIndex = 12;
            this.lblCategoria.Text = "Categoría:";
            // 
            // numCategoria
            // 
            this.numCategoria.Location = new System.Drawing.Point(140, 218);
            this.numCategoria.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numCategoria.Name = "numCategoria";
            this.numCategoria.Size = new System.Drawing.Size(80, 20);
            this.numCategoria.TabIndex = 12;
            // 
            // FormDivipolaCRUD
            // 
            this.ClientSize = new System.Drawing.Size(734, 361);
            this.Controls.Add(this.pnlDetails);
            this.Controls.Add(this.dgvDivipola);
            this.Name = "FormDivipolaCRUD";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestión de División Política (Divipola)";
            ((System.ComponentModel.ISupportInitialize)(this.dgvDivipola)).EndInit();
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numCategoria)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.DataGridView dgvDivipola;
        private System.Windows.Forms.Panel pnlDetails;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Label lblDepartamento;
        private System.Windows.Forms.TextBox txtDepartamento;
        private System.Windows.Forms.Label lblDepCodigo;
        private System.Windows.Forms.TextBox txtDepCodigo;
        private System.Windows.Forms.Label lblDepNorm;
        private System.Windows.Forms.TextBox txtDepNorm;
        private System.Windows.Forms.Label lblMunicipio;
        private System.Windows.Forms.TextBox txtMunicipio;
        private System.Windows.Forms.Label lblMunCodigo;
        private System.Windows.Forms.TextBox txtMunCodigo;
        private System.Windows.Forms.Label lblMunNorm;
        private System.Windows.Forms.TextBox txtMunNorm;
        private System.Windows.Forms.Label lblCategoria;
        private System.Windows.Forms.NumericUpDown numCategoria;
    }
}
