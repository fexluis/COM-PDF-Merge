namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormSedes
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _bs?.Dispose();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.dgvSedes = new System.Windows.Forms.DataGridView();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.numTransporte = new System.Windows.Forms.NumericUpDown();
            this.lblTransporte = new System.Windows.Forms.Label();
            this.txtCategoria = new System.Windows.Forms.TextBox();
            this.lblCategoria = new System.Windows.Forms.Label();
            this.txtDepartamento = new System.Windows.Forms.TextBox();
            this.lblDepartamento = new System.Windows.Forms.Label();
            this.cmbTerritoriales = new System.Windows.Forms.ComboBox();
            this.lblTerritoriales = new System.Windows.Forms.Label();
            this.txtCiudad = new System.Windows.Forms.TextBox();
            this.lblCiudad = new System.Windows.Forms.Label();
            this.txtSubsede = new System.Windows.Forms.TextBox();
            this.lblSubsede = new System.Windows.Forms.Label();
            this.cmbActivo = new System.Windows.Forms.ComboBox();
            this.lblActivo = new System.Windows.Forms.Label();
            this.btnSelectDivipola = new System.Windows.Forms.Button();
            this.btnDivipolaCRUD = new System.Windows.Forms.Button();
            this.btnReintentarSubunidades = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSedes)).BeginInit();
            this.pnlDetails.SuspendLayout();
            this.txtDepartamentoSiif = new System.Windows.Forms.TextBox();
            this.lblDepartamentoSiif = new System.Windows.Forms.Label();
            this.txtCiudadSiif = new System.Windows.Forms.TextBox();
            this.lblCiudadSiif = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // dgvSedes
            // 
            this.dgvSedes.AllowUserToAddRows = false;
            this.dgvSedes.AllowUserToDeleteRows = false;
            this.dgvSedes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSedes.Dock = System.Windows.Forms.DockStyle.Left;
            this.dgvSedes.Location = new System.Drawing.Point(0, 0);
            this.dgvSedes.Name = "dgvSedes";
            this.dgvSedes.ReadOnly = true;
            this.dgvSedes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvSedes.Size = new System.Drawing.Size(400, 400);
            this.dgvSedes.TabIndex = 0;
            // 
            // pnlDetails
            // 
            this.pnlDetails.Controls.Add(this.btnCerrar);
            this.pnlDetails.Controls.Add(this.btnEliminar);
            this.pnlDetails.Controls.Add(this.btnGuardar);
            this.pnlDetails.Controls.Add(this.btnNuevo);
            this.pnlDetails.Controls.Add(this.numTransporte);
            this.pnlDetails.Controls.Add(this.lblTransporte);
            this.pnlDetails.Controls.Add(this.txtCategoria);
            this.pnlDetails.Controls.Add(this.lblCategoria);
            this.pnlDetails.Controls.Add(this.txtDepartamento);
            this.pnlDetails.Controls.Add(this.lblDepartamento);
            this.pnlDetails.Controls.Add(this.btnSelectDivipola);
            this.pnlDetails.Controls.Add(this.btnDivipolaCRUD);
            this.pnlDetails.Controls.Add(this.cmbActivo);
            this.pnlDetails.Controls.Add(this.lblActivo);
            this.pnlDetails.Controls.Add(this.txtSubsede);
            this.pnlDetails.Controls.Add(this.lblSubsede);
            this.pnlDetails.Controls.Add(this.txtCiudad);
            this.pnlDetails.Controls.Add(this.lblCiudad);
            this.pnlDetails.Controls.Add(this.cmbTerritoriales);
            this.pnlDetails.Controls.Add(this.lblTerritoriales);
            this.pnlDetails.Controls.Add(this.txtCiudadSiif);
            this.pnlDetails.Controls.Add(this.lblCiudadSiif);
            this.pnlDetails.Controls.Add(this.txtDepartamentoSiif);
            this.pnlDetails.Controls.Add(this.lblDepartamentoSiif);
            this.pnlDetails.Controls.Add(this.btnReintentarSubunidades);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDetails.Location = new System.Drawing.Point(400, 0);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(300, 400);
            this.pnlDetails.TabIndex = 21;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(210, 355);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 11;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Enabled = false;
            this.btnEliminar.ForeColor = System.Drawing.Color.Red;
            this.btnEliminar.Location = new System.Drawing.Point(129, 355);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(75, 23);
            this.btnEliminar.TabIndex = 10;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.Location = new System.Drawing.Point(48, 355);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 9;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnNuevo
            // 
            this.btnNuevo.Location = new System.Drawing.Point(10, 355);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.Size = new System.Drawing.Size(32, 23);
            this.btnNuevo.TabIndex = 8;
            this.btnNuevo.Text = "+";
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            //
            // txtDepartamento
            // 
            this.txtDepartamento.Location = new System.Drawing.Point(100, 80);
            this.txtDepartamento.Name = "txtDepartamento";
            this.txtDepartamento.ReadOnly = true;
            this.txtDepartamento.Size = new System.Drawing.Size(150, 20);
            this.txtDepartamento.TabIndex = 3;
            // 
            // lblDepartamento
            // 
            this.lblDepartamento.Location = new System.Drawing.Point(20, 83);
            this.lblDepartamento.Name = "lblDepartamento";
            this.lblDepartamento.Size = new System.Drawing.Size(80, 13);
            this.lblDepartamento.TabIndex = 2;
            this.lblDepartamento.Text = "Dep.DIP:";
            // 
            // txtDepartamentoSiif
            // 
            this.txtDepartamentoSiif.Location = new System.Drawing.Point(100, 110);
            this.txtDepartamentoSiif.Name = "txtDepartamentoSiif";
            this.txtDepartamentoSiif.Size = new System.Drawing.Size(150, 20);
            this.txtDepartamentoSiif.TabIndex = 4;
            // 
            // lblDepartamentoSiif
            // 
            this.lblDepartamentoSiif.Location = new System.Drawing.Point(10, 113);
            this.lblDepartamentoSiif.Name = "lblDepartamentoSiif";
            this.lblDepartamentoSiif.Size = new System.Drawing.Size(90, 13);
            this.lblDepartamentoSiif.TabIndex = 3;
            this.lblDepartamentoSiif.Text = "Dep.SIIF(Sin tilde):";
            // 
            // txtCiudadSiif
            // 
            this.txtCiudadSiif.Location = new System.Drawing.Point(100, 140);
            this.txtCiudadSiif.Name = "txtCiudadSiif";
            this.txtCiudadSiif.Size = new System.Drawing.Size(150, 20);
            this.txtCiudadSiif.TabIndex = 5;
            // 
            // lblCiudadSiif
            // 
            this.lblCiudadSiif.Location = new System.Drawing.Point(10, 143);
            this.lblCiudadSiif.Name = "lblCiudadSiif";
            this.lblCiudadSiif.Size = new System.Drawing.Size(90, 13);
            this.lblCiudadSiif.TabIndex = 4;
            this.lblCiudadSiif.Text = "Ciu.SIIF(Sin tilde):";
            // 
            // txtCategoria
            // 
            this.txtCategoria.Location = new System.Drawing.Point(100, 170);
            this.txtCategoria.Name = "txtCategoria";
            this.txtCategoria.ReadOnly = true;
            this.txtCategoria.Size = new System.Drawing.Size(100, 20);
            this.txtCategoria.TabIndex = 6;
            // 
            // lblCategoria
            // 
            this.lblCategoria.Location = new System.Drawing.Point(20, 173);
            this.lblCategoria.Name = "lblCategoria";
            this.lblCategoria.Size = new System.Drawing.Size(80, 13);
            this.lblCategoria.TabIndex = 7;
            this.lblCategoria.Text = "Categoría:";
            // 
            // numTransporte
            // 
            this.numTransporte.DecimalPlaces = 2;
            this.numTransporte.Location = new System.Drawing.Point(100, 200);
            this.numTransporte.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numTransporte.Name = "numTransporte";
            this.numTransporte.Size = new System.Drawing.Size(120, 20);
            this.numTransporte.TabIndex = 8;
            // 
            // lblTransporte
            // 
            this.lblTransporte.Location = new System.Drawing.Point(20, 203);
            this.lblTransporte.Name = "lblTransporte";
            this.lblTransporte.Size = new System.Drawing.Size(80, 13);
            this.lblTransporte.TabIndex = 9;
            this.lblTransporte.Text = "Transporte:";
            // 
            // cmbTerritoriales
            // 
            this.cmbTerritoriales.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTerritoriales.Location = new System.Drawing.Point(100, 20);
            this.cmbTerritoriales.Name = "cmbTerritoriales";
            this.cmbTerritoriales.Size = new System.Drawing.Size(150, 21);
            this.cmbTerritoriales.TabIndex = 18;
            // 
            // lblTerritoriales
            // 
            this.lblTerritoriales.Location = new System.Drawing.Point(20, 23);
            this.lblTerritoriales.Name = "lblTerritoriales";
            this.lblTerritoriales.Size = new System.Drawing.Size(80, 13);
            this.lblTerritoriales.TabIndex = 19;
            this.lblTerritoriales.Text = "Territorial:";
            // 
            // txtCiudad
            // 
            this.txtCiudad.Location = new System.Drawing.Point(100, 50);
            this.txtCiudad.Name = "txtCiudad";
            this.txtCiudad.ReadOnly = true;
            this.txtCiudad.Size = new System.Drawing.Size(150, 20);
            this.txtCiudad.TabIndex = 1;
            // 
            // lblCiudad
            // 
            this.lblCiudad.Location = new System.Drawing.Point(20, 53);
            this.lblCiudad.Name = "lblCiudad";
            this.lblCiudad.Size = new System.Drawing.Size(80, 13);
            this.lblCiudad.TabIndex = 20;
            this.lblCiudad.Text = "Ciudad.DIP:";
            // 
            // txtSubsede
            // 
            this.txtSubsede.Location = new System.Drawing.Point(100, 230);
            this.txtSubsede.Name = "txtSubsede";
            this.txtSubsede.Size = new System.Drawing.Size(150, 20);
            this.txtSubsede.TabIndex = 10;
            // 
            // lblSubsede
            // 
            this.lblSubsede.Location = new System.Drawing.Point(20, 233);
            this.lblSubsede.Name = "lblSubsede";
            this.lblSubsede.Size = new System.Drawing.Size(80, 13);
            this.lblSubsede.TabIndex = 11;
            this.lblSubsede.Text = "Subsede:";
            // 
            // cmbActivo
            // 
            this.cmbActivo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbActivo.Location = new System.Drawing.Point(100, 260);
            this.cmbActivo.Name = "cmbActivo";
            this.cmbActivo.Size = new System.Drawing.Size(120, 21);
            this.cmbActivo.TabIndex = 12;
            // 
            // lblActivo
            // 
            this.lblActivo.Location = new System.Drawing.Point(20, 263);
            this.lblActivo.Name = "lblActivo";
            this.lblActivo.Size = new System.Drawing.Size(80, 13);
            this.lblActivo.TabIndex = 13;
            this.lblActivo.Text = "Estado:";
            // 
            // btnSelectDivipola
            // 
            this.btnSelectDivipola.Location = new System.Drawing.Point(75, 290);
            this.btnSelectDivipola.Name = "btnSelectDivipola";
            this.btnSelectDivipola.Size = new System.Drawing.Size(150, 23);
            this.btnSelectDivipola.TabIndex = 14;
            this.btnSelectDivipola.Text = "SELECCIONAR SEDE";
            this.btnSelectDivipola.UseVisualStyleBackColor = true;
            this.btnSelectDivipola.Click += new System.EventHandler(this.btnSelectDivipola_Click);
            // 
            // btnDivipolaCRUD
            // 
            this.btnDivipolaCRUD.Location = new System.Drawing.Point(10, 320);
            this.btnDivipolaCRUD.Name = "btnDivipolaCRUD";
            this.btnDivipolaCRUD.Size = new System.Drawing.Size(275, 23);
            this.btnDivipolaCRUD.TabIndex = 15;
            this.btnDivipolaCRUD.Text = "DIVISIÓN POLÍTICA";
            this.btnDivipolaCRUD.UseVisualStyleBackColor = true;
            this.btnDivipolaCRUD.Click += new System.EventHandler(this.btnDivipolaCRUD_Click);
            // 
            // btnReintentarSubunidades
            // 
            this.btnReintentarSubunidades.Location = new System.Drawing.Point(255, 19);
            this.btnReintentarSubunidades.Name = "btnReintentarSubunidades";
            this.btnReintentarSubunidades.Size = new System.Drawing.Size(30, 23);
            this.btnReintentarSubunidades.TabIndex = 22;
            this.btnReintentarSubunidades.Text = "↻";
            this.btnReintentarSubunidades.UseVisualStyleBackColor = true;
            this.btnReintentarSubunidades.Click += new System.EventHandler(this.btnReintentarSubunidades_Click);
            // 
            // FormSedes
            // 
            this.ClientSize = new System.Drawing.Size(700, 400);
            this.Controls.Add(this.pnlDetails);
            this.Controls.Add(this.dgvSedes);
            this.Name = "FormSedes";
            this.Text = "Gestión de Sedes";
            ((System.ComponentModel.ISupportInitialize)(this.dgvSedes)).EndInit();
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTransporte)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.DataGridView dgvSedes;
        private System.Windows.Forms.Panel pnlDetails;
        private System.Windows.Forms.TextBox txtCiudad;
        private System.Windows.Forms.Label lblCiudad;
        private System.Windows.Forms.TextBox txtDepartamento;
        private System.Windows.Forms.Label lblDepartamento;
        private System.Windows.Forms.TextBox txtCategoria;
        private System.Windows.Forms.Label lblCategoria;
        private System.Windows.Forms.NumericUpDown numTransporte;
        private System.Windows.Forms.Label lblTransporte;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.TextBox txtSubsede;
        private System.Windows.Forms.Label lblSubsede;
        private System.Windows.Forms.ComboBox cmbActivo;
        private System.Windows.Forms.Label lblActivo;
        private System.Windows.Forms.ComboBox cmbTerritoriales;
        private System.Windows.Forms.Label lblTerritoriales;
        private System.Windows.Forms.Button btnReintentarSubunidades;
        private System.Windows.Forms.Button btnSelectDivipola;
        private System.Windows.Forms.Button btnDivipolaCRUD;
        private System.Windows.Forms.TextBox txtDepartamentoSiif;
        private System.Windows.Forms.Label lblDepartamentoSiif;
        private System.Windows.Forms.TextBox txtCiudadSiif;
        private System.Windows.Forms.Label lblCiudadSiif;
    }
}
