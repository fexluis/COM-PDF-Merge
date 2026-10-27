namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormGastosTerrestres
    {
        private System.ComponentModel.IContainer components = null;

        // El método Dispose(bool) fue movido al code-behind (FormGastosTerrestres.cs) 
        // para manejar adecuadamente la liberación de la conexión a base de datos.

        private void InitializeComponent()
        {
            this.dgvGastos = new System.Windows.Forms.DataGridView();
            this.pnlEdicion = new System.Windows.Forms.Panel();
            this.lblSubunidad = new System.Windows.Forms.Label();
            this.cmbSubunidad = new System.Windows.Forms.ComboBox();
            this.lblCodigo = new System.Windows.Forms.Label();
            this.txtCodigoTrayecto = new System.Windows.Forms.TextBox();
            this.lblTrayecto = new System.Windows.Forms.Label();
            this.txtTrayecto = new System.Windows.Forms.TextBox();
            this.lblValor = new System.Windows.Forms.Label();
            this.numValor = new System.Windows.Forms.NumericUpDown();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            
            ((System.ComponentModel.ISupportInitialize)(this.dgvGastos)).BeginInit();
            this.pnlEdicion.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numValor)).BeginInit();
            this.SuspendLayout();
            
            // dgvGastos
            this.dgvGastos.AllowUserToAddRows = false;
            this.dgvGastos.AllowUserToDeleteRows = false;
            this.dgvGastos.AutoGenerateColumns = false;
            this.dgvGastos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGastos.Location = new System.Drawing.Point(12, 12);
            this.dgvGastos.Name = "dgvGastos";
            this.dgvGastos.ReadOnly = true;
            this.dgvGastos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvGastos.Size = new System.Drawing.Size(560, 250);
            this.dgvGastos.TabIndex = 0;
            
            // Adding columns explicitly
            this.dgvGastos.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn { 
                DataPropertyName = "Id", HeaderText = "Id", Visible = false });
            this.dgvGastos.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn { 
                DataPropertyName = "Territorial", HeaderText = "Territorial", Width = 100 });
            this.dgvGastos.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn { 
                DataPropertyName = "CodigoTrayecto", HeaderText = "Código", Width = 100 });
            this.dgvGastos.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn { 
                DataPropertyName = "Trayecto", HeaderText = "Trayecto", Width = 200 });
            this.dgvGastos.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn { 
                DataPropertyName = "Valor", HeaderText = "Valor", Width = 100 });
            
            this.dgvGastos.SelectionChanged += new System.EventHandler(this.dgvGastos_SelectionChanged);
            
            // pnlEdicion
            this.pnlEdicion.Controls.Add(this.lblSubunidad);
            this.pnlEdicion.Controls.Add(this.cmbSubunidad);
            this.pnlEdicion.Controls.Add(this.lblCodigo);
            this.pnlEdicion.Controls.Add(this.txtCodigoTrayecto);
            this.pnlEdicion.Controls.Add(this.lblTrayecto);
            this.pnlEdicion.Controls.Add(this.txtTrayecto);
            this.pnlEdicion.Controls.Add(this.lblValor);
            this.pnlEdicion.Controls.Add(this.numValor);
            this.pnlEdicion.Location = new System.Drawing.Point(12, 275);
            this.pnlEdicion.Name = "pnlEdicion";
            this.pnlEdicion.Size = new System.Drawing.Size(560, 110);
            this.pnlEdicion.TabIndex = 1;
            
            // lblSubunidad
            this.lblSubunidad.AutoSize = true;
            this.lblSubunidad.Location = new System.Drawing.Point(15, 15);
            this.lblSubunidad.Name = "lblSubunidad";
            this.lblSubunidad.Size = new System.Drawing.Size(117, 13);
            this.lblSubunidad.TabIndex = 0;
            this.lblSubunidad.Text = "Territorial / Subunidad:";
            
            // cmbSubunidad
            this.cmbSubunidad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubunidad.FormattingEnabled = true;
            this.cmbSubunidad.Location = new System.Drawing.Point(135, 12);
            this.cmbSubunidad.Name = "cmbSubunidad";
            this.cmbSubunidad.Size = new System.Drawing.Size(150, 21);
            this.cmbSubunidad.TabIndex = 1;
            
            // lblCodigo
            this.lblCodigo.AutoSize = true;
            this.lblCodigo.Location = new System.Drawing.Point(310, 15);
            this.lblCodigo.Name = "lblCodigo";
            this.lblCodigo.Size = new System.Drawing.Size(87, 13);
            this.lblCodigo.TabIndex = 2;
            this.lblCodigo.Text = "Código Trayecto:";
            
            // txtCodigoTrayecto
            this.txtCodigoTrayecto.Location = new System.Drawing.Point(400, 12);
            this.txtCodigoTrayecto.Name = "txtCodigoTrayecto";
            this.txtCodigoTrayecto.ReadOnly = true;
            this.txtCodigoTrayecto.Size = new System.Drawing.Size(140, 20);
            this.txtCodigoTrayecto.TabIndex = 3;
            
            // lblTrayecto
            this.lblTrayecto.AutoSize = true;
            this.lblTrayecto.Location = new System.Drawing.Point(15, 50);
            this.lblTrayecto.Name = "lblTrayecto";
            this.lblTrayecto.Size = new System.Drawing.Size(52, 13);
            this.lblTrayecto.TabIndex = 4;
            this.lblTrayecto.Text = "Trayecto:";
            
            // txtTrayecto
            this.txtTrayecto.Location = new System.Drawing.Point(135, 47);
            this.txtTrayecto.Name = "txtTrayecto";
            this.txtTrayecto.Size = new System.Drawing.Size(405, 20);
            this.txtTrayecto.TabIndex = 5;
            
            // lblValor
            this.lblValor.AutoSize = true;
            this.lblValor.Location = new System.Drawing.Point(15, 85);
            this.lblValor.Name = "lblValor";
            this.lblValor.Size = new System.Drawing.Size(34, 13);
            this.lblValor.TabIndex = 6;
            this.lblValor.Text = "Valor:";
            
            // numValor
            this.numValor.Location = new System.Drawing.Point(135, 82);
            this.numValor.Maximum = new decimal(new int[] { 100000000, 0, 0, 0 });
            this.numValor.Name = "numValor";
            this.numValor.Size = new System.Drawing.Size(150, 20);
            this.numValor.TabIndex = 7;
            
            // btnNuevo
            this.btnNuevo.Location = new System.Drawing.Point(12, 400);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.Size = new System.Drawing.Size(90, 30);
            this.btnNuevo.TabIndex = 2;
            this.btnNuevo.Text = "Nuevo";
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            
            // btnGuardar
            this.btnGuardar.Location = new System.Drawing.Point(112, 400);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(90, 30);
            this.btnGuardar.TabIndex = 3;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            
            // btnEliminar
            this.btnEliminar.Location = new System.Drawing.Point(212, 400);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(90, 30);
            this.btnEliminar.TabIndex = 4;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            
            // FormGastosTerrestres
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 441);
            this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.btnNuevo);
            this.Controls.Add(this.pnlEdicion);
            this.Controls.Add(this.dgvGastos);
            this.Name = "FormGastosTerrestres";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Base Gastos Terrestres";
            this.Load += new System.EventHandler(this.FormGastosTerrestres_Load);
            
            ((System.ComponentModel.ISupportInitialize)(this.dgvGastos)).EndInit();
            this.pnlEdicion.ResumeLayout(false);
            this.pnlEdicion.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numValor)).EndInit();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.DataGridView dgvGastos;
        private System.Windows.Forms.Panel pnlEdicion;
        private System.Windows.Forms.Label lblSubunidad;
        private System.Windows.Forms.ComboBox cmbSubunidad;
        private System.Windows.Forms.Label lblCodigo;
        private System.Windows.Forms.TextBox txtCodigoTrayecto;
        private System.Windows.Forms.Label lblTrayecto;
        private System.Windows.Forms.TextBox txtTrayecto;
        private System.Windows.Forms.Label lblValor;
        private System.Windows.Forms.NumericUpDown numValor;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnEliminar;
    }
}
