namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class AgFormTerceros
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
            this.pnlTop = new System.Windows.Forms.Panel();
            this.chkMostrarInactivos = new System.Windows.Forms.CheckBox();
            this.chkSoloHabilitadas = new System.Windows.Forms.CheckBox();
            this.txtBuscar = new System.Windows.Forms.TextBox();
            this.lblBuscar = new System.Windows.Forms.Label();
            this.cmbTerritorial = new System.Windows.Forms.ComboBox();
            this.lblTerritorial = new System.Windows.Forms.Label();
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.lblConteo = new System.Windows.Forms.Label();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnAceptar = new System.Windows.Forms.Button();
            this.btnDesmarcarTodos = new System.Windows.Forms.Button();
            this.btnMarcarTodos = new System.Windows.Forms.Button();
            this.dgvTerceros = new System.Windows.Forms.DataGridView();
            this.colSeleccion = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colNombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCedula = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTerritorial = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCategoria = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnlTop.SuspendLayout();
            this.pnlBottom.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerceros)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTop
            // 
            this.pnlTop.Controls.Add(this.chkMostrarInactivos);
            this.pnlTop.Controls.Add(this.chkSoloHabilitadas);
            this.pnlTop.Controls.Add(this.txtBuscar);
            this.pnlTop.Controls.Add(this.lblBuscar);
            this.pnlTop.Controls.Add(this.cmbTerritorial);
            this.pnlTop.Controls.Add(this.lblTerritorial);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(784, 75);
            this.pnlTop.TabIndex = 0;
            // 
            // chkMostrarInactivos
            // 
            this.chkMostrarInactivos.AutoSize = true;
            this.chkMostrarInactivos.Location = new System.Drawing.Point(520, 43);
            this.chkMostrarInactivos.Name = "chkMostrarInactivos";
            this.chkMostrarInactivos.Size = new System.Drawing.Size(117, 19);
            this.chkMostrarInactivos.TabIndex = 5;
            this.chkMostrarInactivos.Text = "Mostrar inactivos";
            this.chkMostrarInactivos.UseVisualStyleBackColor = true;
            // 
            // chkSoloHabilitadas
            // 
            this.chkSoloHabilitadas.AutoSize = true;
            this.chkSoloHabilitadas.Location = new System.Drawing.Point(520, 14);
            this.chkSoloHabilitadas.Name = "chkSoloHabilitadas";
            this.chkSoloHabilitadas.Size = new System.Drawing.Size(188, 19);
            this.chkSoloHabilitadas.TabIndex = 4;
            this.chkSoloHabilitadas.Text = "Solo asignaturas habilitadas";
            this.chkSoloHabilitadas.UseVisualStyleBackColor = true;
            // 
            // txtBuscar
            // 
            this.txtBuscar.Location = new System.Drawing.Point(90, 41);
            this.txtBuscar.Name = "txtBuscar";
            this.txtBuscar.Size = new System.Drawing.Size(400, 23);
            this.txtBuscar.TabIndex = 3;
            // 
            // lblBuscar
            // 
            this.lblBuscar.AutoSize = true;
            this.lblBuscar.Location = new System.Drawing.Point(15, 44);
            this.lblBuscar.Name = "lblBuscar";
            this.lblBuscar.Size = new System.Drawing.Size(45, 15);
            this.lblBuscar.TabIndex = 2;
            this.lblBuscar.Text = "Buscar:";
            // 
            // cmbTerritorial
            // 
            this.cmbTerritorial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTerritorial.FormattingEnabled = true;
            this.cmbTerritorial.Location = new System.Drawing.Point(90, 12);
            this.cmbTerritorial.Name = "cmbTerritorial";
            this.cmbTerritorial.Size = new System.Drawing.Size(400, 23);
            this.cmbTerritorial.TabIndex = 1;
            // 
            // lblTerritorial
            // 
            this.lblTerritorial.AutoSize = true;
            this.lblTerritorial.Location = new System.Drawing.Point(15, 15);
            this.lblTerritorial.Name = "lblTerritorial";
            this.lblTerritorial.Size = new System.Drawing.Size(60, 15);
            this.lblTerritorial.TabIndex = 0;
            this.lblTerritorial.Text = "Territorial:";
            // 
            // pnlBottom
            // 
            this.pnlBottom.Controls.Add(this.lblConteo);
            this.pnlBottom.Controls.Add(this.btnCancelar);
            this.pnlBottom.Controls.Add(this.btnAceptar);
            this.pnlBottom.Controls.Add(this.btnDesmarcarTodos);
            this.pnlBottom.Controls.Add(this.btnMarcarTodos);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 466);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(784, 45);
            this.pnlBottom.TabIndex = 1;
            // 
            // lblConteo
            // 
            this.lblConteo.AutoSize = true;
            this.lblConteo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblConteo.Location = new System.Drawing.Point(240, 15);
            this.lblConteo.Name = "lblConteo";
            this.lblConteo.Size = new System.Drawing.Size(148, 15);
            this.lblConteo.TabIndex = 4;
            this.lblConteo.Text = "0 de 0 seleccionados";
            // 
            // btnCancelar
            // 
            this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancelar.Location = new System.Drawing.Point(685, 10);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(85, 26);
            this.btnCancelar.TabIndex = 3;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            // 
            // btnAceptar
            // 
            this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAceptar.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnAceptar.Location = new System.Drawing.Point(590, 10);
            this.btnAceptar.Name = "btnAceptar";
            this.btnAceptar.Size = new System.Drawing.Size(85, 26);
            this.btnAceptar.TabIndex = 2;
            this.btnAceptar.Text = "Aceptar";
            this.btnAceptar.UseVisualStyleBackColor = true;
            // 
            // btnDesmarcarTodos
            // 
            this.btnDesmarcarTodos.Location = new System.Drawing.Point(120, 10);
            this.btnDesmarcarTodos.Name = "btnDesmarcarTodos";
            this.btnDesmarcarTodos.Size = new System.Drawing.Size(100, 26);
            this.btnDesmarcarTodos.TabIndex = 1;
            this.btnDesmarcarTodos.Text = "Desmarcar";
            this.btnDesmarcarTodos.UseVisualStyleBackColor = true;
            // 
            // btnMarcarTodos
            // 
            this.btnMarcarTodos.Location = new System.Drawing.Point(15, 10);
            this.btnMarcarTodos.Name = "btnMarcarTodos";
            this.btnMarcarTodos.Size = new System.Drawing.Size(95, 26);
            this.btnMarcarTodos.TabIndex = 0;
            this.btnMarcarTodos.Text = "Marcar Todos";
            this.btnMarcarTodos.UseVisualStyleBackColor = true;
            // 
            // dgvTerceros
            // 
            this.dgvTerceros.AllowUserToAddRows = false;
            this.dgvTerceros.AllowUserToDeleteRows = false;
            this.dgvTerceros.AutoGenerateColumns = false;
            this.dgvTerceros.AllowUserToResizeRows = false;
            this.dgvTerceros.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvTerceros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgvTerceros.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTerceros.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSeleccion,
            this.colNombre,
            this.colCedula,
            this.colTipo,
            this.colTerritorial,
            this.colCategoria});
            this.dgvTerceros.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvTerceros.Location = new System.Drawing.Point(0, 75);
            this.dgvTerceros.MultiSelect = false;
            this.dgvTerceros.Name = "dgvTerceros";
            this.dgvTerceros.RowHeadersVisible = false;
            this.dgvTerceros.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvTerceros.Size = new System.Drawing.Size(784, 391);
            this.dgvTerceros.TabIndex = 2;
            // 
            // colSeleccion
            // 
            this.colSeleccion.DataPropertyName = "Seleccionado";
            this.colSeleccion.HeaderText = "";
            this.colSeleccion.Name = "colSeleccion";
            this.colSeleccion.Width = 35;
            // 
            // colNombre
            // 
            this.colNombre.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colNombre.DataPropertyName = "DisplayName";
            this.colNombre.HeaderText = "Nombre y Residencia";
            this.colNombre.Name = "colNombre";
            this.colNombre.ReadOnly = true;
            // 
            // colCedula
            // 
            this.colCedula.DataPropertyName = "Cedula";
            this.colCedula.HeaderText = "Cédula";
            this.colCedula.Name = "colCedula";
            this.colCedula.ReadOnly = true;
            this.colCedula.Width = 100;
            // 
            // colTipo
            // 
            this.colTipo.DataPropertyName = "Tipo";
            this.colTipo.HeaderText = "Tipo";
            this.colTipo.Name = "colTipo";
            this.colTipo.ReadOnly = true;
            this.colTipo.Width = 110;
            // 
            // colTerritorial
            // 
            this.colTerritorial.DataPropertyName = "Subunidad";
            this.colTerritorial.HeaderText = "Territorial";
            this.colTerritorial.Name = "colTerritorial";
            this.colTerritorial.ReadOnly = true;
            this.colTerritorial.Width = 120;
            // 
            // colCategoria
            // 
            this.colCategoria.DataPropertyName = "Categoria";
            this.colCategoria.HeaderText = "Categoría";
            this.colCategoria.Name = "colCategoria";
            this.colCategoria.ReadOnly = true;
            this.colCategoria.Width = 100;
            // 
            // AgFormTerceros
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 511);
            // Fase 1437: Vinculación estándar de teclado (Enter para Aceptar, Escape para Cancelar)
            this.AcceptButton = this.btnAceptar;
            this.CancelButton = this.btnCancelar;
            this.Controls.Add(this.dgvTerceros);
            this.Controls.Add(this.pnlBottom);
            this.Controls.Add(this.pnlTop);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(650, 400);
            this.Name = "AgFormTerceros";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Selección de Terceros y Docentes";
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.pnlBottom.ResumeLayout(false);
            this.pnlBottom.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerceros)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblTerritorial;
        private System.Windows.Forms.ComboBox cmbTerritorial;
        private System.Windows.Forms.Label lblBuscar;
        private System.Windows.Forms.TextBox txtBuscar;
        private System.Windows.Forms.CheckBox chkSoloHabilitadas;
        private System.Windows.Forms.CheckBox chkMostrarInactivos;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnMarcarTodos;
        private System.Windows.Forms.Button btnDesmarcarTodos;
        private System.Windows.Forms.Button btnAceptar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Label lblConteo;
        private System.Windows.Forms.DataGridView dgvTerceros;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colSeleccion;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCedula;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTerritorial;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCategoria;
    }
}
