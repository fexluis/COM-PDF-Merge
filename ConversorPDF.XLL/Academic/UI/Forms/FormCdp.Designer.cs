namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormCdp
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView dgvCdp;
        private System.Windows.Forms.Panel pnlDetails;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnLimpiar;
        private System.Windows.Forms.Button btnEliminar;

        private System.Windows.Forms.Label lblPeriodo;
        private System.Windows.Forms.ComboBox cmbPeriodos;
        private System.Windows.Forms.Label lblCdpNum;
        private System.Windows.Forms.TextBox txtCdpNum;
        private System.Windows.Forms.Label lblFechaCdp;
        private System.Windows.Forms.TextBox txtFechaCdp;
        private System.Windows.Forms.Label lblRubro;
        private System.Windows.Forms.TextBox txtRubro;
        private System.Windows.Forms.Label lblResponsable;
        private System.Windows.Forms.TextBox txtResponsable;
        private System.Windows.Forms.Label lblCargo;
        private System.Windows.Forms.TextBox txtCargo;
        private System.Windows.Forms.Label lblElaboro;
        private System.Windows.Forms.TextBox txtElaboro;
        private System.Windows.Forms.Label lblReviso;
        private System.Windows.Forms.TextBox txtReviso;
        private System.Windows.Forms.Label lblDependencia;
        private System.Windows.Forms.ComboBox cmbDependencia;
        private System.Windows.Forms.Label lblDependenciaFull;
        private System.Windows.Forms.TextBox txtDependenciaFull;
        private System.Windows.Forms.Label lblActivo;
        private System.Windows.Forms.ComboBox cmbActivo;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// Se libera explícitamente la conexión a la base de datos (Fase 4: prevención de Resource Leak).
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Fase 4: Limpiar conexión generada para esta instancia
                _currentConnection?.Dispose();

                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblFiltroPeriodo;
        private System.Windows.Forms.ComboBox cmbPeriodosFiltro;

        private void InitializeComponent()
        {
            this.dgvCdp = new System.Windows.Forms.DataGridView();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.pnlTop = new System.Windows.Forms.Panel();

            this.lblFiltroPeriodo = new System.Windows.Forms.Label();
            this.cmbPeriodosFiltro = new System.Windows.Forms.ComboBox();

            this.lblPeriodo = new System.Windows.Forms.Label();
            this.cmbPeriodos = new System.Windows.Forms.ComboBox();
            this.lblCdpNum = new System.Windows.Forms.Label();
            this.txtCdpNum = new System.Windows.Forms.TextBox();
            this.lblFechaCdp = new System.Windows.Forms.Label();
            this.txtFechaCdp = new System.Windows.Forms.TextBox();
            this.lblRubro = new System.Windows.Forms.Label();
            this.txtRubro = new System.Windows.Forms.TextBox();
            this.lblResponsable = new System.Windows.Forms.Label();
            this.txtResponsable = new System.Windows.Forms.TextBox();
            this.lblCargo = new System.Windows.Forms.Label();
            this.txtCargo = new System.Windows.Forms.TextBox();
            this.lblElaboro = new System.Windows.Forms.Label();
            this.txtElaboro = new System.Windows.Forms.TextBox();
            this.lblReviso = new System.Windows.Forms.Label();
            this.txtReviso = new System.Windows.Forms.TextBox();
            this.lblDependencia = new System.Windows.Forms.Label();
            this.cmbDependencia = new System.Windows.Forms.ComboBox();
            this.lblDependenciaFull = new System.Windows.Forms.Label();
            this.txtDependenciaFull = new System.Windows.Forms.TextBox();
            this.lblActivo = new System.Windows.Forms.Label();
            this.cmbActivo = new System.Windows.Forms.ComboBox();

            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnLimpiar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnNuevo = new System.Windows.Forms.Button();

            ((System.ComponentModel.ISupportInitialize)(this.dgvCdp)).BeginInit();
            this.pnlDetails.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.SuspendLayout();

            // pnlTop
            this.pnlTop.Controls.Add(this.lblFiltroPeriodo);
            this.pnlTop.Controls.Add(this.cmbPeriodosFiltro);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(700, 40);
            this.pnlTop.TabIndex = 2;

            // lblFiltroPeriodo
            this.lblFiltroPeriodo.AutoSize = true;
            this.lblFiltroPeriodo.Location = new System.Drawing.Point(10, 14);
            this.lblFiltroPeriodo.Name = "lblFiltroPeriodo";
            this.lblFiltroPeriodo.Size = new System.Drawing.Size(95, 13);
            this.lblFiltroPeriodo.TabIndex = 0;
            this.lblFiltroPeriodo.Text = "Filtrar por Período:";

            // cmbPeriodosFiltro
            this.cmbPeriodosFiltro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPeriodosFiltro.Location = new System.Drawing.Point(110, 10);
            this.cmbPeriodosFiltro.Name = "cmbPeriodosFiltro";
            this.cmbPeriodosFiltro.Size = new System.Drawing.Size(120, 21);
            this.cmbPeriodosFiltro.TabIndex = 1;
            this.cmbPeriodosFiltro.SelectedIndexChanged += new System.EventHandler(this.cmbPeriodosFiltro_SelectedIndexChanged);

            // dgvCdp
            this.dgvCdp.AllowUserToAddRows = false;
            this.dgvCdp.AllowUserToDeleteRows = false;
            this.dgvCdp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCdp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvCdp.Location = new System.Drawing.Point(0, 40);
            this.dgvCdp.MultiSelect = false;
            this.dgvCdp.Name = "dgvCdp";
            this.dgvCdp.ReadOnly = true;
            this.dgvCdp.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvCdp.Size = new System.Drawing.Size(300, 450);
            this.dgvCdp.TabIndex = 0;
            this.dgvCdp.SelectionChanged += new System.EventHandler(this.dgvCdp_SelectionChanged);

            // pnlDetails
            this.pnlDetails.Controls.Add(this.btnNuevo);
            this.pnlDetails.Controls.Add(this.lblPeriodo);
            this.pnlDetails.Controls.Add(this.cmbPeriodos);
            this.pnlDetails.Controls.Add(this.lblCdpNum);
            this.pnlDetails.Controls.Add(this.txtCdpNum);
            this.pnlDetails.Controls.Add(this.lblFechaCdp);
            this.pnlDetails.Controls.Add(this.txtFechaCdp);
            this.pnlDetails.Controls.Add(this.lblRubro);
            this.pnlDetails.Controls.Add(this.txtRubro);
            this.pnlDetails.Controls.Add(this.lblResponsable);
            this.pnlDetails.Controls.Add(this.txtResponsable);
            this.pnlDetails.Controls.Add(this.lblCargo);
            this.pnlDetails.Controls.Add(this.txtCargo);
            this.pnlDetails.Controls.Add(this.lblElaboro);
            this.pnlDetails.Controls.Add(this.txtElaboro);
            this.pnlDetails.Controls.Add(this.lblReviso);
            this.pnlDetails.Controls.Add(this.txtReviso);
            this.pnlDetails.Controls.Add(this.lblDependencia);
            this.pnlDetails.Controls.Add(this.cmbDependencia);
            this.pnlDetails.Controls.Add(this.lblDependenciaFull);
            this.pnlDetails.Controls.Add(this.txtDependenciaFull);
            this.pnlDetails.Controls.Add(this.lblActivo);
            this.pnlDetails.Controls.Add(this.cmbActivo);
            this.pnlDetails.Controls.Add(this.btnEliminar);
            this.pnlDetails.Controls.Add(this.btnLimpiar);
            this.pnlDetails.Controls.Add(this.btnGuardar);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Right;
            this.pnlDetails.Location = new System.Drawing.Point(700, 0);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(350, 450);
            this.pnlDetails.TabIndex = 1;

            // Labels and Controls setup
            int y = 20;
            int hGap = 30;

            // Buttons
            this.btnNuevo.Location = new System.Drawing.Point(10, y);
            this.btnNuevo.Size = new System.Drawing.Size(70, 23);
            this.btnNuevo.Text = "Nuevo";
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            y += hGap;

            this.lblPeriodo.Location = new System.Drawing.Point(10, y + 3);
            this.lblPeriodo.Size = new System.Drawing.Size(70, 13);
            this.lblPeriodo.Text = "Periodo:";
            this.cmbPeriodos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPeriodos.Location = new System.Drawing.Point(90, y);
            this.cmbPeriodos.Size = new System.Drawing.Size(190, 21);
            y += hGap;

            this.lblCdpNum.Location = new System.Drawing.Point(10, y + 3);
            this.lblCdpNum.Size = new System.Drawing.Size(70, 13);
            this.lblCdpNum.Text = "CDP Num:";
            this.txtCdpNum.Location = new System.Drawing.Point(90, y);
            this.txtCdpNum.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblFechaCdp.Location = new System.Drawing.Point(10, y + 3);
            this.lblFechaCdp.Size = new System.Drawing.Size(70, 13);
            this.lblFechaCdp.Text = "Fecha CDP:";
            this.txtFechaCdp.Location = new System.Drawing.Point(90, y);
            this.txtFechaCdp.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblRubro.Location = new System.Drawing.Point(10, y + 3);
            this.lblRubro.Size = new System.Drawing.Size(70, 13);
            this.lblRubro.Text = "Rubro:";
            this.txtRubro.Location = new System.Drawing.Point(90, y);
            this.txtRubro.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblResponsable.Location = new System.Drawing.Point(10, y + 3);
            this.lblResponsable.Size = new System.Drawing.Size(80, 13);
            this.lblResponsable.Text = "Responsable:";
            this.txtResponsable.Location = new System.Drawing.Point(90, y);
            this.txtResponsable.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblCargo.Location = new System.Drawing.Point(10, y + 3);
            this.lblCargo.Size = new System.Drawing.Size(70, 13);
            this.lblCargo.Text = "Cargo:";
            this.txtCargo.Location = new System.Drawing.Point(90, y);
            this.txtCargo.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblElaboro.Location = new System.Drawing.Point(10, y + 3);
            this.lblElaboro.Size = new System.Drawing.Size(70, 13);
            this.lblElaboro.Text = "Elaboró:";
            this.txtElaboro.Location = new System.Drawing.Point(90, y);
            this.txtElaboro.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblReviso.Location = new System.Drawing.Point(10, y + 3);
            this.lblReviso.Size = new System.Drawing.Size(70, 13);
            this.lblReviso.Text = "Revisó:";
            this.txtReviso.Location = new System.Drawing.Point(90, y);
            this.txtReviso.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblDependencia.Location = new System.Drawing.Point(10, y + 3);
            this.lblDependencia.Size = new System.Drawing.Size(80, 13);
            this.lblDependencia.Text = "Dependencia:";
            this.cmbDependencia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDependencia.Location = new System.Drawing.Point(90, y);
            this.cmbDependencia.Size = new System.Drawing.Size(190, 21);
            y += hGap;

            this.lblDependenciaFull.Location = new System.Drawing.Point(10, y + 3);
            this.lblDependenciaFull.Size = new System.Drawing.Size(80, 13);
            this.lblDependenciaFull.Text = "Dep. Full:";
            this.txtDependenciaFull.Location = new System.Drawing.Point(90, y);
            this.txtDependenciaFull.Size = new System.Drawing.Size(190, 20);
            y += hGap;

            this.lblActivo.Location = new System.Drawing.Point(10, y + 3);
            this.lblActivo.Size = new System.Drawing.Size(70, 13);
            this.lblActivo.Text = "Activo:";
            this.cmbActivo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbActivo.Location = new System.Drawing.Point(90, y);
            this.cmbActivo.Size = new System.Drawing.Size(190, 21);
            y += hGap + 10;

            // Buttons
            this.btnGuardar.Location = new System.Drawing.Point(10, y);
            this.btnGuardar.Size = new System.Drawing.Size(70, 23);
            this.btnGuardar.Text = "Agregar";
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);

            this.btnLimpiar.Location = new System.Drawing.Point(85, y);
            this.btnLimpiar.Size = new System.Drawing.Size(70, 23);
            this.btnLimpiar.Text = "Limpiar";
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);

            this.btnEliminar.Location = new System.Drawing.Point(160, y);
            this.btnEliminar.Size = new System.Drawing.Size(70, 23);
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);

            // FormCdp
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 450);
            this.Controls.Add(this.dgvCdp);
            this.Controls.Add(this.pnlTop);
            this.Controls.Add(this.pnlDetails);
            this.Name = "FormCdp";
            this.Text = "Gestión de CDP Dependencias";
            this.Load += new System.EventHandler(this.FormCdp_Load);

            ((System.ComponentModel.ISupportInitialize)(this.dgvCdp)).EndInit();
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
