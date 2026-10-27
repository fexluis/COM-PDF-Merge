namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormConfig
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnCancelar;

        // Labels
        private System.Windows.Forms.Label lblPeriodos;
        private System.Windows.Forms.Label lblSubunidad;
        private System.Windows.Forms.Label lblIdTerritorial;
        private System.Windows.Forms.Label lblFechaInicio;
        private System.Windows.Forms.Label lblFechaFin;
        private System.Windows.Forms.Label lblCdp;
        private System.Windows.Forms.Label lblSalario;
        private System.Windows.Forms.Label lblViaticoA;
        private System.Windows.Forms.Label lblViaticoB;
        private System.Windows.Forms.Label lblPorcentajeA;
        private System.Windows.Forms.Label lblPorcentajeB;

        private System.Windows.Forms.Label lblDtNombre;
        private System.Windows.Forms.Label lblDtCargo;
        private System.Windows.Forms.Label lblCorpCedula;
        private System.Windows.Forms.Label lblCorpNombre;
        private System.Windows.Forms.Label lblCorpCargo;
        private System.Windows.Forms.Label lblCiudadBase;
        private System.Windows.Forms.Label lblDeptoBase;
        private System.Windows.Forms.Label lblAeropuertoBase;

        // Controls
        private System.Windows.Forms.ComboBox cmbPeriodos;
        private System.Windows.Forms.ComboBox cmbSubunidad;
        private System.Windows.Forms.TextBox txtIdTerritorial;
        private System.Windows.Forms.TextBox txtFechaInicio;
        private System.Windows.Forms.TextBox txtFechaFin;
        private System.Windows.Forms.TextBox txtCdp;
        private System.Windows.Forms.TextBox txtSalario;
        private System.Windows.Forms.TextBox txtViaticoA;
        private System.Windows.Forms.TextBox txtViaticoB;
        private System.Windows.Forms.TextBox txtPorcentajeA;
        private System.Windows.Forms.TextBox txtPorcentajeB;

        private System.Windows.Forms.TextBox txtDtNombre;
        private System.Windows.Forms.TextBox txtDtCargo;
        private System.Windows.Forms.TextBox txtCorpCedula;
        private System.Windows.Forms.TextBox txtCorpNombre;
        private System.Windows.Forms.TextBox txtCorpCargo;
        private System.Windows.Forms.ComboBox cmbCiudadBase;
        private System.Windows.Forms.TextBox txtDeptoBase;
        private System.Windows.Forms.ComboBox cmbAeropuertoBase;

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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();

            this.lblPeriodos = new System.Windows.Forms.Label();
            this.lblSubunidad = new System.Windows.Forms.Label();
            this.lblIdTerritorial = new System.Windows.Forms.Label();
            this.lblFechaInicio = new System.Windows.Forms.Label();
            this.lblFechaFin = new System.Windows.Forms.Label();
            this.lblCdp = new System.Windows.Forms.Label();
            this.lblSalario = new System.Windows.Forms.Label();
            this.lblViaticoA = new System.Windows.Forms.Label();
            this.lblViaticoB = new System.Windows.Forms.Label();
            this.lblPorcentajeA = new System.Windows.Forms.Label();
            this.lblPorcentajeB = new System.Windows.Forms.Label();

            this.lblDtNombre = new System.Windows.Forms.Label();
            this.lblDtCargo = new System.Windows.Forms.Label();
            this.lblCorpCedula = new System.Windows.Forms.Label();
            this.lblCorpNombre = new System.Windows.Forms.Label();
            this.lblCorpCargo = new System.Windows.Forms.Label();
            this.lblCiudadBase = new System.Windows.Forms.Label();
            this.lblDeptoBase = new System.Windows.Forms.Label();
            this.lblAeropuertoBase = new System.Windows.Forms.Label();

            this.cmbPeriodos = new System.Windows.Forms.ComboBox();
            this.cmbSubunidad = new System.Windows.Forms.ComboBox();
            this.txtIdTerritorial = new System.Windows.Forms.TextBox();
            this.txtFechaInicio = new System.Windows.Forms.TextBox();
            this.txtFechaFin = new System.Windows.Forms.TextBox();
            this.txtCdp = new System.Windows.Forms.TextBox();
            this.txtSalario = new System.Windows.Forms.TextBox();
            this.txtViaticoA = new System.Windows.Forms.TextBox();
            this.txtViaticoB = new System.Windows.Forms.TextBox();
            this.txtPorcentajeA = new System.Windows.Forms.TextBox();
            this.txtPorcentajeB = new System.Windows.Forms.TextBox();

            this.txtDtNombre = new System.Windows.Forms.TextBox();
            this.txtDtCargo = new System.Windows.Forms.TextBox();
            this.txtCorpCedula = new System.Windows.Forms.TextBox();
            this.txtCorpNombre = new System.Windows.Forms.TextBox();
            this.txtCorpCargo = new System.Windows.Forms.TextBox();
            this.cmbCiudadBase = new System.Windows.Forms.ComboBox();
            this.txtDeptoBase = new System.Windows.Forms.TextBox();
            this.cmbAeropuertoBase = new System.Windows.Forms.ComboBox();

            this.panel1.SuspendLayout();
            this.SuspendLayout();

            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnCancelar);
            this.panel1.Controls.Add(this.btnGuardar);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(684, 50);
            this.panel1.TabIndex = 1;

            // btnGuardar
            this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGuardar.Location = new System.Drawing.Point(516, 15);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 0;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);

            // btnCancelar
            this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancelar.Location = new System.Drawing.Point(597, 15);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 1;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);

            int xL1 = 20, xC1 = 150, y = 20, dy = 30;
            int xL2 = 340, xC2 = 470;

            // Col 1
            this.lblPeriodos.Location = new System.Drawing.Point(xL1, y); this.lblPeriodos.Text = "Periodo:"; this.lblPeriodos.AutoSize = true;
            this.cmbPeriodos.Location = new System.Drawing.Point(xC1, y); this.cmbPeriodos.Size = new System.Drawing.Size(160, 21); this.cmbPeriodos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; y += dy;

            this.lblSubunidad.Location = new System.Drawing.Point(xL1, y); this.lblSubunidad.Text = "Subunidad:"; this.lblSubunidad.AutoSize = true;
            this.cmbSubunidad.Location = new System.Drawing.Point(xC1, y); this.cmbSubunidad.Size = new System.Drawing.Size(160, 21); this.cmbSubunidad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; y += dy;

            this.lblIdTerritorial.Location = new System.Drawing.Point(xL1, y); this.lblIdTerritorial.Text = "ID Territorial:"; this.lblIdTerritorial.AutoSize = true;
            this.txtIdTerritorial.Location = new System.Drawing.Point(xC1, y); this.txtIdTerritorial.Size = new System.Drawing.Size(160, 20); this.txtIdTerritorial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosEnteros_KeyPress); this.txtIdTerritorial.ReadOnly = true; y += dy;

            this.lblFechaInicio.Location = new System.Drawing.Point(xL1, y); this.lblFechaInicio.Text = "Fecha Inicio:"; this.lblFechaInicio.AutoSize = true;
            this.txtFechaInicio.Location = new System.Drawing.Point(xC1, y); this.txtFechaInicio.Size = new System.Drawing.Size(160, 20); y += dy;

            this.lblFechaFin.Location = new System.Drawing.Point(xL1, y); this.lblFechaFin.Text = "Fecha Fin:"; this.lblFechaFin.AutoSize = true;
            this.txtFechaFin.Location = new System.Drawing.Point(xC1, y); this.txtFechaFin.Size = new System.Drawing.Size(160, 20); y += dy;

            this.lblCdp.Location = new System.Drawing.Point(xL1, y); this.lblCdp.Text = "CDP:"; this.lblCdp.AutoSize = true;
            this.txtCdp.Location = new System.Drawing.Point(xC1, y); this.txtCdp.Size = new System.Drawing.Size(160, 20); y += dy;

            this.lblSalario.Location = new System.Drawing.Point(xL1, y); this.lblSalario.Text = "Salario Decr. 613:"; this.lblSalario.AutoSize = true;
            this.txtSalario.Location = new System.Drawing.Point(xC1, y); this.txtSalario.Size = new System.Drawing.Size(160, 20); this.txtSalario.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosEnteros_KeyPress); y += dy;

            this.lblViaticoA.Location = new System.Drawing.Point(xL1, y); this.lblViaticoA.Text = "Viático Tarifa A:"; this.lblViaticoA.AutoSize = true;
            this.txtViaticoA.Location = new System.Drawing.Point(xC1, y); this.txtViaticoA.Size = new System.Drawing.Size(160, 20); this.txtViaticoA.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosEnteros_KeyPress); y += dy;

            this.lblViaticoB.Location = new System.Drawing.Point(xL1, y); this.lblViaticoB.Text = "Viático Tarifa B:"; this.lblViaticoB.AutoSize = true;
            this.txtViaticoB.Location = new System.Drawing.Point(xC1, y); this.txtViaticoB.Size = new System.Drawing.Size(160, 20); this.txtViaticoB.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosEnteros_KeyPress); y += dy;

            this.lblPorcentajeA.Location = new System.Drawing.Point(xL1, y); this.lblPorcentajeA.Text = "Porcentaje A:"; this.lblPorcentajeA.AutoSize = true;
            this.txtPorcentajeA.Location = new System.Drawing.Point(xC1, y); this.txtPorcentajeA.Size = new System.Drawing.Size(160, 20); this.txtPorcentajeA.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosDecimales_KeyPress); this.txtPorcentajeA.Leave += new System.EventHandler(this.Decimal_Leave); y += dy;

            this.lblPorcentajeB.Location = new System.Drawing.Point(xL1, y); this.lblPorcentajeB.Text = "Porcentaje B:"; this.lblPorcentajeB.AutoSize = true;
            this.txtPorcentajeB.Location = new System.Drawing.Point(xC1, y); this.txtPorcentajeB.Size = new System.Drawing.Size(160, 20); this.txtPorcentajeB.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumerosDecimales_KeyPress); this.txtPorcentajeB.Leave += new System.EventHandler(this.Decimal_Leave); y += dy;

            // Col 2
            y = 20;
            this.lblDtNombre.Location = new System.Drawing.Point(xL2, y); this.lblDtNombre.Text = "Dir. Terr. Nombre:"; this.lblDtNombre.AutoSize = true;
            this.txtDtNombre.Location = new System.Drawing.Point(xC2, y); this.txtDtNombre.Size = new System.Drawing.Size(190, 20); y += dy;

            this.lblDtCargo.Location = new System.Drawing.Point(xL2, y); this.lblDtCargo.Text = "Dir. Terr. Cargo:"; this.lblDtCargo.AutoSize = true;
            this.txtDtCargo.Location = new System.Drawing.Point(xC2, y); this.txtDtCargo.Size = new System.Drawing.Size(190, 20); y += dy;

            this.lblCorpCedula.Location = new System.Drawing.Point(xL2, y); this.lblCorpCedula.Text = "Dir. Terr. Cédula:"; this.lblCorpCedula.AutoSize = true;
            this.txtCorpCedula.Location = new System.Drawing.Point(xC2, y); this.txtCorpCedula.Size = new System.Drawing.Size(190, 20); y += dy;

            this.lblCorpNombre.Location = new System.Drawing.Point(xL2, y); this.lblCorpNombre.Text = "Corporativo:"; this.lblCorpNombre.AutoSize = true;
            this.txtCorpNombre.Location = new System.Drawing.Point(xC2, y); this.txtCorpNombre.Size = new System.Drawing.Size(190, 20); y += dy;

            this.lblCorpCargo.Location = new System.Drawing.Point(xL2, y); this.lblCorpCargo.Text = "Corp. Cargo:"; this.lblCorpCargo.AutoSize = true;
            this.txtCorpCargo.Location = new System.Drawing.Point(xC2, y); this.txtCorpCargo.Size = new System.Drawing.Size(190, 20); y += dy;

            this.lblCiudadBase.Location = new System.Drawing.Point(xL2, y); this.lblCiudadBase.Text = "Ciudad Base:"; this.lblCiudadBase.AutoSize = true;
            this.cmbCiudadBase.Location = new System.Drawing.Point(xC2, y); this.cmbCiudadBase.Size = new System.Drawing.Size(190, 21); this.cmbCiudadBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; this.cmbCiudadBase.SelectedIndexChanged += new System.EventHandler(this.cmbCiudadBase_SelectedIndexChanged); y += dy;

            this.lblDeptoBase.Location = new System.Drawing.Point(xL2, y); this.lblDeptoBase.Text = "Depto Base:"; this.lblDeptoBase.AutoSize = true;
            this.txtDeptoBase.Location = new System.Drawing.Point(xC2, y); this.txtDeptoBase.Size = new System.Drawing.Size(190, 20); this.txtDeptoBase.ReadOnly = true; y += dy;

            this.lblAeropuertoBase.Location = new System.Drawing.Point(xL2, y); this.lblAeropuertoBase.Text = "Aeropuerto Base:"; this.lblAeropuertoBase.AutoSize = true;
            this.cmbAeropuertoBase.Location = new System.Drawing.Point(xC2, y); this.cmbAeropuertoBase.Size = new System.Drawing.Size(190, 21); this.cmbAeropuertoBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; y += dy;


            // FormConfig
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 451);
            
            this.Controls.Add(this.lblPeriodos); this.Controls.Add(this.cmbPeriodos);
            this.Controls.Add(this.lblSubunidad); this.Controls.Add(this.cmbSubunidad);
            this.Controls.Add(this.lblIdTerritorial); this.Controls.Add(this.txtIdTerritorial);
            this.Controls.Add(this.lblFechaInicio); this.Controls.Add(this.txtFechaInicio);
            this.Controls.Add(this.lblFechaFin); this.Controls.Add(this.txtFechaFin);
            this.Controls.Add(this.lblCdp); this.Controls.Add(this.txtCdp);
            this.Controls.Add(this.lblSalario); this.Controls.Add(this.txtSalario);
            this.Controls.Add(this.lblViaticoA); this.Controls.Add(this.txtViaticoA);
            this.Controls.Add(this.lblViaticoB); this.Controls.Add(this.txtViaticoB);
            this.Controls.Add(this.lblPorcentajeA); this.Controls.Add(this.txtPorcentajeA);
            this.Controls.Add(this.lblPorcentajeB); this.Controls.Add(this.txtPorcentajeB);

            this.Controls.Add(this.lblDtNombre); this.Controls.Add(this.txtDtNombre);
            this.Controls.Add(this.lblDtCargo); this.Controls.Add(this.txtDtCargo);
            this.Controls.Add(this.lblCorpCedula); this.Controls.Add(this.txtCorpCedula);
            this.Controls.Add(this.lblCorpNombre); this.Controls.Add(this.txtCorpNombre);
            this.Controls.Add(this.lblCorpCargo); this.Controls.Add(this.txtCorpCargo);
            this.Controls.Add(this.lblCiudadBase); this.Controls.Add(this.cmbCiudadBase);
            this.Controls.Add(this.lblDeptoBase); this.Controls.Add(this.txtDeptoBase);
            this.Controls.Add(this.lblAeropuertoBase); this.Controls.Add(this.cmbAeropuertoBase);

            this.Controls.Add(this.panel1);
            this.Name = "FormConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Configuración";
            this.Load += new System.EventHandler(this.FormConfig_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}
