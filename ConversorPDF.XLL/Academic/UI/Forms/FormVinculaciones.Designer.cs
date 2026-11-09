using System;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormVinculaciones
    {
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            int anchuraForm = 1105;
            this.pnlFilters = new System.Windows.Forms.Panel();
            this.btnVerVinculaciones = new System.Windows.Forms.Button();
            this.btnAdicionar = new System.Windows.Forms.Button();
            this.btnRevisar = new System.Windows.Forms.Button();
            this.btnLimpiar = new System.Windows.Forms.Button();
            this.btnFiltrar = new System.Windows.Forms.Button();
            this.chkSoloViaticas = new System.Windows.Forms.CheckBox();
            this.cmbSemestres = new System.Windows.Forms.ComboBox();
            this.lblSemestres = new System.Windows.Forms.Label();
            this.cmbProgramas = new System.Windows.Forms.ComboBox();
            this.lblProgramas = new System.Windows.Forms.Label();
            this.cmbTercero = new System.Windows.Forms.ComboBox();
            this.lblTercero = new System.Windows.Forms.Label();
            this.cmbSede = new System.Windows.Forms.ComboBox();
            this.lblSede = new System.Windows.Forms.Label();
            this.lblFechaFinal = new System.Windows.Forms.Label();
            this.dtpFechaFinal = new System.Windows.Forms.DateTimePicker();
            this.lblFechaInicio = new System.Windows.Forms.Label();
            this.dtpFechaInicio = new System.Windows.Forms.DateTimePicker();
            this.cmbTerritoriales = new System.Windows.Forms.ComboBox();
            this.cmbPeriodo = new System.Windows.Forms.ComboBox();
            this.lblCorrida = new System.Windows.Forms.Label();
            this.cmbCorrida = new System.Windows.Forms.ComboBox();
            this.btnAjustarFiltrosCorrida = new System.Windows.Forms.Button();
            this.btnEliminarCorrida = new System.Windows.Forms.Button();
            this.lblModo = new System.Windows.Forms.Label();

            this.dgvVinculaciones = new System.Windows.Forms.DataGridView();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblUsuarioActivo = new System.Windows.Forms.ToolStripStatusLabel();

            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.sedesAcadémicasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tercerosDocentesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.díasFestivosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gastosTerrestresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dependenciasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.terminalesAéreasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.configuraciónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.períodosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();

            this.pnlFilters.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVinculaciones)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();

            // pnlFilters
            this.pnlFilters.Controls.Add(this.btnVerVinculaciones);
            this.pnlFilters.Controls.Add(this.btnAdicionar);
            this.pnlFilters.Controls.Add(this.btnRevisar);
            this.pnlFilters.Controls.Add(this.btnLimpiar);
            this.pnlFilters.Controls.Add(this.btnFiltrar);
            this.pnlFilters.Controls.Add(this.chkSoloViaticas);
            this.pnlFilters.Controls.Add(this.cmbSemestres);
            this.pnlFilters.Controls.Add(this.lblSemestres);
            this.pnlFilters.Controls.Add(this.cmbProgramas);
            this.pnlFilters.Controls.Add(this.lblProgramas);
            this.pnlFilters.Controls.Add(this.cmbTercero);
            this.pnlFilters.Controls.Add(this.lblTercero);
            this.pnlFilters.Controls.Add(this.cmbSede);
            this.pnlFilters.Controls.Add(this.lblSede);
            this.pnlFilters.Controls.Add(this.lblFechaFinal);
            this.pnlFilters.Controls.Add(this.dtpFechaFinal);
            this.pnlFilters.Controls.Add(this.lblFechaInicio);
            this.pnlFilters.Controls.Add(this.dtpFechaInicio);
            this.pnlFilters.Controls.Add(this.cmbTerritoriales);
            this.pnlFilters.Controls.Add(this.cmbPeriodo);
            this.pnlFilters.Controls.Add(this.lblCorrida);
            this.pnlFilters.Controls.Add(this.cmbCorrida);
            this.pnlFilters.Controls.Add(this.btnAjustarFiltrosCorrida);
            this.pnlFilters.Controls.Add(this.btnEliminarCorrida);
            this.pnlFilters.Controls.Add(this.lblModo);
            this.pnlFilters.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlFilters.Location = new System.Drawing.Point(0, 24);
            this.pnlFilters.Name = "pnlFilters";
            this.pnlFilters.MinimumSize = new System.Drawing.Size(0, 104);
            this.pnlFilters.Size = new System.Drawing.Size(anchuraForm, 104);
            this.pnlFilters.TabIndex = 0;

            // btnVerVinculaciones
            this.btnVerVinculaciones.Location = new System.Drawing.Point(920, 45);
            this.btnVerVinculaciones.Name = "btnVerVinculaciones";
            this.btnVerVinculaciones.Size = new System.Drawing.Size(90, 23);
            this.btnVerVinculaciones.TabIndex = 18;
            this.btnVerVinculaciones.Text = "Vinculaciones";
            this.btnVerVinculaciones.UseVisualStyleBackColor = true;
            this.btnVerVinculaciones.Click += new System.EventHandler(this.btnVerVinculaciones_Click);

            // btnAdicionar
            this.btnAdicionar.Location = new System.Drawing.Point(920, 15);
            this.btnAdicionar.Name = "btnAdicionar";
            this.btnAdicionar.Size = new System.Drawing.Size(90, 23);
            this.btnAdicionar.TabIndex = 10;
            this.btnAdicionar.Text = "Adicionar";
            this.btnAdicionar.UseVisualStyleBackColor = true;
            this.btnAdicionar.Click += new System.EventHandler(this.btnAdicionar_Click);

            // btnRevisar
            this.btnRevisar.Location = new System.Drawing.Point(835, 45);
            this.btnRevisar.Name = "btnRevisar";
            this.btnRevisar.Size = new System.Drawing.Size(75, 23);
            this.btnRevisar.TabIndex = 9;
            this.btnRevisar.Text = "Revisar";
            this.btnRevisar.UseVisualStyleBackColor = true;
            this.btnRevisar.Click += new System.EventHandler(this.btnRevisar_Click);

            // btnLimpiar
            this.btnLimpiar.Location = new System.Drawing.Point(760, 45);
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.Size = new System.Drawing.Size(60, 23);
            this.btnLimpiar.TabIndex = 6;
            this.btnLimpiar.Text = "Limpiar";
            this.btnLimpiar.UseVisualStyleBackColor = true;
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);

            // btnFiltrar
            this.btnFiltrar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFiltrar.Location = new System.Drawing.Point(835, 15);
            this.btnFiltrar.Name = "btnFiltrar";
            this.btnFiltrar.Size = new System.Drawing.Size(75, 23);
            this.btnFiltrar.TabIndex = 5;
            this.btnFiltrar.Text = "Filtrar";
            this.btnFiltrar.UseVisualStyleBackColor = true;
            this.btnFiltrar.Click += new System.EventHandler(this.btnFiltrar_Click);

            // chkSoloViaticas
            this.chkSoloViaticas.AutoSize = true;
            this.chkSoloViaticas.Location = new System.Drawing.Point(760, 17);
            this.chkSoloViaticas.Name = "chkSoloViaticas";
            this.chkSoloViaticas.Size = new System.Drawing.Size(65, 17);
            this.chkSoloViaticas.TabIndex = 4;
            this.chkSoloViaticas.Text = "Viáticos";
            this.chkSoloViaticas.UseVisualStyleBackColor = true;

            // cmbSemestres
            this.cmbSemestres.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestres.FormattingEnabled = true;
            this.cmbSemestres.Location = new System.Drawing.Point(75, 45);
            this.cmbSemestres.Name = "cmbSemestres";
            this.cmbSemestres.Size = new System.Drawing.Size(130, 21);
            this.cmbSemestres.TabIndex = 23;

            // lblSemestres
            this.lblSemestres.AutoSize = true;
            this.lblSemestres.Location = new System.Drawing.Point(12, 48);
            this.lblSemestres.Name = "lblSemestres";
            this.lblSemestres.Size = new System.Drawing.Size(54, 13);
            this.lblSemestres.TabIndex = 24;
            this.lblSemestres.Text = "Semestre:";

            // cmbProgramas
            this.cmbProgramas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbProgramas.FormattingEnabled = true;
            this.cmbProgramas.Location = new System.Drawing.Point(75, 15);
            this.cmbProgramas.Name = "cmbProgramas";
            this.cmbProgramas.Size = new System.Drawing.Size(130, 21);
            this.cmbProgramas.TabIndex = 19;

            // lblProgramas
            this.lblProgramas.AutoSize = true;
            this.lblProgramas.Location = new System.Drawing.Point(12, 18);
            this.lblProgramas.Name = "lblProgramas";
            this.lblProgramas.Size = new System.Drawing.Size(55, 13);
            this.lblProgramas.TabIndex = 20;
            this.lblProgramas.Text = "Programa:";

            // cmbTercero
            this.cmbTercero.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            this.cmbTercero.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbTercero.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbTercero.FormattingEnabled = true;
            this.cmbTercero.Location = new System.Drawing.Point(260, 45);
            this.cmbTercero.Name = "cmbTercero";
            this.cmbTercero.Size = new System.Drawing.Size(320, 21);
            this.cmbTercero.TabIndex = 3;

            // lblTercero
            this.lblTercero.AutoSize = true;
            this.lblTercero.Location = new System.Drawing.Point(210, 48);
            this.lblTercero.Name = "lblTercero";
            this.lblTercero.Size = new System.Drawing.Size(47, 13);
            this.lblTercero.TabIndex = 2;
            this.lblTercero.Text = "Tercero:";

            // cmbSede
            this.cmbSede.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSede.FormattingEnabled = true;
            this.cmbSede.Location = new System.Drawing.Point(260, 15);
            this.cmbSede.Name = "cmbSede";
            this.cmbSede.Size = new System.Drawing.Size(320, 21);
            this.cmbSede.TabIndex = 1;

            // lblSede
            this.lblSede.AutoSize = true;
            this.lblSede.Location = new System.Drawing.Point(220, 18);
            this.lblSede.Name = "lblSede";
            this.lblSede.Size = new System.Drawing.Size(35, 13);
            this.lblSede.TabIndex = 0;
            this.lblSede.Text = "Sede:";

            // lblFechaFinal
            this.lblFechaFinal.AutoSize = true;
            this.lblFechaFinal.Location = new System.Drawing.Point(595, 48);
            this.lblFechaFinal.Name = "lblFechaFinal";
            this.lblFechaFinal.Size = new System.Drawing.Size(32, 13);
            this.lblFechaFinal.TabIndex = 13;
            this.lblFechaFinal.Text = "Final:";

            // dtpFechaFinal
            this.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFechaFinal.Location = new System.Drawing.Point(640, 45);
            this.dtpFechaFinal.Name = "dtpFechaFinal";
            this.dtpFechaFinal.Size = new System.Drawing.Size(100, 20);
            this.dtpFechaFinal.TabIndex = 14;
            this.dtpFechaFinal.Value = new System.DateTime(2026, 12, 31, 0, 0, 0, 0);

            // lblFechaInicio
            this.lblFechaInicio.AutoSize = true;
            this.lblFechaInicio.Location = new System.Drawing.Point(595, 18);
            this.lblFechaInicio.Name = "lblFechaInicio";
            this.lblFechaInicio.Size = new System.Drawing.Size(35, 13);
            this.lblFechaInicio.TabIndex = 11;
            this.lblFechaInicio.Text = "Inicio:";

            // dtpFechaInicio
            this.dtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFechaInicio.Location = new System.Drawing.Point(640, 15);
            this.dtpFechaInicio.Name = "dtpFechaInicio";
            this.dtpFechaInicio.Size = new System.Drawing.Size(100, 20);
            this.dtpFechaInicio.TabIndex = 12;
            this.dtpFechaInicio.Value = new System.DateTime(2026, 1, 1, 0, 0, 0, 0);

            // cmbTerritoriales (Oculto)
            this.cmbTerritoriales.Location = new System.Drawing.Point(-100, -100);
            this.cmbTerritoriales.Name = "cmbTerritoriales";
            this.cmbTerritoriales.Size = new System.Drawing.Size(10, 21);
            this.cmbTerritoriales.Visible = false;

            // cmbPeriodo (Oculto)
            this.cmbPeriodo.Location = new System.Drawing.Point(-100, -100);
            this.cmbPeriodo.Name = "cmbPeriodo";
            this.cmbPeriodo.Size = new System.Drawing.Size(10, 21);
            this.cmbPeriodo.Visible = false;

            // lblCorrida (fila 3 · visible solo en MODO AG)
            this.lblCorrida.AutoSize = true;
            this.lblCorrida.Location = new System.Drawing.Point(12, 78);
            this.lblCorrida.Name = "lblCorrida";
            this.lblCorrida.Size = new System.Drawing.Size(46, 13);
            this.lblCorrida.TabIndex = 25;
            this.lblCorrida.Text = "Corrida:";
            this.lblCorrida.Visible = false;

            // cmbCorrida (fila 3 · bajo Semestre y Tercero)
            this.cmbCorrida.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCorrida.FormattingEnabled = true;
            this.cmbCorrida.Location = new System.Drawing.Point(75, 75);
            this.cmbCorrida.Name = "cmbCorrida";
            this.cmbCorrida.Size = new System.Drawing.Size(505, 21);
            this.cmbCorrida.DropDownWidth = 640;
            this.cmbCorrida.TabIndex = 26;
            this.cmbCorrida.Visible = false;

            // btnAjustarFiltrosCorrida (fila 3 · antes del carril de botones x=760)
            this.btnAjustarFiltrosCorrida.Location = new System.Drawing.Point(585, 74);
            this.btnAjustarFiltrosCorrida.Name = "btnAjustarFiltrosCorrida";
            this.btnAjustarFiltrosCorrida.Size = new System.Drawing.Size(170, 23);
            this.btnAjustarFiltrosCorrida.TabIndex = 27;
            this.btnAjustarFiltrosCorrida.Text = "Ajustar filtros";
            this.btnAjustarFiltrosCorrida.UseVisualStyleBackColor = true;
            this.btnAjustarFiltrosCorrida.Visible = false;

            // btnEliminarCorrida (fila 3 · junto a "Ajustar filtros")
            this.btnEliminarCorrida.Location = new System.Drawing.Point(765, 74);
            this.btnEliminarCorrida.Name = "btnEliminarCorrida";
            this.btnEliminarCorrida.Size = new System.Drawing.Size(140, 23);
            this.btnEliminarCorrida.TabIndex = 28;
            this.btnEliminarCorrida.Text = "Eliminar corrida";
            this.btnEliminarCorrida.UseVisualStyleBackColor = true;
            this.btnEliminarCorrida.Visible = false;

            // lblModo (etiqueta de modo · fila 3 en AG, esquina superior en Producción)
            this.lblModo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblModo.AutoSize = true;
            this.lblModo.Location = new System.Drawing.Point(920, 78);
            this.lblModo.Name = "lblModo";
            this.lblModo.Size = new System.Drawing.Size(60, 13);
            this.lblModo.TabIndex = 29;
            this.lblModo.Text = "MODO AG (staging)";

            // dgvVinculaciones
            this.dgvVinculaciones.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvVinculaciones.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvVinculaciones.Location = new System.Drawing.Point(0, 104);
            this.dgvVinculaciones.Name = "dgvVinculaciones";
            this.dgvVinculaciones.Size = new System.Drawing.Size(anchuraForm, 474);
            this.dgvVinculaciones.TabIndex = 1;

            // statusStrip1
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus,
            this.lblUsuarioActivo});
            this.statusStrip1.Location = new System.Drawing.Point(0, 578);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(anchuraForm, 22);
            this.statusStrip1.TabIndex = 2;
            this.statusStrip1.Text = "statusStrip1";

            // lblStatus
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(39, 17);
            this.lblStatus.Text = "Listo.";
            this.lblStatus.Spring = true;
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            // lblUsuarioActivo
            this.lblUsuarioActivo.Name = "lblUsuarioActivo";
            this.lblUsuarioActivo.Size = new System.Drawing.Size(0, 17);
            this.lblUsuarioActivo.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;

            // menuStrip1
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sedesAcadémicasToolStripMenuItem,
            this.tercerosDocentesToolStripMenuItem,
            this.díasFestivosToolStripMenuItem,
            this.gastosTerrestresToolStripMenuItem,
            this.dependenciasToolStripMenuItem,
            this.terminalesAéreasToolStripMenuItem,
            this.períodosToolStripMenuItem,
            this.configuraciónToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(anchuraForm, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";

            // sedesAcadémicasToolStripMenuItem
            this.sedesAcadémicasToolStripMenuItem.Name = "sedesAcadémicasToolStripMenuItem";
            this.sedesAcadémicasToolStripMenuItem.Size = new System.Drawing.Size(115, 20);
            this.sedesAcadémicasToolStripMenuItem.Text = "Sedes Académicas";
            this.sedesAcadémicasToolStripMenuItem.Click += new System.EventHandler(this.btnSedes_Click);

            // tercerosDocentesToolStripMenuItem
            this.tercerosDocentesToolStripMenuItem.Name = "tercerosDocentesToolStripMenuItem";
            this.tercerosDocentesToolStripMenuItem.Size = new System.Drawing.Size(121, 20);
            this.tercerosDocentesToolStripMenuItem.Text = "Terceros (Docentes)";
            this.tercerosDocentesToolStripMenuItem.Click += new System.EventHandler(this.btnTerceros_Click);

            // díasFestivosToolStripMenuItem
            this.díasFestivosToolStripMenuItem.Name = "díasFestivosToolStripMenuItem";
            this.díasFestivosToolStripMenuItem.Size = new System.Drawing.Size(86, 20);
            this.díasFestivosToolStripMenuItem.Text = "Días Festivos";
            this.díasFestivosToolStripMenuItem.Click += new System.EventHandler(this.btnFestivos_Click);

            // períodosToolStripMenuItem
            this.períodosToolStripMenuItem.Name = "períodosToolStripMenuItem";
            this.períodosToolStripMenuItem.Size = new System.Drawing.Size(66, 20);
            this.períodosToolStripMenuItem.Text = "Períodos";
            this.períodosToolStripMenuItem.Click += new System.EventHandler(this.btnPeriodos_Click);

            // gastosTerrestresToolStripMenuItem
            this.gastosTerrestresToolStripMenuItem.Name = "gastosTerrestresToolStripMenuItem";
            this.gastosTerrestresToolStripMenuItem.Size = new System.Drawing.Size(107, 20);
            this.gastosTerrestresToolStripMenuItem.Text = "Gastos Terrestres";
            this.gastosTerrestresToolStripMenuItem.Click += new System.EventHandler(this.btnGastosTerrestres_Click);

            // dependenciasToolStripMenuItem
            this.dependenciasToolStripMenuItem.Name = "dependenciasToolStripMenuItem";
            this.dependenciasToolStripMenuItem.Size = new System.Drawing.Size(93, 20);
            this.dependenciasToolStripMenuItem.Text = "Dependencias";
            this.dependenciasToolStripMenuItem.Click += new System.EventHandler(this.btnCdp_Click);

            // terminalesAéreasToolStripMenuItem
            this.terminalesAéreasToolStripMenuItem.Name = "terminalesAéreasToolStripMenuItem";
            this.terminalesAéreasToolStripMenuItem.Size = new System.Drawing.Size(113, 20);
            this.terminalesAéreasToolStripMenuItem.Text = "Terminales Aéreas";
            this.terminalesAéreasToolStripMenuItem.Click += new System.EventHandler(this.btnTerminales_Click);

            // configuraciónToolStripMenuItem
            this.configuraciónToolStripMenuItem.Name = "configuraciónToolStripMenuItem";
            this.configuraciónToolStripMenuItem.Size = new System.Drawing.Size(95, 20);
            this.configuraciónToolStripMenuItem.Text = "Configuración";
            this.configuraciónToolStripMenuItem.Click += new System.EventHandler(this.btnConfig_Click);

            // FormVinculaciones
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(anchuraForm, 600);
            this.Controls.Add(this.dgvVinculaciones);
            this.Controls.Add(this.pnlFilters);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormVinculaciones";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Listado de Vinculaciones Académicas";
            this.pnlFilters.ResumeLayout(false);
            this.pnlFilters.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVinculaciones)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel pnlFilters;
        private System.Windows.Forms.Button btnFiltrar;
        private System.Windows.Forms.CheckBox chkSoloViaticas;
        private System.Windows.Forms.ComboBox cmbTercero;
        private System.Windows.Forms.Label lblTercero;
        private System.Windows.Forms.ComboBox cmbSede;
        private System.Windows.Forms.Label lblSede;
        private System.Windows.Forms.DataGridView dgvVinculaciones;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.ToolStripStatusLabel lblUsuarioActivo;
        private System.Windows.Forms.Button btnLimpiar;
        private System.Windows.Forms.Button btnRevisar;
        private System.Windows.Forms.Button btnAdicionar;
        private System.Windows.Forms.Button btnVerVinculaciones;

        private System.Windows.Forms.DateTimePicker dtpFechaInicio;
        private System.Windows.Forms.DateTimePicker dtpFechaFinal;
        private System.Windows.Forms.Label lblFechaInicio;
        private System.Windows.Forms.Label lblFechaFinal;
        private System.Windows.Forms.ComboBox cmbSemestres;
        private System.Windows.Forms.Label lblSemestres;
        private System.Windows.Forms.ComboBox cmbProgramas;
        private System.Windows.Forms.Label lblProgramas;

        // Retenidos ocultos
        private System.Windows.Forms.ComboBox cmbTerritoriales;
        private System.Windows.Forms.ComboBox cmbPeriodo;

        // Selector de corridas (MODO AG · fila 3)
        private System.Windows.Forms.Button btnAjustarFiltrosCorrida;
        private System.Windows.Forms.Button btnEliminarCorrida;
        private System.Windows.Forms.Label lblCorrida;
        private System.Windows.Forms.Label lblModo;

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem sedesAcadémicasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tercerosDocentesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem díasFestivosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem gastosTerrestresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dependenciasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem terminalesAéreasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem configuraciónToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem períodosToolStripMenuItem;
    }
}
