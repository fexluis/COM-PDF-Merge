namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormTerceros
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
            this.tabControlTerceros = new System.Windows.Forms.TabControl();
            this.tabDatos = new System.Windows.Forms.TabPage();
            this.tabAsignaturas = new System.Windows.Forms.TabPage();
            this.pnlAsignaturasToolbar = new System.Windows.Forms.Panel();
            this.lblDocenteAsignaturasTitulo = new System.Windows.Forms.Label();
            this.btnAgregarAsignatura = new System.Windows.Forms.Button();
            this.btnEditarAsignatura = new System.Windows.Forms.Button();
            this.btnEliminarAsignatura = new System.Windows.Forms.Button();
            this.btnImportarHistorico = new System.Windows.Forms.Button();
            this.btnRefrescarAsignaturas = new System.Windows.Forms.Button();
            this.dgvAsignaturasHabilitadas = new System.Windows.Forms.DataGridView();
            this.dgvTerceros = new System.Windows.Forms.DataGridView();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.btnExportarBase = new System.Windows.Forms.Button();
            this.btnExportarTodos = new System.Windows.Forms.Button();
            this.lblTerritorial = new System.Windows.Forms.Label();
            this.cmbTerritorial = new System.Windows.Forms.ComboBox();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.lblNombre = new System.Windows.Forms.Label();
            this.txtCedula = new System.Windows.Forms.TextBox();
            this.btnEditCedula = new System.Windows.Forms.Button();
            this.lblCedula = new System.Windows.Forms.Label();
            this.chkListDepartamentos = new System.Windows.Forms.CheckedListBox();
            this.lblDepartamento = new System.Windows.Forms.Label();
            this.cmbTipo = new System.Windows.Forms.ComboBox();
            this.lblTipo = new System.Windows.Forms.Label();
            this.numSalario = new System.Windows.Forms.NumericUpDown();
            this.numPuntaje = new System.Windows.Forms.NumericUpDown();
            this.lblPuntaje = new System.Windows.Forms.Label();
            this.lblSalario = new System.Windows.Forms.Label();
            this.cmbCategoria = new System.Windows.Forms.ComboBox();
            this.lblCategoria = new System.Windows.Forms.Label();
            this.txtEntidad = new System.Windows.Forms.TextBox();
            this.lblEntidad = new System.Windows.Forms.Label();
            this.cmbCuenta = new System.Windows.Forms.ComboBox();
            this.lblCuenta = new System.Windows.Forms.Label();
            this.txtTelefono = new System.Windows.Forms.TextBox();
            this.lblTelefono = new System.Windows.Forms.Label();
            this.txtCorreo = new System.Windows.Forms.TextBox();
            this.lblCorreo = new System.Windows.Forms.Label();
            this.dtpFechaNacimiento = new System.Windows.Forms.DateTimePicker();
            this.lblFechaNacimiento = new System.Windows.Forms.Label();
            this.txtContrato = new System.Windows.Forms.TextBox();
            this.lblContrato = new System.Windows.Forms.Label();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.cmbTipoCuenta = new System.Windows.Forms.ComboBox();
            this.txtCorreoPersonal = new System.Windows.Forms.TextBox();
            this.lblCorreoPersonal = new System.Windows.Forms.Label();
            this.txtResidencia = new System.Windows.Forms.TextBox();
            this.btnSelectCiudad = new System.Windows.Forms.Button();
            this.lblResidencia = new System.Windows.Forms.Label();
            this.chkActivo = new System.Windows.Forms.CheckBox();
            this.chkPensionado = new System.Windows.Forms.CheckBox();
            this.chkMostrarInactivos = new System.Windows.Forms.CheckBox();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblBuscar = new System.Windows.Forms.Label();
            this.txtBuscar = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerceros)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAsignaturasHabilitadas)).BeginInit();
            this.pnlDetails.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numSalario)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPuntaje)).BeginInit();
            this.SuspendLayout();
            // 
            // 
            // lblTerritorial
            // 
            this.lblTerritorial.AutoSize = true;
            this.lblTerritorial.Location = new System.Drawing.Point(190, 14);
            this.lblTerritorial.Name = "lblTerritorial";
            this.lblTerritorial.Size = new System.Drawing.Size(56, 13);
            this.lblTerritorial.TabIndex = 39;
            this.lblTerritorial.Text = "Territorial:";
            // 
            // cmbTerritorial
            // 
            this.cmbTerritorial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTerritorial.FormattingEnabled = true;
            this.cmbTerritorial.Location = new System.Drawing.Point(250, 11);
            this.cmbTerritorial.Name = "cmbTerritorial";
            this.cmbTerritorial.Size = new System.Drawing.Size(120, 21);
            this.cmbTerritorial.TabIndex = 40;
            this.cmbTerritorial.SelectedIndexChanged += new System.EventHandler(this.cmbTerritorial_SelectedIndexChanged);
            // 
            // dgvTerceros
            // 
            this.dgvTerceros.AllowUserToAddRows = false;
            this.dgvTerceros.AllowUserToDeleteRows = false;
            this.dgvTerceros.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTerceros.Dock = System.Windows.Forms.DockStyle.Left;
            this.dgvTerceros.Location = new System.Drawing.Point(0, 0);
            this.dgvTerceros.Name = "dgvTerceros";
            this.dgvTerceros.ReadOnly = true;
            this.dgvTerceros.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvTerceros.Size = new System.Drawing.Size(500, 580);
            this.dgvTerceros.TabIndex = 0;
            // 
            // chkMostrarInactivos
            // 
            this.chkMostrarInactivos.AutoSize = true;
            this.chkMostrarInactivos.Location = new System.Drawing.Point(380, 14);
            this.chkMostrarInactivos.Name = "chkMostrarInactivos";
            this.chkMostrarInactivos.Size = new System.Drawing.Size(107, 17);
            this.chkMostrarInactivos.TabIndex = 28;
            this.chkMostrarInactivos.Text = "Mostrar inactivos";
            this.chkMostrarInactivos.UseVisualStyleBackColor = true;
            this.chkMostrarInactivos.CheckedChanged += new System.EventHandler(this.chkMostrarInactivos_CheckedChanged);
            // 
            // 
            // btnExportarBase
            // 
            this.btnExportarBase.Location = new System.Drawing.Point(510, 8);
            this.btnExportarBase.Name = "btnExportarBase";
            this.btnExportarBase.Size = new System.Drawing.Size(150, 25);
            this.btnExportarBase.TabIndex = 37;
            this.btnExportarBase.Text = "Exportar base docente";
            this.btnExportarBase.UseVisualStyleBackColor = true;
            this.btnExportarBase.Click += new System.EventHandler(this.btnExportarBase_Click);
            // 
            // btnExportarTodos
            // 
            this.btnExportarTodos.Location = new System.Drawing.Point(670, 8);
            this.btnExportarTodos.Name = "btnExportarTodos";
            this.btnExportarTodos.Size = new System.Drawing.Size(150, 25);
            this.btnExportarTodos.TabIndex = 38;
            this.btnExportarTodos.Text = "Exportar todos";
            this.btnExportarTodos.UseVisualStyleBackColor = true;
            this.btnExportarTodos.Click += new System.EventHandler(this.btnExportarTodos_Click);
            // 
            // pnlTop
            // 
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Height = 40;
            this.pnlTop.Controls.Add(this.lblBuscar);
            this.pnlTop.Controls.Add(this.txtBuscar);
            this.pnlTop.Controls.Add(this.btnExportarBase);
            this.pnlTop.Controls.Add(this.btnExportarTodos);
            this.pnlTop.Controls.Add(this.chkMostrarInactivos);
            this.pnlTop.Controls.Add(this.lblTerritorial);
            this.pnlTop.Controls.Add(this.cmbTerritorial);
            //
            // lblBuscar
            //
            this.lblBuscar.AutoSize = true;
            this.lblBuscar.Location = new System.Drawing.Point(12, 14);
            this.lblBuscar.Text = "Buscar:";
            //
            // txtBuscar
            //
            this.txtBuscar.Location = new System.Drawing.Point(60, 11);
            this.txtBuscar.Width = 120;
            this.txtBuscar.Name = "txtBuscar";
            // 
            // pnlDetails
            // 

            this.pnlDetails.Controls.Add(this.btnCerrar);
            this.pnlDetails.Controls.Add(this.btnEliminar);
            this.pnlDetails.Controls.Add(this.btnGuardar);
            this.pnlDetails.Controls.Add(this.btnNuevo);
            this.pnlDetails.Controls.Add(this.txtContrato);
            this.pnlDetails.Controls.Add(this.lblContrato);
            this.pnlDetails.Controls.Add(this.dtpFechaNacimiento);
            this.pnlDetails.Controls.Add(this.lblFechaNacimiento);
            this.pnlDetails.Controls.Add(this.txtCorreo);
            this.pnlDetails.Controls.Add(this.lblCorreo);
            this.pnlDetails.Controls.Add(this.txtTelefono);
            this.pnlDetails.Controls.Add(this.lblTelefono);
            this.pnlDetails.Controls.Add(this.cmbCuenta);
            this.pnlDetails.Controls.Add(this.lblCuenta);
            this.pnlDetails.Controls.Add(this.txtEntidad);
            this.pnlDetails.Controls.Add(this.lblEntidad);
            this.pnlDetails.Controls.Add(this.cmbCategoria);
            this.pnlDetails.Controls.Add(this.lblCategoria);
            this.pnlDetails.Controls.Add(this.numSalario);
            this.pnlDetails.Controls.Add(this.numPuntaje);
            this.pnlDetails.Controls.Add(this.lblPuntaje);
            this.pnlDetails.Controls.Add(this.lblSalario);
            this.pnlDetails.Controls.Add(this.cmbTipo);
            this.pnlDetails.Controls.Add(this.lblTipo);
            this.pnlDetails.Controls.Add(this.chkListDepartamentos);
            this.pnlDetails.Controls.Add(this.lblDepartamento);
            this.pnlDetails.Controls.Add(this.btnEditCedula);
            this.pnlDetails.Controls.Add(this.txtCedula);
            this.pnlDetails.Controls.Add(this.lblCedula);
            this.pnlDetails.Controls.Add(this.txtNombre);
            this.pnlDetails.Controls.Add(this.lblNombre);
            this.pnlDetails.Controls.Add(this.cmbTipoCuenta);
            this.pnlDetails.Controls.Add(this.txtCorreoPersonal);
            this.pnlDetails.Controls.Add(this.lblCorreoPersonal);
            this.pnlDetails.Controls.Add(this.txtResidencia);
            this.pnlDetails.Controls.Add(this.btnSelectCiudad);
            this.pnlDetails.Controls.Add(this.lblResidencia);
            this.pnlDetails.Controls.Add(this.chkPensionado);
            this.pnlDetails.Controls.Add(this.chkActivo);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDetails.Location = new System.Drawing.Point(500, 0);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(400, 700);
            this.pnlDetails.TabIndex = 1;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(294, 620);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 27;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Enabled = false;
            this.btnEliminar.ForeColor = System.Drawing.Color.Red;
            this.btnEliminar.Location = new System.Drawing.Point(213, 620);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(75, 23);
            this.btnEliminar.TabIndex = 26;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.Location = new System.Drawing.Point(132, 620);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 25;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnNuevo
            // 
            this.btnNuevo.Location = new System.Drawing.Point(51, 620);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.Size = new System.Drawing.Size(75, 23);
            this.btnNuevo.TabIndex = 24;
            this.btnNuevo.Text = "Nuevo";
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            // 
            // cmbTipoCuenta
            // 
            this.cmbTipoCuenta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoCuenta.FormattingEnabled = true;
            this.cmbTipoCuenta.Items.AddRange(new object[] {
            "Ahorros",
            "Corriente"});
            this.cmbTipoCuenta.Location = new System.Drawing.Point(280, 382);
            this.cmbTipoCuenta.Name = "cmbTipoCuenta";
            this.cmbTipoCuenta.Size = new System.Drawing.Size(90, 21);
            this.cmbTipoCuenta.TabIndex = 29;
            // 
            // txtCorreoPersonal
            // 
            this.txtCorreoPersonal.Location = new System.Drawing.Point(120, 472);
            this.txtCorreoPersonal.Name = "txtCorreoPersonal";
            this.txtCorreoPersonal.Size = new System.Drawing.Size(250, 20);
            this.txtCorreoPersonal.TabIndex = 31;
            // 
            // lblCorreoPersonal
            // 
            this.lblCorreoPersonal.Location = new System.Drawing.Point(20, 475);
            this.lblCorreoPersonal.Name = "lblCorreoPersonal";
            this.lblCorreoPersonal.Size = new System.Drawing.Size(100, 13);
            this.lblCorreoPersonal.TabIndex = 32;
            this.lblCorreoPersonal.Text = "Correo Pers.:";
            // 
            // txtResidencia
            // 
            this.txtResidencia.Location = new System.Drawing.Point(120, 561);
            this.txtResidencia.Name = "txtResidencia";
            this.txtResidencia.ReadOnly = true;
            this.txtResidencia.Size = new System.Drawing.Size(160, 20);
            this.txtResidencia.TabIndex = 33;
            // 
            // btnSelectCiudad
            // 
            this.btnSelectCiudad.Location = new System.Drawing.Point(285, 560);
            this.btnSelectCiudad.Name = "btnSelectCiudad";
            this.btnSelectCiudad.Size = new System.Drawing.Size(85, 23);
            this.btnSelectCiudad.TabIndex = 35;
            this.btnSelectCiudad.Text = "Ciudad";
            this.btnSelectCiudad.UseVisualStyleBackColor = true;
            this.btnSelectCiudad.Click += new System.EventHandler(this.btnSelectCiudad_Click);
            // 
            // lblResidencia
            // 
            this.lblResidencia.Location = new System.Drawing.Point(20, 565);
            this.lblResidencia.Name = "lblResidencia";
            this.lblResidencia.Size = new System.Drawing.Size(100, 13);
            this.lblResidencia.TabIndex = 34;
            this.lblResidencia.Text = "Residencia:";
            // 
            // chkActivo
            // 
            this.chkActivo.AutoSize = true;
            this.chkActivo.Location = new System.Drawing.Point(120, 592);
            this.chkActivo.Name = "chkActivo";
            this.chkActivo.Size = new System.Drawing.Size(56, 17);
            this.chkActivo.TabIndex = 35;
            this.chkActivo.Text = "Activo";
            this.chkActivo.UseVisualStyleBackColor = true;
            // 
            // chkPensionado
            // 
            this.chkPensionado.AutoSize = true;
            this.chkPensionado.Location = new System.Drawing.Point(200, 592);
            this.chkPensionado.Name = "chkPensionado";
            this.chkPensionado.Size = new System.Drawing.Size(82, 17);
            this.chkPensionado.TabIndex = 36;
            this.chkPensionado.Text = "Pensionado";
            this.chkPensionado.UseVisualStyleBackColor = true;
            // 
            // txtContrato
            // 
            this.txtContrato.Location = new System.Drawing.Point(120, 532);
            this.txtContrato.Name = "txtContrato";
            this.txtContrato.Size = new System.Drawing.Size(250, 20);
            this.txtContrato.TabIndex = 23;
            // 
            // lblContrato
            // 
            this.lblContrato.Location = new System.Drawing.Point(20, 535);
            this.lblContrato.Name = "lblContrato";
            this.lblContrato.Size = new System.Drawing.Size(100, 13);
            this.lblContrato.TabIndex = 22;
            this.lblContrato.Text = "Contrato:";
            // 
            // 
            // lblFechaNacimiento
            // 
            this.lblFechaNacimiento.Location = new System.Drawing.Point(20, 505);
            this.lblFechaNacimiento.Name = "lblFechaNacimiento";
            this.lblFechaNacimiento.Size = new System.Drawing.Size(100, 13);
            this.lblFechaNacimiento.TabIndex = 20;
            this.lblFechaNacimiento.Text = "Nacimiento:";
            // 
            // dtpFechaNacimiento
            // 
            this.dtpFechaNacimiento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFechaNacimiento.Location = new System.Drawing.Point(120, 502);
            this.dtpFechaNacimiento.Name = "dtpFechaNacimiento";
            this.dtpFechaNacimiento.Size = new System.Drawing.Size(120, 20);
            this.dtpFechaNacimiento.TabIndex = 21;
            // 
            // txtCorreo
            // 
            this.txtCorreo.Location = new System.Drawing.Point(120, 442);
            this.txtCorreo.Name = "txtCorreo";
            this.txtCorreo.Size = new System.Drawing.Size(250, 20);
            this.txtCorreo.TabIndex = 19;
            // 
            // lblCorreo
            // 
            this.lblCorreo.Location = new System.Drawing.Point(20, 445);
            this.lblCorreo.Name = "lblCorreo";
            this.lblCorreo.Size = new System.Drawing.Size(100, 13);
            this.lblCorreo.TabIndex = 18;
            this.lblCorreo.Text = "Correo:";
            // 
            // txtTelefono
            // 
            this.txtTelefono.Location = new System.Drawing.Point(120, 412);
            this.txtTelefono.Name = "txtTelefono";
            this.txtTelefono.Size = new System.Drawing.Size(150, 20);
            this.txtTelefono.TabIndex = 17;
            // 
            // lblTelefono
            // 
            this.lblTelefono.Location = new System.Drawing.Point(20, 415);
            this.lblTelefono.Name = "lblTelefono";
            this.lblTelefono.Size = new System.Drawing.Size(100, 13);
            this.lblTelefono.TabIndex = 16;
            this.lblTelefono.Text = "Teléfono:";
            // 
            // cmbCuenta
            // 
            this.cmbCuenta.FormattingEnabled = true;
            this.cmbCuenta.Items.AddRange(new object[] {
            "Ahorro",
            "Corriente"});
            this.cmbCuenta.Location = new System.Drawing.Point(120, 382);
            this.cmbCuenta.Name = "cmbCuenta";
            this.cmbCuenta.Size = new System.Drawing.Size(150, 21);
            this.cmbCuenta.TabIndex = 15;
            // 
            // lblCuenta
            // 
            this.lblCuenta.Location = new System.Drawing.Point(20, 385);
            this.lblCuenta.Name = "lblCuenta";
            this.lblCuenta.Size = new System.Drawing.Size(100, 13);
            this.lblCuenta.TabIndex = 14;
            this.lblCuenta.Text = "Cuenta:";
            // 
            // txtEntidad
            // 
            this.txtEntidad.Location = new System.Drawing.Point(120, 352);
            this.txtEntidad.Name = "txtEntidad";
            this.txtEntidad.Size = new System.Drawing.Size(250, 20);
            this.txtEntidad.TabIndex = 13;
            // 
            // lblEntidad
            // 
            this.lblEntidad.Location = new System.Drawing.Point(20, 355);
            this.lblEntidad.Name = "lblEntidad";
            this.lblEntidad.Size = new System.Drawing.Size(100, 13);
            this.lblEntidad.TabIndex = 12;
            this.lblEntidad.Text = "Entidad Banc.:";
            // 
            // cmbCategoria
            // 
            this.cmbCategoria.FormattingEnabled = true;
            this.cmbCategoria.Items.AddRange(new object[] {
            "N/A",
            "TITULAR",
            "ASOCIADO",
            "ASISTENTE",
            "AUXILIAR"});
            this.cmbCategoria.Location = new System.Drawing.Point(120, 322);
            this.cmbCategoria.Name = "cmbCategoria";
            this.cmbCategoria.Size = new System.Drawing.Size(200, 21);
            this.cmbCategoria.TabIndex = 11;
            // 
            // lblCategoria
            // 
            this.lblCategoria.Location = new System.Drawing.Point(20, 325);
            this.lblCategoria.Name = "lblCategoria";
            this.lblCategoria.Size = new System.Drawing.Size(100, 13);
            this.lblCategoria.TabIndex = 10;
            this.lblCategoria.Text = "Categoría:";
            // 
            // numSalario
            // 
            this.numSalario.DecimalPlaces = 2;
            this.numSalario.Location = new System.Drawing.Point(120, 292);
            this.numSalario.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.numSalario.Name = "numSalario";
            this.numSalario.Size = new System.Drawing.Size(100, 20);
            this.numSalario.TabIndex = 9;
            // 
            // lblSalario
            // 
            this.lblSalario.Location = new System.Drawing.Point(20, 295);
            this.lblSalario.Name = "lblSalario";
            this.lblSalario.Size = new System.Drawing.Size(100, 13);
            this.lblSalario.TabIndex = 8;
            this.lblSalario.Text = "Salario:";
            // 
            // numPuntaje
            // 
            this.numPuntaje.DecimalPlaces = 2;
            this.numPuntaje.Location = new System.Drawing.Point(280, 292);
            this.numPuntaje.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.numPuntaje.Name = "numPuntaje";
            this.numPuntaje.Size = new System.Drawing.Size(90, 20);
            this.numPuntaje.TabIndex = 22;
            // 
            // lblPuntaje
            // 
            this.lblPuntaje.Location = new System.Drawing.Point(230, 295);
            this.lblPuntaje.Name = "lblPuntaje";
            this.lblPuntaje.Size = new System.Drawing.Size(50, 13);
            this.lblPuntaje.TabIndex = 23;
            this.lblPuntaje.Text = "Puntaje:";
            // 
            // cmbTipo
            // 
            this.cmbTipo.FormattingEnabled = true;
            this.cmbTipo.Items.AddRange(new object[] {
            "CARRERA",
            "OCASIONAL",
            "CATEDRA",
            "ESPECIAL/VISITANTE",
            "CONTRATISTA"});
            this.cmbTipo.Location = new System.Drawing.Point(120, 262);
            this.cmbTipo.Name = "cmbTipo";
            this.cmbTipo.Size = new System.Drawing.Size(200, 21);
            this.cmbTipo.TabIndex = 7;
            // 
            // lblTipo
            // 
            this.lblTipo.Location = new System.Drawing.Point(20, 265);
            this.lblTipo.Name = "lblTipo";
            this.lblTipo.Size = new System.Drawing.Size(100, 13);
            this.lblTipo.TabIndex = 6;
            this.lblTipo.Text = "Tipo:";
            // 
            // chkListDepartamentos
            // 
            this.chkListDepartamentos.CheckOnClick = true;
            this.chkListDepartamentos.FormattingEnabled = true;
            this.chkListDepartamentos.Location = new System.Drawing.Point(120, 70);
            this.chkListDepartamentos.Name = "chkListDepartamentos";
            this.chkListDepartamentos.Size = new System.Drawing.Size(200, 184);
            this.chkListDepartamentos.TabIndex = 5;
            this.chkListDepartamentos.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkListDepartamentos_ItemCheck);
            // 
            // lblDepartamento
            // 
            this.lblDepartamento.AutoSize = true;
            this.lblDepartamento.Location = new System.Drawing.Point(20, 73);
            this.lblDepartamento.Name = "lblDepartamento";
            this.lblDepartamento.Size = new System.Drawing.Size(100, 26);
            this.lblDepartamento.TabIndex = 4;
            this.lblDepartamento.Text = "Territorial:\n(Para dictar clases)";
            // 
            // txtCedula
            // 
            this.txtCedula.Location = new System.Drawing.Point(120, 40);
            this.txtCedula.Name = "txtCedula";
            this.txtCedula.Size = new System.Drawing.Size(120, 20);
            this.txtCedula.TabIndex = 3;
            this.txtCedula.ReadOnly = true;
            // 
            // btnEditCedula
            // 
            this.btnEditCedula.Location = new System.Drawing.Point(245, 39);
            this.btnEditCedula.Name = "btnEditCedula";
            this.btnEditCedula.Size = new System.Drawing.Size(25, 22);
            this.btnEditCedula.TabIndex = 41;
            this.btnEditCedula.Text = "E";
            this.btnEditCedula.UseVisualStyleBackColor = true;
            this.btnEditCedula.Click += new System.EventHandler(this.btnEditCedula_Click);
            // 
            // lblCedula
            // 
            this.lblCedula.Location = new System.Drawing.Point(20, 43);
            this.lblCedula.Name = "lblCedula";
            this.lblCedula.Size = new System.Drawing.Size(100, 13);
            this.lblCedula.TabIndex = 2;
            this.lblCedula.Text = "Cédula:";
            // 
            // txtNombre
            // 
            this.txtNombre.Location = new System.Drawing.Point(120, 10);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(250, 20);
            this.txtNombre.TabIndex = 1;
            this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            // 
            // lblNombre
            // 
            this.lblNombre.Location = new System.Drawing.Point(20, 13);
            this.lblNombre.Name = "lblNombre";
            this.lblNombre.Size = new System.Drawing.Size(100, 13);
            this.lblNombre.TabIndex = 0;
            this.lblNombre.Text = "Nombre:";
            // 
            // tabControlTerceros
            // 
            this.tabControlTerceros.Controls.Add(this.tabDatos);
            this.tabControlTerceros.Controls.Add(this.tabAsignaturas);
            this.tabControlTerceros.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlTerceros.Location = new System.Drawing.Point(0, 0);
            this.tabControlTerceros.Name = "tabControlTerceros";
            this.tabControlTerceros.SelectedIndex = 0;
            this.tabControlTerceros.Size = new System.Drawing.Size(900, 700);
            this.tabControlTerceros.TabIndex = 0;
            // 
            // tabDatos
            // 
            this.tabDatos.Controls.Add(this.pnlDetails);
            this.tabDatos.Controls.Add(this.dgvTerceros);
            this.tabDatos.Controls.Add(this.pnlTop);
            this.tabDatos.Location = new System.Drawing.Point(4, 22);
            this.tabDatos.Name = "tabDatos";
            this.tabDatos.Padding = new System.Windows.Forms.Padding(3);
            this.tabDatos.Size = new System.Drawing.Size(892, 674);
            this.tabDatos.TabIndex = 0;
            this.tabDatos.Text = "Datos del Tercero";
            this.tabDatos.UseVisualStyleBackColor = true;
            // 
            // tabAsignaturas
            // 
            this.tabAsignaturas.Controls.Add(this.dgvAsignaturasHabilitadas);
            this.tabAsignaturas.Controls.Add(this.pnlAsignaturasToolbar);
            this.tabAsignaturas.Location = new System.Drawing.Point(4, 22);
            this.tabAsignaturas.Name = "tabAsignaturas";
            this.tabAsignaturas.Padding = new System.Windows.Forms.Padding(3);
            this.tabAsignaturas.Size = new System.Drawing.Size(892, 674);
            this.tabAsignaturas.TabIndex = 1;
            this.tabAsignaturas.Text = "Asignaturas Habilitadas";
            this.tabAsignaturas.UseVisualStyleBackColor = true;
            // 
            // pnlAsignaturasToolbar
            // 
            this.pnlAsignaturasToolbar.Controls.Add(this.lblDocenteAsignaturasTitulo);
            this.pnlAsignaturasToolbar.Controls.Add(this.btnAgregarAsignatura);
            this.pnlAsignaturasToolbar.Controls.Add(this.btnEditarAsignatura);
            this.pnlAsignaturasToolbar.Controls.Add(this.btnEliminarAsignatura);
            this.pnlAsignaturasToolbar.Controls.Add(this.btnImportarHistorico);
            this.pnlAsignaturasToolbar.Controls.Add(this.btnRefrescarAsignaturas);
            this.pnlAsignaturasToolbar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlAsignaturasToolbar.Location = new System.Drawing.Point(3, 3);
            this.pnlAsignaturasToolbar.Name = "pnlAsignaturasToolbar";
            this.pnlAsignaturasToolbar.Size = new System.Drawing.Size(886, 45);
            this.pnlAsignaturasToolbar.TabIndex = 0;
            // 
            // lblDocenteAsignaturasTitulo
            // 
            this.lblDocenteAsignaturasTitulo.AutoSize = true;
            this.lblDocenteAsignaturasTitulo.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold);
            this.lblDocenteAsignaturasTitulo.Location = new System.Drawing.Point(10, 14);
            this.lblDocenteAsignaturasTitulo.Name = "lblDocenteAsignaturasTitulo";
            this.lblDocenteAsignaturasTitulo.Size = new System.Drawing.Size(185, 17);
            this.lblDocenteAsignaturasTitulo.TabIndex = 0;
            this.lblDocenteAsignaturasTitulo.Text = "Docente: (ninguno seleccionado)";
            // 
            // btnAgregarAsignatura
            // 
            this.btnAgregarAsignatura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAgregarAsignatura.Location = new System.Drawing.Point(400, 8);
            this.btnAgregarAsignatura.Name = "btnAgregarAsignatura";
            this.btnAgregarAsignatura.Size = new System.Drawing.Size(85, 28);
            this.btnAgregarAsignatura.TabIndex = 1;
            this.btnAgregarAsignatura.Text = "Agregar...";
            this.btnAgregarAsignatura.UseVisualStyleBackColor = true;
            this.btnAgregarAsignatura.Click += new System.EventHandler(this.BtnAgregarAsignatura_Click);
            // 
            // btnEditarAsignatura
            // 
            this.btnEditarAsignatura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEditarAsignatura.Location = new System.Drawing.Point(490, 8);
            this.btnEditarAsignatura.Name = "btnEditarAsignatura";
            this.btnEditarAsignatura.Size = new System.Drawing.Size(85, 28);
            this.btnEditarAsignatura.TabIndex = 2;
            this.btnEditarAsignatura.Text = "Editar...";
            this.btnEditarAsignatura.UseVisualStyleBackColor = true;
            this.btnEditarAsignatura.Click += new System.EventHandler(this.BtnEditarAsignatura_Click);
            // 
            // btnEliminarAsignatura
            // 
            this.btnEliminarAsignatura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEliminarAsignatura.Location = new System.Drawing.Point(580, 8);
            this.btnEliminarAsignatura.Name = "btnEliminarAsignatura";
            this.btnEliminarAsignatura.Size = new System.Drawing.Size(85, 28);
            this.btnEliminarAsignatura.TabIndex = 3;
            this.btnEliminarAsignatura.Text = "Eliminar";
            this.btnEliminarAsignatura.UseVisualStyleBackColor = true;
            this.btnEliminarAsignatura.Click += new System.EventHandler(this.BtnEliminarAsignatura_Click);
            // 
            // btnImportarHistorico
            // 
            this.btnImportarHistorico.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImportarHistorico.Location = new System.Drawing.Point(670, 8);
            this.btnImportarHistorico.Name = "btnImportarHistorico";
            this.btnImportarHistorico.Size = new System.Drawing.Size(125, 28);
            this.btnImportarHistorico.TabIndex = 4;
            this.btnImportarHistorico.Text = "Importar Histórico";
            this.btnImportarHistorico.UseVisualStyleBackColor = true;
            this.btnImportarHistorico.Click += new System.EventHandler(this.BtnImportarHistorico_Click);
            // 
            // btnRefrescarAsignaturas
            // 
            this.btnRefrescarAsignaturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefrescarAsignaturas.Location = new System.Drawing.Point(800, 8);
            this.btnRefrescarAsignaturas.Name = "btnRefrescarAsignaturas";
            this.btnRefrescarAsignaturas.Size = new System.Drawing.Size(80, 28);
            this.btnRefrescarAsignaturas.TabIndex = 5;
            this.btnRefrescarAsignaturas.Text = "Refrescar";
            this.btnRefrescarAsignaturas.UseVisualStyleBackColor = true;
            this.btnRefrescarAsignaturas.Click += new System.EventHandler(this.BtnRefrescarAsignaturas_Click);
            // 
            // dgvAsignaturasHabilitadas
            // 
            this.dgvAsignaturasHabilitadas.AllowUserToAddRows = false;
            this.dgvAsignaturasHabilitadas.AllowUserToDeleteRows = false;
            this.dgvAsignaturasHabilitadas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAsignaturasHabilitadas.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvAsignaturasHabilitadas.Location = new System.Drawing.Point(3, 48);
            this.dgvAsignaturasHabilitadas.Name = "dgvAsignaturasHabilitadas";
            this.dgvAsignaturasHabilitadas.ReadOnly = true;
            this.dgvAsignaturasHabilitadas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAsignaturasHabilitadas.Size = new System.Drawing.Size(886, 623);
            this.dgvAsignaturasHabilitadas.TabIndex = 1;
            // 
            // FormTerceros
            // 
            this.ClientSize = new System.Drawing.Size(900, 700);
            this.Controls.Add(this.tabControlTerceros);
            this.Name = "FormTerceros";
            this.Text = "Gestión de Terceros y Habilitaciones";
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerceros)).EndInit();
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numSalario)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPuntaje)).EndInit();
            this.tabControlTerceros.ResumeLayout(false);
            this.tabDatos.ResumeLayout(false);
            this.tabAsignaturas.ResumeLayout(false);
            this.pnlAsignaturasToolbar.ResumeLayout(false);
            this.pnlAsignaturasToolbar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAsignaturasHabilitadas)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.TabControl tabControlTerceros;
        private System.Windows.Forms.TabPage tabDatos;
        private System.Windows.Forms.TabPage tabAsignaturas;
        private System.Windows.Forms.Panel pnlAsignaturasToolbar;
        private System.Windows.Forms.Label lblDocenteAsignaturasTitulo;
        private System.Windows.Forms.Button btnAgregarAsignatura;
        private System.Windows.Forms.Button btnEditarAsignatura;
        private System.Windows.Forms.Button btnEliminarAsignatura;
        private System.Windows.Forms.Button btnImportarHistorico;
        private System.Windows.Forms.Button btnRefrescarAsignaturas;
        private System.Windows.Forms.DataGridView dgvAsignaturasHabilitadas;
        private System.Windows.Forms.DataGridView dgvTerceros;
        private System.Windows.Forms.Panel pnlDetails;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label lblNombre;
        private System.Windows.Forms.TextBox txtCedula;
        private System.Windows.Forms.Button btnEditCedula;
        private System.Windows.Forms.Label lblCedula;
        private System.Windows.Forms.CheckedListBox chkListDepartamentos;
        private System.Windows.Forms.Label lblDepartamento;
        private System.Windows.Forms.ComboBox cmbTipo;
        private System.Windows.Forms.Label lblTipo;
        private System.Windows.Forms.NumericUpDown numSalario;
        private System.Windows.Forms.NumericUpDown numPuntaje;
        private System.Windows.Forms.Label lblPuntaje;
        private System.Windows.Forms.Label lblSalario;
        private System.Windows.Forms.ComboBox cmbCategoria;
        private System.Windows.Forms.Label lblCategoria;
        private System.Windows.Forms.TextBox txtEntidad;
        private System.Windows.Forms.Label lblEntidad;
        private System.Windows.Forms.ComboBox cmbCuenta;
        private System.Windows.Forms.Label lblCuenta;
        private System.Windows.Forms.TextBox txtTelefono;
        private System.Windows.Forms.Label lblTelefono;
        private System.Windows.Forms.TextBox txtCorreo;
        private System.Windows.Forms.Label lblCorreo;
        private System.Windows.Forms.DateTimePicker dtpFechaNacimiento;
        private System.Windows.Forms.Label lblFechaNacimiento;
        private System.Windows.Forms.TextBox txtContrato;
        private System.Windows.Forms.Label lblContrato;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.ComboBox cmbTipoCuenta;
        private System.Windows.Forms.TextBox txtCorreoPersonal;
        private System.Windows.Forms.Label lblCorreoPersonal;
        private System.Windows.Forms.TextBox txtResidencia;
        private System.Windows.Forms.Button btnSelectCiudad;
        private System.Windows.Forms.Label lblResidencia;
        private System.Windows.Forms.CheckBox chkActivo;
        private System.Windows.Forms.CheckBox chkPensionado;
        private System.Windows.Forms.CheckBox chkMostrarInactivos;
        private System.Windows.Forms.Label lblTerritorial;
        private System.Windows.Forms.ComboBox cmbTerritorial;
        private System.Windows.Forms.Button btnExportarBase;
        private System.Windows.Forms.Button btnExportarTodos;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblBuscar;
        private System.Windows.Forms.TextBox txtBuscar;
    }
}

