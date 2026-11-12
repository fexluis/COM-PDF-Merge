namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class AgFormGeneradorHorarios
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
            this.components = new System.ComponentModel.Container();
            this.tabControlPrincipal = new System.Windows.Forms.TabControl();
            this.tabSimulacion = new System.Windows.Forms.TabPage();
            this.grpParametrosAG = new System.Windows.Forms.GroupBox();
            this.numElitismo = new System.Windows.Forms.NumericUpDown();
            this.lblElitismo = new System.Windows.Forms.Label();
            this.numTasaMutacion = new System.Windows.Forms.NumericUpDown();
            this.lblMutacion = new System.Windows.Forms.Label();
            this.numTasaCruce = new System.Windows.Forms.NumericUpDown();
            this.lblCruce = new System.Windows.Forms.Label();
            this.numGeneraciones = new System.Windows.Forms.NumericUpDown();
            this.lblGeneraciones = new System.Windows.Forms.Label();
            this.numPoblacion = new System.Windows.Forms.NumericUpDown();
            this.lblPoblacion = new System.Windows.Forms.Label();
            this.grpVentana = new System.Windows.Forms.GroupBox();
            this.lblOrigenVentana = new System.Windows.Forms.Label();
            this.chkManualVentana = new System.Windows.Forms.CheckBox();
            this.dtpVentanaFin = new System.Windows.Forms.DateTimePicker();
            this.lblFechaFin = new System.Windows.Forms.Label();
            this.dtpVentanaInicio = new System.Windows.Forms.DateTimePicker();
            this.lblFechaInicio = new System.Windows.Forms.Label();
            this.grpSeleccion = new System.Windows.Forms.GroupBox();
            this.splitContenedorSim = new System.Windows.Forms.SplitContainer();
            this.pnlFiltros = new System.Windows.Forms.Panel();
            this.treeContextoAcademico = new System.Windows.Forms.TreeView();
            this.pnlCabecerasNivel6 = new System.Windows.Forms.Panel();
            this.sepNivel6A = new System.Windows.Forms.Panel();
            this.sepNivel6B = new System.Windows.Forms.Panel();
            this.pnlArbolToolbar = new System.Windows.Forms.Panel();
            this.pnlToolbarIzq = new System.Windows.Forms.FlowLayoutPanel();
            this.lblPerfilArbol = new System.Windows.Forms.Label();
            this.cmbArboles = new System.Windows.Forms.ComboBox();
            this.btnCargarArbol = new System.Windows.Forms.Button();
            this.btnNuevoArbol = new System.Windows.Forms.Button();
            this.btnSobrescribirArbol = new System.Windows.Forms.Button();
            this.btnEliminarArbol = new System.Windows.Forms.Button();
            this.sepToolbar1 = new System.Windows.Forms.Panel();
            this.btnExpandirArbol = new System.Windows.Forms.Button();
            this.btnContraerArbol = new System.Windows.Forms.Button();
            this.lblConteoCombinaciones = new System.Windows.Forms.Label();
            this.btnGenerarAulas = new System.Windows.Forms.Button();
            this.cmbSedes = new ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox();
            this.lblSede = new System.Windows.Forms.Label();
            this.cmbProgramas = new ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox();
            this.lblPrograma = new System.Windows.Forms.Label();
            this.cmbPeriodos = new ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox();
            this.lblPeriodo = new System.Windows.Forms.Label();
            this.lblTipoPrograma = new System.Windows.Forms.Label();
            this.cmbTipoPrograma = new System.Windows.Forms.ComboBox();
            this.btnFiltrarArbol = new System.Windows.Forms.Button();
            this.btnSeleccionarTerceros = new System.Windows.Forms.Button();
            this.lblColJornada = new System.Windows.Forms.Label();
            this.lblColDias = new System.Windows.Forms.Label();
            this.lblColHorasDia = new System.Windows.Forms.Label();
            this.pnlConfiguracionInferior = new System.Windows.Forms.Panel();
            this.pnlSeparadorConfig = new System.Windows.Forms.Panel();
            this.pnlBotonesSim = new System.Windows.Forms.Panel();
            this.pnlProgreso = new System.Windows.Forms.Panel();
            this.progressBarAG = new System.Windows.Forms.ProgressBar();
            this.lblStatusAG = new System.Windows.Forms.Label();
            this.flowAcciones = new System.Windows.Forms.FlowLayoutPanel();
            this.chkModoDebug = new System.Windows.Forms.CheckBox();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEjecutarAG = new System.Windows.Forms.Button();
            this.btnValidarFactibilidad = new System.Windows.Forms.Button();
            this.txtLogAG = new System.Windows.Forms.TextBox();
            this.toolTipConteo = new System.Windows.Forms.ToolTip(this.components);
            this.tabResultados = new System.Windows.Forms.TabPage();
            this.pnlAgModo = new System.Windows.Forms.Panel();
            this.btnAgVinculaciones = new System.Windows.Forms.Button();
            this.btnAgVerVinculaciones = new System.Windows.Forms.Button();
            this.btnAgSesiones = new System.Windows.Forms.Button();
            this.btnAprobarCorrida = new System.Windows.Forms.Button();
            this.btnRecalcularViolaciones = new System.Windows.Forms.Button();
            this.lblAgModo = new System.Windows.Forms.Label();
            this.pnlHeaderResultados = new System.Windows.Forms.Panel();
            this.lblHardViolaciones = new System.Windows.Forms.Label();
            this.lblMejorFitness = new System.Windows.Forms.Label();
            this.cmbCargarCorridas = new System.Windows.Forms.ComboBox();
            this.lblCargarCorridas = new System.Windows.Forms.Label();
            this.btnPromoverProduccion = new System.Windows.Forms.Button();
            this.splitContainerResultados = new System.Windows.Forms.SplitContainer();
            this.dgvHorariosGenerados = new System.Windows.Forms.DataGridView();
            this.dgvViolaciones = new System.Windows.Forms.DataGridView();
            this.lblVacantes = new System.Windows.Forms.Label();
            this.tabAuditoriaResultados = new System.Windows.Forms.TabControl();
            this.tabViolaciones = new System.Windows.Forms.TabPage();
            this.tabVacantes = new System.Windows.Forms.TabPage();
            this.pnlFiltroVacantes = new System.Windows.Forms.Panel();
            this.lblFiltroVacante = new System.Windows.Forms.Label();
            this.cmbMotivoVacante = new System.Windows.Forms.ComboBox();
            this.btnAccionVacante = new System.Windows.Forms.Button();
            this.lblConteoVacantesTab = new System.Windows.Forms.Label();
            this.dgvVacantes = new System.Windows.Forms.DataGridView();
            this.tabControlPrincipal.SuspendLayout();
            this.tabSimulacion.SuspendLayout();
            this.grpParametrosAG.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numElitismo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTasaMutacion)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTasaCruce)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numGeneraciones)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPoblacion)).BeginInit();
            this.grpVentana.SuspendLayout();
            this.grpSeleccion.SuspendLayout();
            this.pnlConfiguracionInferior.SuspendLayout();
            this.pnlBotonesSim.SuspendLayout();
            this.pnlProgreso.SuspendLayout();
            this.flowAcciones.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContenedorSim)).BeginInit();
            this.splitContenedorSim.Panel1.SuspendLayout();
            this.splitContenedorSim.Panel2.SuspendLayout();
            this.splitContenedorSim.SuspendLayout();
            this.pnlFiltros.SuspendLayout();
            this.pnlCabecerasNivel6.SuspendLayout();
            this.pnlArbolToolbar.SuspendLayout();
            this.pnlToolbarIzq.SuspendLayout();
            this.tabResultados.SuspendLayout();
            this.pnlHeaderResultados.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerResultados)).BeginInit();
            this.splitContainerResultados.Panel1.SuspendLayout();
            this.splitContainerResultados.Panel2.SuspendLayout();
            this.splitContainerResultados.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvHorariosGenerados)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvViolaciones)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControlPrincipal
            // 
            this.tabControlPrincipal.Controls.Add(this.tabSimulacion);
            this.tabControlPrincipal.Controls.Add(this.tabResultados);
            this.tabControlPrincipal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlPrincipal.Location = new System.Drawing.Point(0, 0);
            this.tabControlPrincipal.Name = "tabControlPrincipal";
            this.tabControlPrincipal.SelectedIndex = 0;
            this.tabControlPrincipal.Size = new System.Drawing.Size(1120, 661);
            this.tabControlPrincipal.TabIndex = 0;
            // 
            // tabSimulacion
            // 
            this.tabSimulacion.Controls.Add(this.splitContenedorSim);
            this.tabSimulacion.Location = new System.Drawing.Point(4, 22);
            this.tabSimulacion.Name = "tabSimulacion";
            this.tabSimulacion.Padding = new System.Windows.Forms.Padding(10);
            this.tabSimulacion.Size = new System.Drawing.Size(1112, 635);
            this.tabSimulacion.TabIndex = 0;
            this.tabSimulacion.Text = "Configuración & Ejecución (Módulo S)";
            this.tabSimulacion.UseVisualStyleBackColor = true;
            // 
            // splitContenedorSim
            // 
            this.splitContenedorSim.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContenedorSim.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContenedorSim.Location = new System.Drawing.Point(10, 10);
            this.splitContenedorSim.Name = "splitContenedorSim";
            this.splitContenedorSim.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContenedorSim.Panel1
            // 
            this.splitContenedorSim.Panel1.Controls.Add(this.grpSeleccion);
            this.splitContenedorSim.Panel1.Controls.Add(this.pnlConfiguracionInferior);
            // 
            // splitContenedorSim.Panel2
            // 
            this.splitContenedorSim.Panel2.Controls.Add(this.txtLogAG);
            this.splitContenedorSim.Panel2.Controls.Add(this.pnlBotonesSim);
            this.splitContenedorSim.Size = new System.Drawing.Size(1092, 615);
            this.splitContenedorSim.SplitterDistance = 525;
            this.splitContenedorSim.SplitterWidth = 6;
            this.splitContenedorSim.TabIndex = 0;
            // 
            // grpSeleccion
            // 
            // Orden crítico de adición: Fill primero, Top después en orden inverso al visual
            this.grpSeleccion.Controls.Add(this.treeContextoAcademico);
            this.grpSeleccion.Controls.Add(this.pnlCabecerasNivel6);
            this.grpSeleccion.Controls.Add(this.pnlArbolToolbar);
            this.grpSeleccion.Controls.Add(this.pnlFiltros);
            this.grpSeleccion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpSeleccion.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpSeleccion.Location = new System.Drawing.Point(0, 0);
            this.grpSeleccion.MinimumSize = new System.Drawing.Size(0, 190);
            this.grpSeleccion.Name = "grpSeleccion";
            this.grpSeleccion.Size = new System.Drawing.Size(1092, 360);
            this.grpSeleccion.TabIndex = 0;
            this.grpSeleccion.TabStop = false;
            this.grpSeleccion.Text = "1. Contexto Académico (Jerarquía 6 Niveles)";
            // 
            // treeContextoAcademico
            // 
            this.treeContextoAcademico.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.treeContextoAcademico.CheckBoxes = false;
            this.treeContextoAcademico.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeContextoAcademico.DrawMode = System.Windows.Forms.TreeViewDrawMode.OwnerDrawText;
            this.treeContextoAcademico.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.treeContextoAcademico.FullRowSelect = true;
            this.treeContextoAcademico.HideSelection = false;
            this.treeContextoAcademico.Indent = 14;
            this.treeContextoAcademico.ItemHeight = 20;
            this.treeContextoAcademico.Location = new System.Drawing.Point(3, 106);
            this.treeContextoAcademico.Name = "treeContextoAcademico";
            this.treeContextoAcademico.Scrollable = true;
            this.treeContextoAcademico.ShowNodeToolTips = true;
            this.treeContextoAcademico.Size = new System.Drawing.Size(1086, 360);
            this.treeContextoAcademico.TabIndex = 3;
            // 
            // pnlCabecerasNivel6
            // 
            this.pnlCabecerasNivel6.Controls.Add(this.lblColHorasDia);
            this.pnlCabecerasNivel6.Controls.Add(this.sepNivel6B);
            this.pnlCabecerasNivel6.Controls.Add(this.lblColDias);
            this.pnlCabecerasNivel6.Controls.Add(this.sepNivel6A);
            this.pnlCabecerasNivel6.Controls.Add(this.lblColJornada);
            this.pnlCabecerasNivel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlCabecerasNivel6.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.pnlCabecerasNivel6.Location = new System.Drawing.Point(3, 86);
            this.pnlCabecerasNivel6.Name = "pnlCabecerasNivel6";
            this.pnlCabecerasNivel6.Size = new System.Drawing.Size(1086, 20);
            this.pnlCabecerasNivel6.TabIndex = 2;
            this.pnlCabecerasNivel6.Visible = false;
            // 
            // lblColJornada
            // 
            this.lblColJornada.AutoSize = true;
            this.lblColJornada.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.lblColJornada.ForeColor = System.Drawing.Color.Navy;
            this.lblColJornada.Location = new System.Drawing.Point(150, 3);
            this.lblColJornada.Name = "lblColJornada";
            this.lblColJornada.Size = new System.Drawing.Size(52, 13);
            this.lblColJornada.TabIndex = 0;
            this.lblColJornada.Text = "Jornada:";
            // 
            // sepNivel6A
            // 
            this.sepNivel6A.BackColor = System.Drawing.Color.FromArgb(160, 160, 160);
            this.sepNivel6A.Location = new System.Drawing.Point(360, 2);
            this.sepNivel6A.Name = "sepNivel6A";
            this.sepNivel6A.Size = new System.Drawing.Size(1, 16);
            this.sepNivel6A.TabIndex = 1;
            // 
            // lblColDias
            // 
            this.lblColDias.AutoSize = true;
            this.lblColDias.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.lblColDias.ForeColor = System.Drawing.Color.Navy;
            this.lblColDias.Location = new System.Drawing.Point(375, 3);
            this.lblColDias.Name = "lblColDias";
            this.lblColDias.Size = new System.Drawing.Size(35, 13);
            this.lblColDias.TabIndex = 2;
            this.lblColDias.Text = "Días:";
            // 
            // sepNivel6B
            // 
            this.sepNivel6B.BackColor = System.Drawing.Color.FromArgb(160, 160, 160);
            this.sepNivel6B.Location = new System.Drawing.Point(775, 2);
            this.sepNivel6B.Name = "sepNivel6B";
            this.sepNivel6B.Size = new System.Drawing.Size(1, 16);
            this.sepNivel6B.TabIndex = 3;
            // 
            // lblColHorasDia
            // 
            this.lblColHorasDia.AutoSize = true;
            this.lblColHorasDia.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.lblColHorasDia.ForeColor = System.Drawing.Color.Navy;
            this.lblColHorasDia.Location = new System.Drawing.Point(790, 3);
            this.lblColHorasDia.Name = "lblColHorasDia";
            this.lblColHorasDia.Size = new System.Drawing.Size(61, 13);
            this.lblColHorasDia.TabIndex = 4;
            this.lblColHorasDia.Text = "Horas/día:";
            // 
            // pnlArbolToolbar
            // 
            this.pnlArbolToolbar.Controls.Add(this.lblConteoCombinaciones);
            this.pnlArbolToolbar.Controls.Add(this.btnGenerarAulas);
            this.pnlArbolToolbar.Controls.Add(this.pnlToolbarIzq);
            this.pnlArbolToolbar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlArbolToolbar.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.pnlArbolToolbar.Location = new System.Drawing.Point(3, 52);
            this.pnlArbolToolbar.Name = "pnlArbolToolbar";
            this.pnlArbolToolbar.Size = new System.Drawing.Size(1086, 34);
            this.pnlArbolToolbar.TabIndex = 1;
            // 
            // pnlToolbarIzq
            // 
            this.pnlToolbarIzq.AutoSize = true;
            this.pnlToolbarIzq.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pnlToolbarIzq.Controls.Add(this.lblPerfilArbol);
            this.pnlToolbarIzq.Controls.Add(this.cmbArboles);
            this.pnlToolbarIzq.Controls.Add(this.btnCargarArbol);
            this.pnlToolbarIzq.Controls.Add(this.btnNuevoArbol);
            this.pnlToolbarIzq.Controls.Add(this.btnSobrescribirArbol);
            this.pnlToolbarIzq.Controls.Add(this.btnEliminarArbol);
            this.pnlToolbarIzq.Controls.Add(this.sepToolbar1);
            this.pnlToolbarIzq.Controls.Add(this.btnExpandirArbol);
            this.pnlToolbarIzq.Controls.Add(this.btnContraerArbol);
            this.pnlToolbarIzq.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlToolbarIzq.Location = new System.Drawing.Point(0, 0);
            this.pnlToolbarIzq.Name = "pnlToolbarIzq";
            this.pnlToolbarIzq.Size = new System.Drawing.Size(703, 34);
            this.pnlToolbarIzq.TabIndex = 0;
            this.pnlToolbarIzq.WrapContents = false;
            // 
            // sepToolbar1
            // 
            this.sepToolbar1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(160)))), ((int)(((byte)(160)))), ((int)(((byte)(160)))));
            this.sepToolbar1.Location = new System.Drawing.Point(587, 5);
            this.sepToolbar1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 0);
            this.sepToolbar1.Name = "sepToolbar1";
            this.sepToolbar1.Size = new System.Drawing.Size(1, 20);
            this.sepToolbar1.TabIndex = 6;
            // 
            // lblPerfilArbol
            // 
            this.lblPerfilArbol.AutoSize = true;
            this.lblPerfilArbol.Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Bold);
            this.lblPerfilArbol.Location = new System.Drawing.Point(3, 8);
            this.lblPerfilArbol.Margin = new System.Windows.Forms.Padding(3, 8, 3, 0);
            this.lblPerfilArbol.Name = "lblPerfilArbol";
            this.lblPerfilArbol.Size = new System.Drawing.Size(73, 15);
            this.lblPerfilArbol.TabIndex = 0;
            this.lblPerfilArbol.Text = "Perfil Árbol:";
            // 
            // cmbArboles
            // 
            this.cmbArboles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArboles.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.cmbArboles.FormattingEnabled = true;
            this.cmbArboles.Location = new System.Drawing.Point(82, 5);
            this.cmbArboles.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.cmbArboles.Name = "cmbArboles";
            this.cmbArboles.Size = new System.Drawing.Size(165, 21);
            this.cmbArboles.TabIndex = 1;
            // 
            // btnCargarArbol
            // 
            this.btnCargarArbol.Location = new System.Drawing.Point(253, 3);
            this.btnCargarArbol.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnCargarArbol.Name = "btnCargarArbol";
            this.btnCargarArbol.Size = new System.Drawing.Size(58, 26);
            this.btnCargarArbol.TabIndex = 2;
            this.btnCargarArbol.Text = "Cargar";
            this.btnCargarArbol.UseVisualStyleBackColor = true;
            // 
            // btnNuevoArbol
            // 
            this.btnNuevoArbol.Location = new System.Drawing.Point(317, 3);
            this.btnNuevoArbol.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnNuevoArbol.Name = "btnNuevoArbol";
            this.btnNuevoArbol.Size = new System.Drawing.Size(105, 26);
            this.btnNuevoArbol.TabIndex = 3;
            this.btnNuevoArbol.Text = "Guardar como...";
            this.btnNuevoArbol.UseVisualStyleBackColor = true;
            // 
            // btnSobrescribirArbol
            // 
            this.btnSobrescribirArbol.Location = new System.Drawing.Point(428, 3);
            this.btnSobrescribirArbol.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnSobrescribirArbol.Name = "btnSobrescribirArbol";
            this.btnSobrescribirArbol.Size = new System.Drawing.Size(86, 26);
            this.btnSobrescribirArbol.TabIndex = 4;
            this.btnSobrescribirArbol.Text = "Sobrescribir";
            this.btnSobrescribirArbol.UseVisualStyleBackColor = true;
            // 
            // btnEliminarArbol
            // 
            this.btnEliminarArbol.Location = new System.Drawing.Point(520, 3);
            this.btnEliminarArbol.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnEliminarArbol.Name = "btnEliminarArbol";
            this.btnEliminarArbol.Size = new System.Drawing.Size(64, 26);
            this.btnEliminarArbol.TabIndex = 5;
            this.btnEliminarArbol.Text = "Eliminar";
            this.btnEliminarArbol.UseVisualStyleBackColor = true;
            // 
            // btnExpandirArbol
            // 
            this.btnExpandirArbol.Location = new System.Drawing.Point(590, 3);
            this.btnExpandirArbol.Margin = new System.Windows.Forms.Padding(3, 3, 2, 3);
            this.btnExpandirArbol.Name = "btnExpandirArbol";
            this.btnExpandirArbol.Size = new System.Drawing.Size(48, 26);
            this.btnExpandirArbol.TabIndex = 6;
            this.btnExpandirArbol.Text = "Exp.";
            this.btnExpandirArbol.UseVisualStyleBackColor = true;
            // 
            // btnContraerArbol
            // 
            this.btnContraerArbol.Location = new System.Drawing.Point(642, 3);
            this.btnContraerArbol.Margin = new System.Windows.Forms.Padding(2, 3, 3, 3);
            this.btnContraerArbol.Name = "btnContraerArbol";
            this.btnContraerArbol.Size = new System.Drawing.Size(48, 26);
            this.btnContraerArbol.TabIndex = 7;
            this.btnContraerArbol.Text = "Cont.";
            this.btnContraerArbol.UseVisualStyleBackColor = true;
            // 
            // btnGenerarAulas
            // 
            this.btnGenerarAulas.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnGenerarAulas.Location = new System.Drawing.Point(936, 0);
            this.btnGenerarAulas.Margin = new System.Windows.Forms.Padding(0);
            this.btnGenerarAulas.Name = "btnGenerarAulas";
            this.btnGenerarAulas.Size = new System.Drawing.Size(150, 34);
            this.btnGenerarAulas.TabIndex = 2;
            this.btnGenerarAulas.Text = "Generar Aulas (0)";
            this.btnGenerarAulas.UseVisualStyleBackColor = true;
            // 
            // lblConteoCombinaciones
            // 
            this.lblConteoCombinaciones.AutoEllipsis = true;
            this.lblConteoCombinaciones.AutoSize = false;
            this.lblConteoCombinaciones.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblConteoCombinaciones.Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Bold);
            this.lblConteoCombinaciones.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblConteoCombinaciones.Location = new System.Drawing.Point(696, 0);
            this.lblConteoCombinaciones.Name = "lblConteoCombinaciones";
            this.lblConteoCombinaciones.Padding = new System.Windows.Forms.Padding(10, 0, 5, 0);
            this.lblConteoCombinaciones.Size = new System.Drawing.Size(240, 34);
            this.lblConteoCombinaciones.TabIndex = 1;
            this.lblConteoCombinaciones.Text = "0 combinaciones seleccionadas";
            this.lblConteoCombinaciones.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlFiltros
            // 
            this.pnlFiltros.Controls.Add(this.btnSeleccionarTerceros);
            this.pnlFiltros.Controls.Add(this.btnFiltrarArbol);
            this.pnlFiltros.Controls.Add(this.cmbSedes);
            this.pnlFiltros.Controls.Add(this.lblSede);
            this.pnlFiltros.Controls.Add(this.cmbProgramas);
            this.pnlFiltros.Controls.Add(this.lblPrograma);
            this.pnlFiltros.Controls.Add(this.cmbPeriodos);
            this.pnlFiltros.Controls.Add(this.lblPeriodo);
            this.pnlFiltros.Controls.Add(this.cmbTipoPrograma);
            this.pnlFiltros.Controls.Add(this.lblTipoPrograma);
            this.pnlFiltros.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlFiltros.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.pnlFiltros.Location = new System.Drawing.Point(3, 19);
            this.pnlFiltros.Name = "pnlFiltros";
            this.pnlFiltros.Size = new System.Drawing.Size(1086, 33);
            this.pnlFiltros.TabIndex = 0;
            // 
            // lblTipoPrograma
            // 
            this.lblTipoPrograma.AutoSize = true;
            this.lblTipoPrograma.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblTipoPrograma.Location = new System.Drawing.Point(8, 8);
            this.lblTipoPrograma.Name = "lblTipoPrograma";
            this.lblTipoPrograma.Size = new System.Drawing.Size(33, 15);
            this.lblTipoPrograma.TabIndex = 0;
            this.lblTipoPrograma.Text = "Tipo:";
            // 
            // cmbTipoPrograma
            // 
            this.cmbTipoPrograma.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoPrograma.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbTipoPrograma.FormattingEnabled = true;
            this.cmbTipoPrograma.Location = new System.Drawing.Point(46, 5);
            this.cmbTipoPrograma.Name = "cmbTipoPrograma";
            this.cmbTipoPrograma.Size = new System.Drawing.Size(100, 23);
            this.cmbTipoPrograma.TabIndex = 1;
            // 
            // lblPeriodo
            // 
            this.lblPeriodo.AutoSize = true;
            this.lblPeriodo.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPeriodo.Location = new System.Drawing.Point(154, 8);
            this.lblPeriodo.Name = "lblPeriodo";
            this.lblPeriodo.Size = new System.Drawing.Size(51, 15);
            this.lblPeriodo.TabIndex = 2;
            this.lblPeriodo.Text = "Periodo:";
            // 
            // cmbPeriodos
            // 
            this.cmbPeriodos.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbPeriodos.Location = new System.Drawing.Point(208, 5);
            this.cmbPeriodos.Name = "cmbPeriodos";
            this.cmbPeriodos.Size = new System.Drawing.Size(95, 23);
            this.cmbPeriodos.TabIndex = 3;
            // 
            // lblPrograma
            // 
            this.lblPrograma.AutoSize = true;
            this.lblPrograma.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrograma.Location = new System.Drawing.Point(311, 8);
            this.lblPrograma.Name = "lblPrograma";
            this.lblPrograma.Size = new System.Drawing.Size(61, 15);
            this.lblPrograma.TabIndex = 4;
            this.lblPrograma.Text = "Programa:";
            // 
            // cmbProgramas
            // 
            this.cmbProgramas.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbProgramas.Location = new System.Drawing.Point(376, 5);
            this.cmbProgramas.Name = "cmbProgramas";
            this.cmbProgramas.Size = new System.Drawing.Size(225, 23);
            this.cmbProgramas.TabIndex = 5;
            // 
            // lblSede
            // 
            this.lblSede.AutoSize = true;
            this.lblSede.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblSede.Location = new System.Drawing.Point(609, 8);
            this.lblSede.Name = "lblSede";
            this.lblSede.Size = new System.Drawing.Size(35, 15);
            this.lblSede.TabIndex = 6;
            this.lblSede.Text = "Sede:";
            // 
            // cmbSedes
            // 
            this.cmbSedes.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbSedes.Location = new System.Drawing.Point(648, 5);
            this.cmbSedes.Name = "cmbSedes";
            this.cmbSedes.Size = new System.Drawing.Size(180, 23);
            this.cmbSedes.TabIndex = 7;
            // 
            // btnFiltrarArbol
            // 
            this.btnFiltrarArbol.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnFiltrarArbol.Location = new System.Drawing.Point(838, 4);
            this.btnFiltrarArbol.Name = "btnFiltrarArbol";
            this.btnFiltrarArbol.Size = new System.Drawing.Size(100, 25);
            this.btnFiltrarArbol.TabIndex = 8;
            this.btnFiltrarArbol.Text = "Filtrar Árbol";
            this.btnFiltrarArbol.UseVisualStyleBackColor = true;
            // 
            // btnSeleccionarTerceros
            // 
            this.btnSeleccionarTerceros.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnSeleccionarTerceros.Location = new System.Drawing.Point(946, 4);
            this.btnSeleccionarTerceros.Name = "btnSeleccionarTerceros";
            this.btnSeleccionarTerceros.Size = new System.Drawing.Size(110, 25);
            this.btnSeleccionarTerceros.TabIndex = 9;
            this.btnSeleccionarTerceros.Text = "Terceros...";
            this.btnSeleccionarTerceros.UseVisualStyleBackColor = true;
            // 
            // pnlConfiguracionInferior
            // 
            this.pnlConfiguracionInferior.Controls.Add(this.grpParametrosAG);
            this.pnlConfiguracionInferior.Controls.Add(this.pnlSeparadorConfig);
            this.pnlConfiguracionInferior.Controls.Add(this.grpVentana);
            this.pnlConfiguracionInferior.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlConfiguracionInferior.Location = new System.Drawing.Point(0, 323);
            this.pnlConfiguracionInferior.Name = "pnlConfiguracionInferior";
            this.pnlConfiguracionInferior.Size = new System.Drawing.Size(1092, 56);
            this.pnlConfiguracionInferior.TabIndex = 1;
            // 
            // pnlSeparadorConfig
            // 
            this.pnlSeparadorConfig.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlSeparadorConfig.Location = new System.Drawing.Point(490, 0);
            this.pnlSeparadorConfig.Name = "pnlSeparadorConfig";
            this.pnlSeparadorConfig.Size = new System.Drawing.Size(6, 56);
            this.pnlSeparadorConfig.TabIndex = 1;
            // 
            // grpVentana
            // 
            this.grpVentana.Controls.Add(this.lblOrigenVentana);
            this.grpVentana.Controls.Add(this.chkManualVentana);
            this.grpVentana.Controls.Add(this.dtpVentanaFin);
            this.grpVentana.Controls.Add(this.lblFechaFin);
            this.grpVentana.Controls.Add(this.dtpVentanaInicio);
            this.grpVentana.Controls.Add(this.lblFechaInicio);
            this.grpVentana.Dock = System.Windows.Forms.DockStyle.Left;
            this.grpVentana.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpVentana.Location = new System.Drawing.Point(0, 0);
            this.grpVentana.Name = "grpVentana";
            this.grpVentana.Size = new System.Drawing.Size(490, 56);
            this.grpVentana.TabIndex = 0;
            this.grpVentana.TabStop = false;
            this.grpVentana.Text = "2. Ventana Temporal Semestral (α)";
            // 
            // lblFechaInicio
            // 
            this.lblFechaInicio.AutoSize = true;
            this.lblFechaInicio.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblFechaInicio.Location = new System.Drawing.Point(8, 25);
            this.lblFechaInicio.Name = "lblFechaInicio";
            this.lblFechaInicio.Size = new System.Drawing.Size(71, 15);
            this.lblFechaInicio.TabIndex = 0;
            this.lblFechaInicio.Text = "Fecha Inicio:";
            // 
            // dtpVentanaInicio
            // 
            this.dtpVentanaInicio.CustomFormat = "dd/MM/yyyy";
            this.dtpVentanaInicio.Enabled = false;
            this.dtpVentanaInicio.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpVentanaInicio.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpVentanaInicio.Location = new System.Drawing.Point(82, 22);
            this.dtpVentanaInicio.Name = "dtpVentanaInicio";
            this.dtpVentanaInicio.Size = new System.Drawing.Size(100, 23);
            this.dtpVentanaInicio.TabIndex = 1;
            // 
            // lblFechaFin
            // 
            this.lblFechaFin.AutoSize = true;
            this.lblFechaFin.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblFechaFin.Location = new System.Drawing.Point(188, 25);
            this.lblFechaFin.Name = "lblFechaFin";
            this.lblFechaFin.Size = new System.Drawing.Size(55, 15);
            this.lblFechaFin.TabIndex = 2;
            this.lblFechaFin.Text = "Fecha Fin:";
            // 
            // dtpVentanaFin
            // 
            this.dtpVentanaFin.CustomFormat = "dd/MM/yyyy";
            this.dtpVentanaFin.Enabled = false;
            this.dtpVentanaFin.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpVentanaFin.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpVentanaFin.Location = new System.Drawing.Point(247, 22);
            this.dtpVentanaFin.Name = "dtpVentanaFin";
            this.dtpVentanaFin.Size = new System.Drawing.Size(100, 23);
            this.dtpVentanaFin.TabIndex = 3;
            // 
            // chkManualVentana
            // 
            this.chkManualVentana.AutoSize = true;
            this.chkManualVentana.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkManualVentana.Location = new System.Drawing.Point(355, 24);
            this.chkManualVentana.Name = "chkManualVentana";
            this.chkManualVentana.Size = new System.Drawing.Size(105, 19);
            this.chkManualVentana.TabIndex = 4;
            this.chkManualVentana.Text = "Ajustar manual";
            this.chkManualVentana.UseVisualStyleBackColor = true;
            // 
            // lblOrigenVentana
            // 
            this.lblOrigenVentana.AutoSize = true;
            this.lblOrigenVentana.Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Italic);
            this.lblOrigenVentana.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblOrigenVentana.Location = new System.Drawing.Point(206, 0);
            this.lblOrigenVentana.Name = "lblOrigenVentana";
            this.lblOrigenVentana.Size = new System.Drawing.Size(185, 15);
            this.lblOrigenVentana.TabIndex = 5;
            this.lblOrigenVentana.Text = "· Origen: Festivos (PRESENCIAL)";
            // 
            // grpParametrosAG
            // 
            this.grpParametrosAG.Controls.Add(this.numElitismo);
            this.grpParametrosAG.Controls.Add(this.lblElitismo);
            this.grpParametrosAG.Controls.Add(this.numTasaMutacion);
            this.grpParametrosAG.Controls.Add(this.lblMutacion);
            this.grpParametrosAG.Controls.Add(this.numTasaCruce);
            this.grpParametrosAG.Controls.Add(this.lblCruce);
            this.grpParametrosAG.Controls.Add(this.numGeneraciones);
            this.grpParametrosAG.Controls.Add(this.lblGeneraciones);
            this.grpParametrosAG.Controls.Add(this.numPoblacion);
            this.grpParametrosAG.Controls.Add(this.lblPoblacion);
            this.grpParametrosAG.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpParametrosAG.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpParametrosAG.Location = new System.Drawing.Point(496, 0);
            this.grpParametrosAG.Name = "grpParametrosAG";
            this.grpParametrosAG.Size = new System.Drawing.Size(600, 56);
            this.grpParametrosAG.TabIndex = 2;
            this.grpParametrosAG.TabStop = false;
            this.grpParametrosAG.Text = "3. Parámetros Evolutivos";
            // 
            // lblPoblacion
            // 
            this.lblPoblacion.AutoSize = true;
            this.lblPoblacion.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblPoblacion.Location = new System.Drawing.Point(8, 25);
            this.lblPoblacion.Name = "lblPoblacion";
            this.lblPoblacion.Size = new System.Drawing.Size(63, 15);
            this.lblPoblacion.TabIndex = 0;
            this.lblPoblacion.Text = "Población:";
            // 
            // numPoblacion
            // 
            this.numPoblacion.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numPoblacion.Location = new System.Drawing.Point(74, 22);
            this.numPoblacion.Maximum = new decimal(new int[] { 500, 0, 0, 0 });
            this.numPoblacion.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
            this.numPoblacion.Name = "numPoblacion";
            this.numPoblacion.Size = new System.Drawing.Size(50, 23);
            this.numPoblacion.TabIndex = 1;
            this.numPoblacion.Value = new decimal(new int[] { 50, 0, 0, 0 });
            // 
            // lblGeneraciones
            // 
            this.lblGeneraciones.AutoSize = true;
            this.lblGeneraciones.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblGeneraciones.Location = new System.Drawing.Point(128, 25);
            this.lblGeneraciones.Name = "lblGeneraciones";
            this.lblGeneraciones.Size = new System.Drawing.Size(42, 15);
            this.lblGeneraciones.TabIndex = 2;
            this.lblGeneraciones.Text = "Gener.:";
            // 
            // numGeneraciones
            // 
            this.numGeneraciones.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numGeneraciones.Location = new System.Drawing.Point(173, 22);
            this.numGeneraciones.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.numGeneraciones.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
            this.numGeneraciones.Name = "numGeneraciones";
            this.numGeneraciones.Size = new System.Drawing.Size(55, 23);
            this.numGeneraciones.TabIndex = 3;
            this.numGeneraciones.Value = new decimal(new int[] { 200, 0, 0, 0 });
            // 
            // lblCruce
            // 
            this.lblCruce.AutoSize = true;
            this.lblCruce.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblCruce.Location = new System.Drawing.Point(232, 25);
            this.lblCruce.Name = "lblCruce";
            this.lblCruce.Size = new System.Drawing.Size(41, 15);
            this.lblCruce.TabIndex = 4;
            this.lblCruce.Text = "Cruce:";
            // 
            // numTasaCruce
            // 
            this.numTasaCruce.DecimalPlaces = 2;
            this.numTasaCruce.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numTasaCruce.Increment = new decimal(new int[] { 5, 0, 0, 131072 });
            this.numTasaCruce.Location = new System.Drawing.Point(276, 22);
            this.numTasaCruce.Maximum = new decimal(new int[] { 1, 0, 0, 0 });
            this.numTasaCruce.Minimum = new decimal(new int[] { 1, 0, 0, 65536 });
            this.numTasaCruce.Name = "numTasaCruce";
            this.numTasaCruce.Size = new System.Drawing.Size(50, 23);
            this.numTasaCruce.TabIndex = 5;
            this.numTasaCruce.Value = new decimal(new int[] { 85, 0, 0, 131072 });
            // 
            // lblMutacion
            // 
            this.lblMutacion.AutoSize = true;
            this.lblMutacion.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblMutacion.Location = new System.Drawing.Point(330, 25);
            this.lblMutacion.Name = "lblMutacion";
            this.lblMutacion.Size = new System.Drawing.Size(44, 15);
            this.lblMutacion.TabIndex = 6;
            this.lblMutacion.Text = "Mutac.:";
            // 
            // numTasaMutacion
            // 
            this.numTasaMutacion.DecimalPlaces = 2;
            this.numTasaMutacion.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numTasaMutacion.Increment = new decimal(new int[] { 5, 0, 0, 131072 });
            this.numTasaMutacion.Location = new System.Drawing.Point(377, 22);
            this.numTasaMutacion.Maximum = new decimal(new int[] { 50, 0, 0, 131072 });
            this.numTasaMutacion.Minimum = new decimal(new int[] { 1, 0, 0, 131072 });
            this.numTasaMutacion.Name = "numTasaMutacion";
            this.numTasaMutacion.Size = new System.Drawing.Size(50, 23);
            this.numTasaMutacion.TabIndex = 7;
            this.numTasaMutacion.Value = new decimal(new int[] { 10, 0, 0, 131072 });
            // 
            // lblElitismo
            // 
            this.lblElitismo.AutoSize = true;
            this.lblElitismo.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblElitismo.Location = new System.Drawing.Point(431, 25);
            this.lblElitismo.Name = "lblElitismo";
            this.lblElitismo.Size = new System.Drawing.Size(51, 15);
            this.lblElitismo.TabIndex = 8;
            this.lblElitismo.Text = "Elitismo:";
            // 
            // numElitismo
            // 
            this.numElitismo.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numElitismo.Location = new System.Drawing.Point(485, 22);
            this.numElitismo.Maximum = new decimal(new int[] { 10, 0, 0, 0 });
            this.numElitismo.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.numElitismo.Name = "numElitismo";
            this.numElitismo.Size = new System.Drawing.Size(42, 23);
            this.numElitismo.TabIndex = 9;
            this.numElitismo.Value = new decimal(new int[] { 2, 0, 0, 0 });
            // 
            // pnlBotonesSim
            // 
            this.pnlBotonesSim.Controls.Add(this.pnlProgreso);
            this.pnlBotonesSim.Controls.Add(this.flowAcciones);
            this.pnlBotonesSim.Controls.Add(this.chkModoDebug);
            this.pnlBotonesSim.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlBotonesSim.Location = new System.Drawing.Point(0, 0);
            this.pnlBotonesSim.Name = "pnlBotonesSim";
            this.pnlBotonesSim.Size = new System.Drawing.Size(1092, 38);
            this.pnlBotonesSim.TabIndex = 0;
            // 
            // pnlProgreso
            // 
            this.pnlProgreso.Controls.Add(this.lblStatusAG);
            this.pnlProgreso.Controls.Add(this.progressBarAG);
            this.pnlProgreso.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlProgreso.Location = new System.Drawing.Point(0, 0);
            this.pnlProgreso.Name = "pnlProgreso";
            this.pnlProgreso.Padding = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.pnlProgreso.Size = new System.Drawing.Size(430, 38);
            this.pnlProgreso.TabIndex = 0;
            // 
            // progressBarAG
            // 
            this.progressBarAG.Location = new System.Drawing.Point(5, 9);
            this.progressBarAG.Name = "progressBarAG";
            this.progressBarAG.Size = new System.Drawing.Size(130, 20);
            this.progressBarAG.TabIndex = 0;
            // 
            // lblStatusAG
            // 
            this.lblStatusAG.AutoEllipsis = true;
            this.lblStatusAG.AutoSize = false;
            this.lblStatusAG.Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Bold);
            this.lblStatusAG.Location = new System.Drawing.Point(142, 9);
            this.lblStatusAG.Name = "lblStatusAG";
            this.lblStatusAG.Size = new System.Drawing.Size(280, 20);
            this.lblStatusAG.TabIndex = 1;
            this.lblStatusAG.Text = "Estado: En espera";
            this.lblStatusAG.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkModoDebug
            // 
            this.chkModoDebug.AutoEllipsis = true;
            this.chkModoDebug.AutoSize = false;
            this.chkModoDebug.Checked = true;
            this.chkModoDebug.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkModoDebug.Dock = System.Windows.Forms.DockStyle.Right;
            this.chkModoDebug.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.chkModoDebug.Location = new System.Drawing.Point(912, 0);
            this.chkModoDebug.Name = "chkModoDebug";
            this.chkModoDebug.Size = new System.Drawing.Size(180, 38);
            this.chkModoDebug.TabIndex = 2;
            this.chkModoDebug.Text = "Modo Diagnóstico (DEBUG)";
            this.chkModoDebug.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.chkModoDebug.UseVisualStyleBackColor = true;
            // 
            // flowAcciones
            // 
            this.flowAcciones.AutoSize = true;
            this.flowAcciones.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flowAcciones.Controls.Add(this.btnCancelar);
            this.flowAcciones.Controls.Add(this.btnEjecutarAG);
            this.flowAcciones.Controls.Add(this.btnValidarFactibilidad);
            this.flowAcciones.Dock = System.Windows.Forms.DockStyle.Right;
            this.flowAcciones.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowAcciones.Location = new System.Drawing.Point(482, 0);
            this.flowAcciones.Name = "flowAcciones";
            this.flowAcciones.Padding = new System.Windows.Forms.Padding(0, 4, 8, 4);
            this.flowAcciones.Size = new System.Drawing.Size(430, 38);
            this.flowAcciones.TabIndex = 1;
            this.flowAcciones.WrapContents = false;
            // 
            // btnCancelar
            // 
            this.btnCancelar.Enabled = false;
            this.btnCancelar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnCancelar.Location = new System.Drawing.Point(327, 4);
            this.btnCancelar.Margin = new System.Windows.Forms.Padding(4, 0, 0, 0);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(95, 30);
            this.btnCancelar.TabIndex = 2;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            // 
            // btnEjecutarAG
            // 
            this.btnEjecutarAG.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(167)))), ((int)(((byte)(69)))));
            this.btnEjecutarAG.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnEjecutarAG.ForeColor = System.Drawing.Color.White;
            this.btnEjecutarAG.Location = new System.Drawing.Point(159, 4);
            this.btnEjecutarAG.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.btnEjecutarAG.Name = "btnEjecutarAG";
            this.btnEjecutarAG.Size = new System.Drawing.Size(160, 30);
            this.btnEjecutarAG.TabIndex = 1;
            this.btnEjecutarAG.Text = "Optimizar Horarios";
            this.btnEjecutarAG.UseVisualStyleBackColor = false;
            // 
            // btnValidarFactibilidad
            // 
            this.btnValidarFactibilidad.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnValidarFactibilidad.Location = new System.Drawing.Point(0, 4);
            this.btnValidarFactibilidad.Margin = new System.Windows.Forms.Padding(0, 0, 4, 0);
            this.btnValidarFactibilidad.Name = "btnValidarFactibilidad";
            this.btnValidarFactibilidad.Size = new System.Drawing.Size(155, 30);
            this.btnValidarFactibilidad.TabIndex = 0;
            this.btnValidarFactibilidad.Text = "Diagnóstico Pre-vuelo";
            this.btnValidarFactibilidad.UseVisualStyleBackColor = true;
            // 
            // txtLogAG
            // 
            this.txtLogAG.BackColor = System.Drawing.Color.Black;
            this.txtLogAG.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLogAG.Font = new System.Drawing.Font("Consolas", 9F);
            this.txtLogAG.ForeColor = System.Drawing.Color.LimeGreen;
            this.txtLogAG.Location = new System.Drawing.Point(0, 38);
            this.txtLogAG.Multiline = true;
            this.txtLogAG.Name = "txtLogAG";
            this.txtLogAG.ReadOnly = true;
            this.txtLogAG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLogAG.Size = new System.Drawing.Size(1092, 46);
            this.txtLogAG.TabIndex = 1;
            this.txtLogAG.WordWrap = false;
            // 
            // tabResultados
            // 
            this.tabResultados.Controls.Add(this.splitContainerResultados);
            this.tabResultados.Controls.Add(this.pnlAgModo);
            this.tabResultados.Controls.Add(this.pnlHeaderResultados);
            this.tabResultados.Location = new System.Drawing.Point(4, 22);
            this.tabResultados.Name = "tabResultados";
            this.tabResultados.Padding = new System.Windows.Forms.Padding(10);
            this.tabResultados.Size = new System.Drawing.Size(976, 635);
            this.tabResultados.TabIndex = 1;
            this.tabResultados.Text = "Auditoría & Promoción (Módulo R)";
            this.tabResultados.UseVisualStyleBackColor = true;
            // 
            // pnlAgModo
            // 
            this.pnlAgModo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(248)))), ((int)(((byte)(240)))));
            this.pnlAgModo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlAgModo.Controls.Add(this.lblAgModo);
            this.pnlAgModo.Controls.Add(this.btnRecalcularViolaciones);
            this.pnlAgModo.Controls.Add(this.btnAgSesiones);
            this.pnlAgModo.Controls.Add(this.btnAgVerVinculaciones);
            this.pnlAgModo.Controls.Add(this.btnAgVinculaciones);
            this.pnlAgModo.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlAgModo.Location = new System.Drawing.Point(10, 60);
            this.pnlAgModo.Name = "pnlAgModo";
            this.pnlAgModo.Size = new System.Drawing.Size(956, 66);
            this.pnlAgModo.TabIndex = 2;
            // 
            // btnAgVinculaciones
            // 
            this.btnAgVinculaciones.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAgVinculaciones.Location = new System.Drawing.Point(10, 4);
            this.btnAgVinculaciones.Name = "btnAgVinculaciones";
            this.btnAgVinculaciones.Size = new System.Drawing.Size(170, 28);
            this.btnAgVinculaciones.TabIndex = 0;
            this.btnAgVinculaciones.Text = "🔗 Vinculaciones (AG)";
            this.btnAgVinculaciones.UseVisualStyleBackColor = true;
            // 
            // btnAgVerVinculaciones
            // 
            this.btnAgVerVinculaciones.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAgVerVinculaciones.Location = new System.Drawing.Point(190, 4);
            this.btnAgVerVinculaciones.Name = "btnAgVerVinculaciones";
            this.btnAgVerVinculaciones.Size = new System.Drawing.Size(150, 28);
            this.btnAgVerVinculaciones.TabIndex = 1;
            this.btnAgVerVinculaciones.Text = "📋 Detalle (AG)";
            this.btnAgVerVinculaciones.UseVisualStyleBackColor = true;
            // 
            // btnAgSesiones
            // 
            this.btnAgSesiones.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAgSesiones.Location = new System.Drawing.Point(350, 4);
            this.btnAgSesiones.Name = "btnAgSesiones";
            this.btnAgSesiones.Size = new System.Drawing.Size(150, 28);
            this.btnAgSesiones.TabIndex = 2;
            this.btnAgSesiones.Text = "✏️ Sesión (AG)";
            this.btnAgSesiones.UseVisualStyleBackColor = true;
            // 
            // btnAprobarCorrida
            // 
            this.btnAprobarCorrida.Enabled = false;
            this.btnAprobarCorrida.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAprobarCorrida.Location = new System.Drawing.Point(510, 4);
            this.btnAprobarCorrida.Name = "btnAprobarCorrida";
            this.btnAprobarCorrida.Size = new System.Drawing.Size(150, 28);
            this.btnAprobarCorrida.TabIndex = 3;
            this.btnAprobarCorrida.Text = "✅ Aprobar corrida";
            this.btnAprobarCorrida.UseVisualStyleBackColor = true;
            // 
            // 
            // btnRecalcularViolaciones
            // 
            this.btnRecalcularViolaciones.Enabled = false;
            this.btnRecalcularViolaciones.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRecalcularViolaciones.Location = new System.Drawing.Point(10, 34);
            this.btnRecalcularViolaciones.Name = "btnRecalcularViolaciones";
            this.btnRecalcularViolaciones.Size = new System.Drawing.Size(200, 28);
            this.btnRecalcularViolaciones.TabIndex = 4;
            this.btnRecalcularViolaciones.Text = "🔁 Recalcular violaciones";
            this.btnRecalcularViolaciones.UseVisualStyleBackColor = true;
            // lblAgModo
            // 
            this.lblAgModo.AutoEllipsis = true;
            this.lblAgModo.AutoSize = false;
            this.lblAgModo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAgModo.ForeColor = System.Drawing.Color.OrangeRed;
            this.lblAgModo.Location = new System.Drawing.Point(220, 40);
            this.lblAgModo.Name = "lblAgModo";
            this.lblAgModo.Size = new System.Drawing.Size(726, 15);
            this.lblAgModo.TabIndex = 3;
            this.lblAgModo.Text = "Modo AG · staging: ag_vinculaciones / ag_sesiones";
            // 
            // pnlHeaderResultados
            // 
            this.pnlHeaderResultados.Controls.Add(this.cmbCargarCorridas);
            this.pnlHeaderResultados.Controls.Add(this.lblCargarCorridas);
            this.pnlHeaderResultados.Controls.Add(this.btnPromoverProduccion);
            this.pnlHeaderResultados.Controls.Add(this.lblVacantes);
            this.pnlHeaderResultados.Controls.Add(this.lblHardViolaciones);
            this.pnlHeaderResultados.Controls.Add(this.lblMejorFitness);
            this.pnlHeaderResultados.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlHeaderResultados.Location = new System.Drawing.Point(10, 10);
            this.pnlHeaderResultados.Name = "pnlHeaderResultados";
            this.pnlHeaderResultados.Size = new System.Drawing.Size(1100, 50);
            this.pnlHeaderResultados.TabIndex = 0;
            // 
            // lblHardViolaciones
            // 
            this.lblHardViolaciones.AutoSize = true;
            this.lblHardViolaciones.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblHardViolaciones.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblHardViolaciones.Location = new System.Drawing.Point(195, 15);
            this.lblHardViolaciones.Name = "lblHardViolaciones";
            this.lblHardViolaciones.Size = new System.Drawing.Size(147, 19);
            this.lblHardViolaciones.TabIndex = 1;
            this.lblHardViolaciones.Text = "Restricciones Hard: 0";
            // 
            // lblMejorFitness
            // 
            this.lblMejorFitness.AutoSize = true;
            this.lblMejorFitness.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblMejorFitness.Location = new System.Drawing.Point(15, 15);
            this.lblMejorFitness.Name = "lblMejorFitness";
            this.lblMejorFitness.Size = new System.Drawing.Size(155, 19);
            this.lblMejorFitness.TabIndex = 0;
            this.lblMejorFitness.Text = "Aptitud (Fitness): 1.000";
            // 
            // lblVacantes
            // 
            this.lblVacantes.AutoSize = true;
            this.lblVacantes.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblVacantes.ForeColor = System.Drawing.Color.DarkOrange;
            this.lblVacantes.Location = new System.Drawing.Point(355, 15);
            this.lblVacantes.Name = "lblVacantes";
            this.lblVacantes.Size = new System.Drawing.Size(85, 19);
            this.lblVacantes.TabIndex = 4;
            this.lblVacantes.Text = "Vacantes: 0";
            // 
            // cmbCargarCorridas
            // 
            this.cmbCargarCorridas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCargarCorridas.DropDownWidth = 480;
            this.cmbCargarCorridas.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbCargarCorridas.FormattingEnabled = true;
            this.cmbCargarCorridas.Location = new System.Drawing.Point(545, 14);
            this.cmbCargarCorridas.Name = "cmbCargarCorridas";
            this.cmbCargarCorridas.Size = new System.Drawing.Size(330, 25);
            this.cmbCargarCorridas.TabIndex = 4;
            // 
            // lblCargarCorridas
            // 
            this.lblCargarCorridas.AutoSize = true;
            this.lblCargarCorridas.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCargarCorridas.Location = new System.Drawing.Point(444, 18);
            this.lblCargarCorridas.Name = "lblCargarCorridas";
            this.lblCargarCorridas.Size = new System.Drawing.Size(95, 15);
            this.lblCargarCorridas.TabIndex = 3;
            this.lblCargarCorridas.Text = "Cargar corridas:";
            // 
            // btnPromoverProduccion
            // 
            this.btnPromoverProduccion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPromoverProduccion.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnPromoverProduccion.Enabled = false;
            this.btnPromoverProduccion.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnPromoverProduccion.ForeColor = System.Drawing.Color.White;
            this.btnPromoverProduccion.Location = new System.Drawing.Point(750, 10);
            this.btnPromoverProduccion.Name = "btnPromoverProduccion";
            this.btnPromoverProduccion.Size = new System.Drawing.Size(200, 32);
            this.btnPromoverProduccion.TabIndex = 2;
            this.btnPromoverProduccion.Text = "🚀 Promover a Producción";
            this.btnPromoverProduccion.UseVisualStyleBackColor = false;
            // 
            // splitContainerResultados
            // 
            this.splitContainerResultados.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerResultados.Location = new System.Drawing.Point(10, 60);
            this.splitContainerResultados.Name = "splitContainerResultados";
            this.splitContainerResultados.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainerResultados.Panel1
            // 
            this.splitContainerResultados.Panel1.Controls.Add(this.dgvHorariosGenerados);
            // 
            // splitContainerResultados.Panel2
            // 
            this.splitContainerResultados.Panel2.Controls.Add(this.tabAuditoriaResultados);
            this.splitContainerResultados.Size = new System.Drawing.Size(956, 565);
            this.splitContainerResultados.SplitterDistance = 380;
            this.splitContainerResultados.TabIndex = 1;
            // 
            // tabAuditoriaResultados
            // 
            this.tabAuditoriaResultados.Controls.Add(this.tabViolaciones);
            this.tabAuditoriaResultados.Controls.Add(this.tabVacantes);
            this.tabAuditoriaResultados.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabAuditoriaResultados.Location = new System.Drawing.Point(0, 0);
            this.tabAuditoriaResultados.Name = "tabAuditoriaResultados";
            this.tabAuditoriaResultados.SelectedIndex = 0;
            this.tabAuditoriaResultados.Size = new System.Drawing.Size(956, 181);
            this.tabAuditoriaResultados.TabIndex = 0;
            // 
            // tabViolaciones
            // 
            this.tabViolaciones.Controls.Add(this.dgvViolaciones);
            this.tabViolaciones.Location = new System.Drawing.Point(4, 22);
            this.tabViolaciones.Name = "tabViolaciones";
            this.tabViolaciones.Padding = new System.Windows.Forms.Padding(3);
            this.tabViolaciones.Size = new System.Drawing.Size(948, 155);
            this.tabViolaciones.TabIndex = 0;
            this.tabViolaciones.Text = "⚠️ Violaciones de Reglas";
            this.tabViolaciones.UseVisualStyleBackColor = true;
            // 
            // tabVacantes
            // 
            this.tabVacantes.Controls.Add(this.dgvVacantes);
            this.tabVacantes.Controls.Add(this.pnlFiltroVacantes);
            this.tabVacantes.Location = new System.Drawing.Point(4, 22);
            this.tabVacantes.Name = "tabVacantes";
            this.tabVacantes.Padding = new System.Windows.Forms.Padding(3);
            this.tabVacantes.Size = new System.Drawing.Size(948, 155);
            this.tabVacantes.TabIndex = 1;
            this.tabVacantes.Text = "👤 Panel de Vacantes";
            this.tabVacantes.UseVisualStyleBackColor = true;
            // 
            // pnlFiltroVacantes
            // 
            this.pnlFiltroVacantes.Controls.Add(this.lblConteoVacantesTab);
            this.pnlFiltroVacantes.Controls.Add(this.btnAccionVacante);
            this.pnlFiltroVacantes.Controls.Add(this.cmbMotivoVacante);
            this.pnlFiltroVacantes.Controls.Add(this.lblFiltroVacante);
            this.pnlFiltroVacantes.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlFiltroVacantes.Location = new System.Drawing.Point(3, 3);
            this.pnlFiltroVacantes.Name = "pnlFiltroVacantes";
            this.pnlFiltroVacantes.Size = new System.Drawing.Size(942, 34);
            this.pnlFiltroVacantes.TabIndex = 0;
            // 
            // lblFiltroVacante
            // 
            this.lblFiltroVacante.AutoSize = true;
            this.lblFiltroVacante.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblFiltroVacante.Location = new System.Drawing.Point(8, 9);
            this.lblFiltroVacante.Name = "lblFiltroVacante";
            this.lblFiltroVacante.Size = new System.Drawing.Size(107, 15);
            this.lblFiltroVacante.TabIndex = 0;
            this.lblFiltroVacante.Text = "Filtrar por motivo:";
            // 
            // cmbMotivoVacante
            // 
            this.cmbMotivoVacante.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMotivoVacante.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbMotivoVacante.FormattingEnabled = true;
            this.cmbMotivoVacante.Location = new System.Drawing.Point(120, 6);
            this.cmbMotivoVacante.Name = "cmbMotivoVacante";
            this.cmbMotivoVacante.Size = new System.Drawing.Size(180, 23);
            this.cmbMotivoVacante.TabIndex = 1;
            // 
            // btnAccionVacante
            // 
            this.btnAccionVacante.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnAccionVacante.Location = new System.Drawing.Point(310, 4);
            this.btnAccionVacante.Name = "btnAccionVacante";
            this.btnAccionVacante.Size = new System.Drawing.Size(140, 26);
            this.btnAccionVacante.TabIndex = 2;
            this.btnAccionVacante.Text = "Cargar pool";
            this.btnAccionVacante.UseVisualStyleBackColor = true;
            // 
            // lblConteoVacantesTab
            // 
            this.lblConteoVacantesTab.AutoSize = true;
            this.lblConteoVacantesTab.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Italic);
            this.lblConteoVacantesTab.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.lblConteoVacantesTab.Location = new System.Drawing.Point(460, 9);
            this.lblConteoVacantesTab.Name = "lblConteoVacantesTab";
            this.lblConteoVacantesTab.Size = new System.Drawing.Size(71, 15);
            this.lblConteoVacantesTab.TabIndex = 3;
            this.lblConteoVacantesTab.Text = "0 vacante(s)";
            // 
            // dgvVacantes
            // 
            this.dgvVacantes.AllowUserToAddRows = false;
            this.dgvVacantes.AllowUserToDeleteRows = false;
            this.dgvVacantes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvVacantes.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.dgvVacantes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvVacantes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvVacantes.Location = new System.Drawing.Point(3, 37);
            this.dgvVacantes.Name = "dgvVacantes";
            this.dgvVacantes.ReadOnly = true;
            this.dgvVacantes.RowHeadersVisible = false;
            this.dgvVacantes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvVacantes.Size = new System.Drawing.Size(942, 115);
            this.dgvVacantes.TabIndex = 1;
            // 
            // dgvHorariosGenerados
            // 
            this.dgvHorariosGenerados.AllowUserToAddRows = false;
            this.dgvHorariosGenerados.AllowUserToDeleteRows = false;
            this.dgvHorariosGenerados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvHorariosGenerados.BackgroundColor = System.Drawing.Color.White;
            this.dgvHorariosGenerados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvHorariosGenerados.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvHorariosGenerados.Location = new System.Drawing.Point(0, 0);
            this.dgvHorariosGenerados.Name = "dgvHorariosGenerados";
            this.dgvHorariosGenerados.ReadOnly = true;
            this.dgvHorariosGenerados.RowHeadersVisible = false;
            this.dgvHorariosGenerados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvHorariosGenerados.Size = new System.Drawing.Size(956, 380);
            this.dgvHorariosGenerados.TabIndex = 0;
            // 
            // dgvViolaciones
            // 
            this.dgvViolaciones.AllowUserToAddRows = false;
            this.dgvViolaciones.AllowUserToDeleteRows = false;
            this.dgvViolaciones.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvViolaciones.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.dgvViolaciones.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvViolaciones.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvViolaciones.Location = new System.Drawing.Point(3, 3);
            this.dgvViolaciones.Name = "dgvViolaciones";
            this.dgvViolaciones.ReadOnly = true;
            this.dgvViolaciones.RowHeadersVisible = false;
            this.dgvViolaciones.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvViolaciones.Size = new System.Drawing.Size(942, 149);
            this.dgvViolaciones.TabIndex = 0;
            // 
            // AgFormGeneradorHorarios
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1120, 720);
            this.Controls.Add(this.tabControlPrincipal);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimumSize = new System.Drawing.Size(1120, 700);
            this.Name = "AgFormGeneradorHorarios";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Optimización Automática de Horarios Académicos (Algoritmo Genético)";
            this.tabControlPrincipal.ResumeLayout(false);
            this.tabSimulacion.ResumeLayout(false);
            this.grpParametrosAG.ResumeLayout(false);
            this.grpParametrosAG.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numElitismo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTasaMutacion)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTasaCruce)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numGeneraciones)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPoblacion)).EndInit();
            this.grpVentana.ResumeLayout(false);
            this.grpVentana.PerformLayout();
            this.grpSeleccion.ResumeLayout(false);
            this.pnlFiltros.ResumeLayout(false);
            this.pnlFiltros.PerformLayout();
            this.pnlCabecerasNivel6.ResumeLayout(false);
            this.pnlCabecerasNivel6.PerformLayout();
            this.pnlArbolToolbar.ResumeLayout(false);
            this.pnlToolbarIzq.ResumeLayout(false);
            this.pnlToolbarIzq.PerformLayout();
            this.pnlConfiguracionInferior.ResumeLayout(false);
            this.pnlBotonesSim.ResumeLayout(false);
            this.pnlBotonesSim.PerformLayout();
            this.pnlProgreso.ResumeLayout(false);
            this.flowAcciones.ResumeLayout(false);
            this.splitContenedorSim.Panel1.ResumeLayout(false);
            this.splitContenedorSim.Panel2.ResumeLayout(false);
            this.splitContenedorSim.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContenedorSim)).EndInit();
            this.splitContenedorSim.ResumeLayout(false);
            this.tabResultados.ResumeLayout(false);
            this.pnlAgModo.ResumeLayout(false);
            this.pnlAgModo.PerformLayout();
            this.pnlHeaderResultados.ResumeLayout(false);
            this.pnlHeaderResultados.PerformLayout();
            this.splitContainerResultados.Panel1.ResumeLayout(false);
            this.splitContainerResultados.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerResultados)).EndInit();
            this.splitContainerResultados.ResumeLayout(false);
            this.tabAuditoriaResultados.ResumeLayout(false);
            this.tabViolaciones.ResumeLayout(false);
            this.tabVacantes.ResumeLayout(false);
            this.pnlFiltroVacantes.ResumeLayout(false);
            this.pnlFiltroVacantes.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVacantes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvHorariosGenerados)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvViolaciones)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.TabControl tabControlPrincipal;
        private System.Windows.Forms.TabPage tabSimulacion;
        private System.Windows.Forms.SplitContainer splitContenedorSim;
        private System.Windows.Forms.GroupBox grpSeleccion;
        private System.Windows.Forms.TreeView treeContextoAcademico;
        private System.Windows.Forms.Panel pnlFiltros;
        private System.Windows.Forms.Panel pnlCabecerasNivel6;
        private System.Windows.Forms.Panel sepNivel6A;
        private System.Windows.Forms.Panel sepNivel6B;
        private System.Windows.Forms.Panel pnlArbolToolbar;
        private System.Windows.Forms.FlowLayoutPanel pnlToolbarIzq;
        private System.Windows.Forms.Panel sepToolbar1;
        private System.Windows.Forms.Button btnExpandirArbol;
        private System.Windows.Forms.Button btnContraerArbol;
        private System.Windows.Forms.Button btnGenerarAulas;
        private System.Windows.Forms.Label lblConteoCombinaciones;
        private System.Windows.Forms.Label lblPerfilArbol;
        private System.Windows.Forms.ComboBox cmbArboles;
        private System.Windows.Forms.Button btnNuevoArbol;
        private System.Windows.Forms.Button btnCargarArbol;
        private System.Windows.Forms.Button btnSobrescribirArbol;
        private System.Windows.Forms.Button btnEliminarArbol;
        private ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox cmbSedes;
        private System.Windows.Forms.Label lblSede;
        private ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox cmbProgramas;
        private System.Windows.Forms.Label lblPrograma;
        private ConversorPDF.XLL.Academic.UI.Controls.CheckedComboBox cmbPeriodos;
        private System.Windows.Forms.Label lblPeriodo;
        private System.Windows.Forms.ComboBox cmbTipoPrograma;
        private System.Windows.Forms.Label lblTipoPrograma;
        private System.Windows.Forms.Button btnFiltrarArbol;
        private System.Windows.Forms.Button btnSeleccionarTerceros;
        private System.Windows.Forms.Panel pnlConfiguracionInferior;
        private System.Windows.Forms.Panel pnlSeparadorConfig;
        private System.Windows.Forms.GroupBox grpVentana;
        private System.Windows.Forms.Label lblOrigenVentana;
        private System.Windows.Forms.CheckBox chkManualVentana;
        private System.Windows.Forms.DateTimePicker dtpVentanaFin;
        private System.Windows.Forms.Label lblFechaFin;
        private System.Windows.Forms.DateTimePicker dtpVentanaInicio;
        private System.Windows.Forms.Label lblFechaInicio;
        private System.Windows.Forms.GroupBox grpParametrosAG;
        private System.Windows.Forms.NumericUpDown numElitismo;
        private System.Windows.Forms.Label lblElitismo;
        private System.Windows.Forms.NumericUpDown numTasaMutacion;
        private System.Windows.Forms.Label lblMutacion;
        private System.Windows.Forms.NumericUpDown numTasaCruce;
        private System.Windows.Forms.Label lblCruce;
        private System.Windows.Forms.NumericUpDown numGeneraciones;
        private System.Windows.Forms.Label lblGeneraciones;
        private System.Windows.Forms.NumericUpDown numPoblacion;
        private System.Windows.Forms.Label lblPoblacion;
        private System.Windows.Forms.Panel pnlBotonesSim;
        private System.Windows.Forms.Panel pnlProgreso;
        private System.Windows.Forms.ProgressBar progressBarAG;
        private System.Windows.Forms.Label lblStatusAG;
        private System.Windows.Forms.FlowLayoutPanel flowAcciones;
        private System.Windows.Forms.CheckBox chkModoDebug;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEjecutarAG;
        private System.Windows.Forms.Button btnValidarFactibilidad;
        private System.Windows.Forms.TextBox txtLogAG;
        private System.Windows.Forms.ToolTip toolTipConteo;
        private System.Windows.Forms.TabPage tabResultados;
        private System.Windows.Forms.Panel pnlAgModo;
        private System.Windows.Forms.Button btnAgVinculaciones;
        private System.Windows.Forms.Button btnAgVerVinculaciones;
        private System.Windows.Forms.Button btnAgSesiones;
        private System.Windows.Forms.Label lblAgModo;
        private System.Windows.Forms.Panel pnlHeaderResultados;
        private System.Windows.Forms.ComboBox cmbCargarCorridas;
        private System.Windows.Forms.Label lblCargarCorridas;
        private System.Windows.Forms.Button btnPromoverProduccion;
        private System.Windows.Forms.Button btnAprobarCorrida;
        private System.Windows.Forms.Button btnRecalcularViolaciones;
        private System.Windows.Forms.Label lblHardViolaciones;
        private System.Windows.Forms.Label lblMejorFitness;
        private System.Windows.Forms.Label lblVacantes;
        private System.Windows.Forms.SplitContainer splitContainerResultados;
        private System.Windows.Forms.DataGridView dgvHorariosGenerados;
        private System.Windows.Forms.DataGridView dgvViolaciones;
        private System.Windows.Forms.TabControl tabAuditoriaResultados;
        private System.Windows.Forms.TabPage tabViolaciones;
        private System.Windows.Forms.TabPage tabVacantes;
        private System.Windows.Forms.Panel pnlFiltroVacantes;
        private System.Windows.Forms.Label lblFiltroVacante;
        private System.Windows.Forms.ComboBox cmbMotivoVacante;
        private System.Windows.Forms.Button btnAccionVacante;
        private System.Windows.Forms.Label lblConteoVacantesTab;
        private System.Windows.Forms.DataGridView dgvVacantes;
        private System.Windows.Forms.Label lblColJornada;
        private System.Windows.Forms.Label lblColDias;
        private System.Windows.Forms.Label lblColHorasDia;
    }
}
