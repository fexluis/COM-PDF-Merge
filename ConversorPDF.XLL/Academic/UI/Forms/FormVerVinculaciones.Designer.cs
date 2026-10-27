namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormVerVinculaciones
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            // C2: Liberar los objetos GDI Font (ahora son instancias independientes clonadas/creadas)
            if (disposing)
            {
                _boldFont?.Dispose();
                _regularFont?.Dispose();
                _monoFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnNivel3 = new System.Windows.Forms.Button();
            this.btnNivel2 = new System.Windows.Forms.Button();
            this.btnNivel1 = new System.Windows.Forms.Button();
            this.lblTotal = new System.Windows.Forms.Label();
            this.btnPorAsignaturas = new System.Windows.Forms.Button();
            this.btnPorSemestres = new System.Windows.Forms.Button();
            this.btnPorFechas = new System.Windows.Forms.Button();
            this.btnViaticos = new System.Windows.Forms.Button();
            this.btnVisualizar = new System.Windows.Forms.Button();
            this.btnExportar = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.treeVinculaciones = new System.Windows.Forms.TreeView();
            this.pnlBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlBottom
            // 
            this.pnlBottom.Controls.Add(this.btnNivel3);
            this.pnlBottom.Controls.Add(this.btnNivel2);
            this.pnlBottom.Controls.Add(this.btnNivel1);
            this.pnlBottom.Controls.Add(this.lblTotal);
            this.pnlBottom.Controls.Add(this.btnPorAsignaturas);
            this.pnlBottom.Controls.Add(this.btnPorSemestres);
            this.pnlBottom.Controls.Add(this.btnPorFechas);
            this.pnlBottom.Controls.Add(this.btnViaticos);
            this.pnlBottom.Controls.Add(this.btnVisualizar);
            this.pnlBottom.Controls.Add(this.btnExportar);
            this.pnlBottom.Controls.Add(this.btnCerrar);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 511);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(1000, 50);
            this.pnlBottom.TabIndex = 0;
            // 
            // btnNivel3
            // 
            this.btnNivel3.Location = new System.Drawing.Point(74, 13);
            this.btnNivel3.Name = "btnNivel3";
            this.btnNivel3.Size = new System.Drawing.Size(25, 23);
            this.btnNivel3.TabIndex = 8;
            this.btnNivel3.Text = "3";
            this.btnNivel3.UseVisualStyleBackColor = true;
            this.btnNivel3.Click += new System.EventHandler(this.btnNivel3_Click);
            // 
            // btnNivel2
            // 
            this.btnNivel2.Location = new System.Drawing.Point(43, 13);
            this.btnNivel2.Name = "btnNivel2";
            this.btnNivel2.Size = new System.Drawing.Size(25, 23);
            this.btnNivel2.TabIndex = 7;
            this.btnNivel2.Text = "2";
            this.btnNivel2.UseVisualStyleBackColor = true;
            this.btnNivel2.Click += new System.EventHandler(this.btnNivel2_Click);
            // 
            // btnNivel1
            // 
            this.btnNivel1.Location = new System.Drawing.Point(12, 13);
            this.btnNivel1.Name = "btnNivel1";
            this.btnNivel1.Size = new System.Drawing.Size(25, 23);
            this.btnNivel1.TabIndex = 6;
            this.btnNivel1.Text = "1";
            this.btnNivel1.UseVisualStyleBackColor = true;
            this.btnNivel1.Click += new System.EventHandler(this.btnNivel1_Click);
            // 
            // lblTotal
            // 
            this.lblTotal.AutoSize = true;
            this.lblTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotal.Location = new System.Drawing.Point(105, 18);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.Size = new System.Drawing.Size(110, 13);
            this.lblTotal.TabIndex = 5;
            this.lblTotal.Text = "0 sesiones totales";
            // 
            // btnPorAsignaturas
            // 
            this.btnPorAsignaturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPorAsignaturas.Location = new System.Drawing.Point(324, 13);
            this.btnPorAsignaturas.Name = "btnPorAsignaturas";
            this.btnPorAsignaturas.Size = new System.Drawing.Size(90, 23);
            this.btnPorAsignaturas.TabIndex = 4;
            this.btnPorAsignaturas.Text = "Asignaturas";
            this.btnPorAsignaturas.UseVisualStyleBackColor = true;
            this.btnPorAsignaturas.Click += new System.EventHandler(this.btnPorAsignaturas_Click);
            // 
            // btnPorSemestres
            // 
            this.btnPorSemestres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPorSemestres.Location = new System.Drawing.Point(420, 13);
            this.btnPorSemestres.Name = "btnPorSemestres";
            this.btnPorSemestres.Size = new System.Drawing.Size(90, 23);
            this.btnPorSemestres.TabIndex = 9;
            this.btnPorSemestres.Text = "Semestres";
            this.btnPorSemestres.UseVisualStyleBackColor = true;
            this.btnPorSemestres.Click += new System.EventHandler(this.btnPorSemestres_Click);
            // 
            // btnPorFechas
            // 
            this.btnPorFechas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPorFechas.Location = new System.Drawing.Point(516, 13);
            this.btnPorFechas.Name = "btnPorFechas";
            this.btnPorFechas.Size = new System.Drawing.Size(90, 23);
            this.btnPorFechas.TabIndex = 3;
            this.btnPorFechas.Text = "Por Fechas";
            this.btnPorFechas.UseVisualStyleBackColor = true;
            this.btnPorFechas.Click += new System.EventHandler(this.btnPorFechas_Click);
            // 
            // btnViaticos
            // 
            this.btnViaticos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnViaticos.Location = new System.Drawing.Point(612, 13);
            this.btnViaticos.Name = "btnViaticos";
            this.btnViaticos.Size = new System.Drawing.Size(64, 23);
            this.btnViaticos.TabIndex = 10;
            this.btnViaticos.Text = "Viaticos";
            this.btnViaticos.UseVisualStyleBackColor = true;
            this.btnViaticos.Click += new System.EventHandler(this.btnViaticos_Click);
            // 
            // btnVisualizar
            // 
            this.btnVisualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnVisualizar.Location = new System.Drawing.Point(682, 13);
            this.btnVisualizar.Name = "btnVisualizar";
            this.btnVisualizar.Size = new System.Drawing.Size(100, 23);
            this.btnVisualizar.TabIndex = 2;
            this.btnVisualizar.Text = "Timeline";
            this.btnVisualizar.UseVisualStyleBackColor = true;
            this.btnVisualizar.Click += new System.EventHandler(this.btnVisualizar_Click);
            // 
            // btnExportar
            // 
            this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportar.Location = new System.Drawing.Point(788, 13);
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.Size = new System.Drawing.Size(100, 23);
            this.btnExportar.TabIndex = 1;
            this.btnExportar.Text = "Exportar excel";
            this.btnExportar.UseVisualStyleBackColor = true;
            this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.Location = new System.Drawing.Point(894, 13);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(94, 23);
            this.btnCerrar.TabIndex = 0;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // treeVinculaciones
            // 
            this.treeVinculaciones.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeVinculaciones.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeVinculaciones.Location = new System.Drawing.Point(0, 0);
            this.treeVinculaciones.Name = "treeVinculaciones";
            this.treeVinculaciones.ShowNodeToolTips = true;
            this.treeVinculaciones.Size = new System.Drawing.Size(1000, 511);
            this.treeVinculaciones.TabIndex = 1;
            // 
            // FormVerVinculaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1000, 561);
            // M1: MinimumSize previene que los botones anclados a la derecha se superpongan
            // al redimensionar el formulario a un ancho menor que el de diseño.
            this.MinimumSize = new System.Drawing.Size(1016, 600);
            this.Controls.Add(this.treeVinculaciones);
            this.Controls.Add(this.pnlBottom);
            this.Name = "FormVerVinculaciones";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Visualización de Vinculaciones Académicas";
            this.pnlBottom.ResumeLayout(false);
            this.pnlBottom.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.Button btnVisualizar;
        private System.Windows.Forms.Button btnPorFechas;
        private System.Windows.Forms.Button btnViaticos;
        private System.Windows.Forms.Button btnPorAsignaturas;
        private System.Windows.Forms.Button btnPorSemestres;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.TreeView treeVinculaciones;
        private System.Windows.Forms.Button btnNivel1;
        private System.Windows.Forms.Button btnNivel2;
        private System.Windows.Forms.Button btnNivel3;
    }
}
