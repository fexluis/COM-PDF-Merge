namespace ConversorPDF.XLL.Academic.UI.Forms
{
    partial class FormUsuarios
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                picFirma?.Image?.Dispose();
                picFirma.Image = null;
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.dgvUsuarios = new System.Windows.Forms.DataGridView();
            this.txtNombreUsuario = new System.Windows.Forms.TextBox();
            this.lblNombreUsuario = new System.Windows.Forms.Label();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.lblNombre = new System.Windows.Forms.Label();
            this.cmbPerfil = new System.Windows.Forms.ComboBox();
            this.lblPerfil = new System.Windows.Forms.Label();
            this.cmbSubunidad = new System.Windows.Forms.ComboBox();
            this.lblSubunidad = new System.Windows.Forms.Label();
            this.cmbActivo = new System.Windows.Forms.ComboBox();
            this.lblActivo = new System.Windows.Forms.Label();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.txtFirma = new System.Windows.Forms.TextBox();
            this.lblFirma = new System.Windows.Forms.Label();
            this.btnSeleccionarFirma = new System.Windows.Forms.Button();
            this.btnLimpiarFirma = new System.Windows.Forms.Button();
            this.picFirma = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsuarios)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picFirma)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvUsuarios
            // 
            this.dgvUsuarios.AllowUserToAddRows = false;
            this.dgvUsuarios.AllowUserToDeleteRows = false;
            this.dgvUsuarios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUsuarios.Location = new System.Drawing.Point(12, 12);
            this.dgvUsuarios.MultiSelect = false;
            this.dgvUsuarios.Name = "dgvUsuarios";
            this.dgvUsuarios.ReadOnly = true;
            this.dgvUsuarios.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvUsuarios.Size = new System.Drawing.Size(534, 187);
            this.dgvUsuarios.TabIndex = 0;
            // 
            // txtNombreUsuario
            // 
            this.txtNombreUsuario.Location = new System.Drawing.Point(111, 219);
            this.txtNombreUsuario.Name = "txtNombreUsuario";
            this.txtNombreUsuario.Size = new System.Drawing.Size(183, 20);
            this.txtNombreUsuario.TabIndex = 2;
            // 
            // lblNombreUsuario
            // 
            this.lblNombreUsuario.AutoSize = true;
            this.lblNombreUsuario.Location = new System.Drawing.Point(12, 222);
            this.lblNombreUsuario.Name = "lblNombreUsuario";
            this.lblNombreUsuario.Size = new System.Drawing.Size(86, 13);
            this.lblNombreUsuario.TabIndex = 1;
            this.lblNombreUsuario.Text = "Nombre Usuario:";
            // 
            // txtNombre
            // 
            this.txtNombre.Location = new System.Drawing.Point(111, 245);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(183, 20);
            this.txtNombre.TabIndex = 4;
            // 
            // lblNombre
            // 
            this.lblNombre.AutoSize = true;
            this.lblNombre.Location = new System.Drawing.Point(12, 248);
            this.lblNombre.Name = "lblNombre";
            this.lblNombre.Size = new System.Drawing.Size(47, 13);
            this.lblNombre.TabIndex = 3;
            this.lblNombre.Text = "Nombre:";
            // 
            // cmbPerfil
            // 
            this.cmbPerfil.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPerfil.FormattingEnabled = true;
            this.cmbPerfil.Location = new System.Drawing.Point(111, 271);
            this.cmbPerfil.Name = "cmbPerfil";
            this.cmbPerfil.Size = new System.Drawing.Size(183, 21);
            this.cmbPerfil.TabIndex = 6;
            // 
            // lblPerfil
            // 
            this.lblPerfil.AutoSize = true;
            this.lblPerfil.Location = new System.Drawing.Point(12, 274);
            this.lblPerfil.Name = "lblPerfil";
            this.lblPerfil.Size = new System.Drawing.Size(33, 13);
            this.lblPerfil.TabIndex = 5;
            this.lblPerfil.Text = "Perfil:";
            // 
            // 
            // cmbSubunidad
            // 
            this.cmbSubunidad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubunidad.FormattingEnabled = true;
            this.cmbSubunidad.Location = new System.Drawing.Point(111, 298);
            this.cmbSubunidad.Name = "cmbSubunidad";
            this.cmbSubunidad.Size = new System.Drawing.Size(183, 21);
            this.cmbSubunidad.TabIndex = 7;
            // 
            // lblSubunidad
            // 
            this.lblSubunidad.AutoSize = true;
            this.lblSubunidad.Location = new System.Drawing.Point(12, 301);
            this.lblSubunidad.Name = "lblSubunidad";
            this.lblSubunidad.Size = new System.Drawing.Size(61, 13);
            this.lblSubunidad.TabIndex = 8;
            this.lblSubunidad.Text = "Subunidad:";
            // 
            // cmbActivo
            // 
            this.cmbActivo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbActivo.FormattingEnabled = true;
            this.cmbActivo.Location = new System.Drawing.Point(111, 325);
            this.cmbActivo.Name = "cmbActivo";
            this.cmbActivo.Size = new System.Drawing.Size(183, 21);
            this.cmbActivo.TabIndex = 9;
            // 
            // lblActivo
            // 
            this.lblActivo.AutoSize = true;
            this.lblActivo.Location = new System.Drawing.Point(12, 328);
            this.lblActivo.Name = "lblActivo";
            this.lblActivo.Size = new System.Drawing.Size(40, 13);
            this.lblActivo.TabIndex = 10;
            this.lblActivo.Text = "Activo:";
            // 
            // btnNuevo
            // 
            this.btnNuevo.Location = new System.Drawing.Point(185, 360);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.Size = new System.Drawing.Size(75, 23);
            this.btnNuevo.TabIndex = 11;
            this.btnNuevo.Text = "Nuevo";
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            // 
            // txtFirma
            // 
            this.txtFirma.Location = new System.Drawing.Point(12, 415);
            this.txtFirma.MaxLength = 0;
            this.txtFirma.Multiline = true;
            this.txtFirma.ReadOnly = true;
            this.txtFirma.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtFirma.Name = "txtFirma";
            this.txtFirma.Size = new System.Drawing.Size(534, 60);
            this.txtFirma.TabIndex = 15;
            // 
            // lblFirma
            // 
            this.lblFirma.AutoSize = true;
            this.lblFirma.Location = new System.Drawing.Point(12, 395);
            this.lblFirma.Name = "lblFirma";
            this.lblFirma.Size = new System.Drawing.Size(35, 13);
            this.lblFirma.TabIndex = 16;
            this.lblFirma.Text = "Firma:";
            // 
            // btnGuardar
            // 
            this.btnGuardar.Location = new System.Drawing.Point(355, 360);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 12;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Location = new System.Drawing.Point(270, 360);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(75, 23);
            this.btnEliminar.TabIndex = 13;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(440, 360);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 14;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnSeleccionarFirma
            // 
            this.btnSeleccionarFirma.Location = new System.Drawing.Point(12, 360);
            this.btnSeleccionarFirma.Name = "btnSeleccionarFirma";
            this.btnSeleccionarFirma.Size = new System.Drawing.Size(105, 23);
            this.btnSeleccionarFirma.TabIndex = 17;
            this.btnSeleccionarFirma.Text = "Seleccionar Firma";
            this.btnSeleccionarFirma.UseVisualStyleBackColor = true;
            this.btnSeleccionarFirma.Click += new System.EventHandler(this.btnSeleccionarFirma_Click);
            // 
            // btnLimpiarFirma
            // 
            this.btnLimpiarFirma.Location = new System.Drawing.Point(120, 360);
            this.btnLimpiarFirma.Name = "btnLimpiarFirma";
            this.btnLimpiarFirma.Size = new System.Drawing.Size(60, 23);
            this.btnLimpiarFirma.TabIndex = 19;
            this.btnLimpiarFirma.Text = "Limpiar";
            this.btnLimpiarFirma.UseVisualStyleBackColor = true;
            this.btnLimpiarFirma.Click += new System.EventHandler(this.btnLimpiarFirma_Click);
            // 
            // picFirma
            // 
            this.picFirma.BackColor = System.Drawing.Color.LightYellow;
            this.picFirma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picFirma.Location = new System.Drawing.Point(320, 219);
            this.picFirma.Name = "picFirma";
            this.picFirma.Size = new System.Drawing.Size(226, 127);
            this.picFirma.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picFirma.TabIndex = 18;
            this.picFirma.TabStop = false;
            // 
            // FormUsuarios
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 520);
            this.Controls.Add(this.picFirma);
            this.Controls.Add(this.btnLimpiarFirma);
            this.Controls.Add(this.btnSeleccionarFirma);
            this.Controls.Add(this.lblFirma);
            this.Controls.Add(this.txtFirma);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.btnNuevo);
            this.Controls.Add(this.lblActivo);
            this.Controls.Add(this.cmbActivo);
            this.Controls.Add(this.lblSubunidad);
            this.Controls.Add(this.cmbSubunidad);
            this.Controls.Add(this.lblPerfil);
            this.Controls.Add(this.cmbPerfil);
            this.Controls.Add(this.lblNombre);
            this.Controls.Add(this.txtNombre);
            this.Controls.Add(this.lblNombreUsuario);
            this.Controls.Add(this.txtNombreUsuario);
            this.Controls.Add(this.dgvUsuarios);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormUsuarios";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Gestión de Usuarios";
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsuarios)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picFirma)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.DataGridView dgvUsuarios;
        private System.Windows.Forms.TextBox txtNombreUsuario;
        private System.Windows.Forms.Label lblNombreUsuario;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label lblNombre;
        private System.Windows.Forms.ComboBox cmbPerfil;
        private System.Windows.Forms.Label lblPerfil;
        private System.Windows.Forms.ComboBox cmbSubunidad;
        private System.Windows.Forms.Label lblSubunidad;
        private System.Windows.Forms.ComboBox cmbActivo;
        private System.Windows.Forms.Label lblActivo;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.TextBox txtFirma;
        private System.Windows.Forms.Label lblFirma;
        private System.Windows.Forms.PictureBox picFirma;
        private System.Windows.Forms.Button btnSeleccionarFirma;
        private System.Windows.Forms.Button btnLimpiarFirma;
    }
}
