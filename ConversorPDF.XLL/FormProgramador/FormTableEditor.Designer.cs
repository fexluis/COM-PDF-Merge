using System.Drawing;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Windows.Forms;
namespace ConversorPDF.XLL.FormProgramador
{
    partial class FormTableEditor
    {
        private System.ComponentModel.IContainer components = null;
        private DataGridView dgvData;
        private Button btnGuardar;
        private Button btnCancelar;
        private Label lblInstrucciones;

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
            this.dgvData = new DataGridView();
            this.btnGuardar = new Button();
            this.btnCancelar = new Button();
            this.lblInstrucciones = new Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.SuspendLayout();

            // dgvData
            this.dgvData.AllowUserToOrderColumns = true;
            this.dgvData.Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom) | AnchorStyles.Left) | AnchorStyles.Right)));
            this.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Location = new Point(12, 40);
            this.dgvData.Name = "dgvData";
            this.dgvData.Size = new Size(560, 300);
            this.dgvData.TabIndex = 0;

            // btnGuardar
            this.btnGuardar.Anchor = ((AnchorStyles)((AnchorStyles.Bottom | AnchorStyles.Right)));
            this.btnGuardar.Location = new Point(416, 350);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new Size(75, 23);
            this.btnGuardar.TabIndex = 1;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = true;

            // btnCancelar
            this.btnCancelar.Anchor = ((AnchorStyles)((AnchorStyles.Bottom | AnchorStyles.Right)));
            this.btnCancelar.Location = new Point(497, 350);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new Size(75, 23);
            this.btnCancelar.TabIndex = 2;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;

            // lblInstrucciones
            this.lblInstrucciones.AutoSize = true;
            this.lblInstrucciones.Location = new Point(12, 15);
            this.lblInstrucciones.Name = "lblInstrucciones";
            this.lblInstrucciones.Size = new Size(350, 13);
            this.lblInstrucciones.TabIndex = 3;
            this.lblInstrucciones.Text = "Edite los datos directamente en la tabla. Use Supr para eliminar filas.";

            // FormTableEditor
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(584, 385);
            this.Controls.Add(this.lblInstrucciones);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.dgvData);
            this.MinimizeBox = false;
            this.Name = "FormTableEditor";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Editor de Tabla";
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}