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
    partial class FormGestionDocumental
    {
        private System.ComponentModel.IContainer components = null;

        private PictureBox picHero;
        private Label lblViaticoInicial;
        private TextBox txtViaticoInicial;
        private Label lblViaticoFinal;
        private TextBox txtViaticoFinal;
        private Button btnRenombrar;
        private Button btnUnificar;

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
            this.picHero = new PictureBox();
            this.lblViaticoInicial = new Label();
            this.txtViaticoInicial = new TextBox();
            this.lblViaticoFinal = new Label();
            this.txtViaticoFinal = new TextBox();
            this.btnRenombrar = new Button();
            this.btnUnificar = new Button();
            this.SuspendLayout();

            // picHero
            this.picHero.Image = App.GetResourceImage("ConversorPDF.XLL.img.hero.png");
            this.picHero.Location = new Point(0, 0);
            this.picHero.Size = new Size(360, 180);
            this.picHero.SizeMode = PictureBoxSizeMode.StretchImage;

            // lblViaticoInicial
            this.lblViaticoInicial.AutoSize = true;
            this.lblViaticoInicial.Location = new Point(30, 205);
            this.lblViaticoInicial.Text = "Viático Inicial:";
            this.lblViaticoInicial.Font = new Font("Segoe UI", 10F, FontStyle.Regular);

            // txtViaticoInicial
            this.txtViaticoInicial.Location = new Point(30, 230);
            this.txtViaticoInicial.Size = new Size(140, 25);
            this.txtViaticoInicial.Font = new Font("Segoe UI", 11F);
            this.txtViaticoInicial.TextAlign = HorizontalAlignment.Center;

            // lblViaticoFinal
            this.lblViaticoFinal.AutoSize = true;
            this.lblViaticoFinal.Location = new Point(190, 205);
            this.lblViaticoFinal.Text = "Viático Final:";
            this.lblViaticoFinal.Font = new Font("Segoe UI", 10F, FontStyle.Regular);

            // txtViaticoFinal
            this.txtViaticoFinal.Location = new Point(190, 230);
            this.txtViaticoFinal.Size = new Size(140, 25);
            this.txtViaticoFinal.Font = new Font("Segoe UI", 11F);
            this.txtViaticoFinal.TextAlign = HorizontalAlignment.Center;

            // btnRenombrar
            this.btnRenombrar.Location = new Point(30, 285);
            this.btnRenombrar.Size = new Size(300, 45);
            this.btnRenombrar.Text = "RENOMBRAR SOLICITUDES";
            this.btnRenombrar.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.btnRenombrar.FlatStyle = FlatStyle.Flat;
            this.btnRenombrar.BackColor = Color.White;
            this.btnRenombrar.FlatAppearance.BorderColor = Color.LightGray;
            this.btnRenombrar.Cursor = Cursors.Hand;

            // btnUnificar
            this.btnUnificar.Location = new Point(30, 345);
            this.btnUnificar.Size = new Size(300, 45);
            this.btnUnificar.Text = "UNIFICAR SOLICITUDES";
            this.btnUnificar.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.btnUnificar.FlatStyle = FlatStyle.Flat;
            this.btnUnificar.BackColor = Color.White;
            this.btnUnificar.FlatAppearance.BorderColor = Color.LightGray;
            this.btnUnificar.Cursor = Cursors.Hand;

            // FormGestionDocumental
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(360, 420);
            this.Controls.Add(this.picHero);
            this.Controls.Add(this.lblViaticoInicial);
            this.Controls.Add(this.txtViaticoInicial);
            this.Controls.Add(this.lblViaticoFinal);
            this.Controls.Add(this.txtViaticoFinal);
            this.Controls.Add(this.btnRenombrar);
            this.Controls.Add(this.btnUnificar);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Gestión Documental Viáticos";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}