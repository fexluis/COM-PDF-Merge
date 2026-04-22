using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Drawing;
using System.Windows.Forms;
namespace ConversorPDF.XLL.FormProgramador
{
    partial class FormConfig
    {
        private System.ComponentModel.IContainer components = null;

        private Label lblSalario; private TextBox txtSalario;
        private Label lblViaticoA; private TextBox txtViaticoA;
        private Label lblViaticoB; private TextBox txtViaticoB;
        private Label lblPorcA; private TextBox txtPorcA;
        private Label lblPorcB; private TextBox txtPorcB;
        private Label lblDeptoBase; private TextBox txtDeptoBase;
        private Label lblNombreDT; private TextBox txtNombreDT;
        private Label lblCargoDT; private TextBox txtCargoDT;
        private Label lblCedulaGerente; private TextBox txtCedulaGerente;
        private Label lblGerenteSustituto; private TextBox txtGerenteSustituto;
        private Label lblCargoSustituto; private TextBox txtCargoSustituto;
        private Label lblCiudadBase; private TextBox txtCiudadBase;
        private Label lblAeropuertoBase; private TextBox txtAeropuertoBase;

        private Button btnGuardar;
        private Button btnCancelar;

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
            this.SuspendLayout();

            int lblX = 15;
            int txtX = 220;
            int startY = 15;
            int spacing = 32;

            int currentY = startY;

            // Labels and TextBoxes
            lblSalario = CreateLabel("Salario Decreto 613 de 2025:", lblX, currentY);
            txtSalario = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblViaticoA = CreateLabel("Valor Viático Diario Tarifa A:", lblX, currentY);
            txtViaticoA = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblViaticoB = CreateLabel("Valor Viático Diario Tarifa B:", lblX, currentY);
            txtViaticoB = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblPorcA = CreateLabel("Porcentaje Sobre Decreto A:", lblX, currentY);
            txtPorcA = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblPorcB = CreateLabel("Porcentaje Sobre Decreto B:", lblX, currentY);
            txtPorcB = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblDeptoBase = CreateLabel("Departamento Base:", lblX, currentY);
            txtDeptoBase = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblNombreDT = CreateLabel("Nombre DT:", lblX, currentY);
            txtNombreDT = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblCargoDT = CreateLabel("Cargo DT:", lblX, currentY);
            txtCargoDT = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblCedulaGerente = CreateLabel("Cédula Gerente:", lblX, currentY);
            txtCedulaGerente = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblGerenteSustituto = CreateLabel("Subdirector Nal. Corporativo:", lblX, currentY);
            txtGerenteSustituto = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblCargoSustituto = CreateLabel("Cargo Subdirector:", lblX, currentY);
            txtCargoSustituto = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblCiudadBase = CreateLabel("Ciudad Base:", lblX, currentY);
            txtCiudadBase = CreateTextBox(txtX, currentY);
            currentY += spacing;

            lblAeropuertoBase = CreateLabel("Aeropuerto Base:", lblX, currentY);
            txtAeropuertoBase = CreateTextBox(txtX, currentY);
            currentY += (spacing + 10);

            // Buttons
            btnGuardar = new Button
            {
                Text = "GUARDAR",
                Location = new Point(120, currentY),
                Size = new Size(130, 35),
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            btnGuardar.Click += BtnGuardar_Click;

            btnCancelar = new Button
            {
                Text = "CANCELAR",
                Location = new Point(260, currentY),
                Size = new Size(130, 35),
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI", 9F)
            };
            btnCancelar.Click += BtnCancelar_Click;

            // Form settings
            this.ClientSize = new Size(520, currentY + 60);
            this.Controls.Add(lblSalario); this.Controls.Add(txtSalario);
            this.Controls.Add(lblViaticoA); this.Controls.Add(txtViaticoA);
            this.Controls.Add(lblViaticoB); this.Controls.Add(txtViaticoB);
            this.Controls.Add(lblPorcA); this.Controls.Add(txtPorcA);
            this.Controls.Add(lblPorcB); this.Controls.Add(txtPorcB);
            this.Controls.Add(lblDeptoBase); this.Controls.Add(txtDeptoBase);
            this.Controls.Add(lblNombreDT); this.Controls.Add(txtNombreDT);
            this.Controls.Add(lblCargoDT); this.Controls.Add(txtCargoDT);
            this.Controls.Add(lblCedulaGerente); this.Controls.Add(txtCedulaGerente);
            this.Controls.Add(lblGerenteSustituto); this.Controls.Add(txtGerenteSustituto);
            this.Controls.Add(lblCargoSustituto); this.Controls.Add(txtCargoSustituto);
            this.Controls.Add(lblCiudadBase); this.Controls.Add(txtCiudadBase);
            this.Controls.Add(lblAeropuertoBase); this.Controls.Add(txtAeropuertoBase);
            this.Controls.Add(btnGuardar);
            this.Controls.Add(btnCancelar);

            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormConfig";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "CONFIGURACION DE VIATICOS";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private Label CreateLabel(string text, int x, int y)
        {
            return new Label { Text = text, Location = new Point(x, y), AutoSize = true, Font = new Font("Segoe UI", 9F) };
        }

        private TextBox CreateTextBox(int x, int y)
        {
            return new TextBox { Location = new Point(x, y - 3), Size = new Size(270, 23), Font = new Font("Segoe UI", 9F) };
        }
    }
}