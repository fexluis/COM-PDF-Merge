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
    partial class FormProgramador
    {
        private System.ComponentModel.IContainer components = null;

        // Left pane controls
        private Label lblFechaInicio;
        private TextBox txtFechaInicio;
        private Button btnCalInicio;

        private Label lblFechaFinal;
        private TextBox txtFechaFinal;
        private Button btnCalFinal;

        private Label lblCambiarCDP;
        private ComboBox cmbCDP;
        
        private Label lblCDPHeader;
        private Label lblCDPValue;
        private Label lblCDPDependencia;
        private Label lblCDPDependenciaDescripcion;

        private Button btnImprimirPDF;
        private Button btnImprimirMasivo;
        private Button btnImprimirCumplido;
        private Button btnEditarDependencias;
        private Button btnEditarConfig;
        private Button btnEditarDestinos;
        private Button btnEditarCategoria1;

        // Right pane controls (Calendar)
        private Panel pnlCalendar;
        private Label lblMonthYear;
        private Button btnPrevMonth;
        private Button btnNextMonth;
        private Panel pnlDaysHeader;
        private Panel pnlDaysGrid;

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
            this.SuspendLayout();

            int leftColX = 15;
            int rightColX = 260; // Calendar X

            // lblFechaInicio
            this.lblFechaInicio = new Label { Text = "Fecha de Inicio:", Location = new Point(leftColX, 15), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            // txtFechaInicio
            this.txtFechaInicio = new TextBox { Location = new Point(leftColX, 35), Size = new Size(180, 25), Font = new Font("Segoe UI", 12F) };
            // btnCalInicio
            this.btnCalInicio = new Button { Text = "📅", Location = new Point(200, 34), Size = new Size(35, 27), Font = new Font("Segoe UI", 10F), FlatStyle = FlatStyle.System };

            // lblFechaFinal
            this.lblFechaFinal = new Label { Text = "Fecha Final:", Location = new Point(leftColX, 75), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            // txtFechaFinal
            this.txtFechaFinal = new TextBox { Location = new Point(leftColX, 95), Size = new Size(180, 25), Font = new Font("Segoe UI", 12F) };
            // btnCalFinal
            this.btnCalFinal = new Button { Text = "📅", Location = new Point(200, 94), Size = new Size(35, 27), Font = new Font("Segoe UI", 10F), FlatStyle = FlatStyle.System };

            // lblCambiarCDP
            this.lblCambiarCDP = new Label { Text = "Cambiar CDP:", Location = new Point(leftColX, 140), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            // cmbCDP
            this.cmbCDP = new ComboBox { Location = new Point(leftColX, 160), Size = new Size(100, 25), Font = new Font("Segoe UI", 10F), DropDownStyle = ComboBoxStyle.DropDownList };
            
            // lblCDPHeader
            this.lblCDPHeader = new Label { Text = "CDP:", Location = new Point(130, 140), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            // lblCDPValue
            this.lblCDPValue = new Label { Text = "----", Location = new Point(130, 155), AutoSize = true, Font = new Font("Segoe UI", 24F, FontStyle.Regular), ForeColor = Color.Black };

            // lblCDPDependencia
            this.lblCDPDependencia = new Label { Text = "DEPENDENCIA", Location = new Point(leftColX, 195), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Black };

            // lblCDPDependenciaDescripcion
            this.lblCDPDependenciaDescripcion = new Label { Text = "DESCRIPCION", Location = new Point(leftColX, 215), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.DimGray };

            // btnImprimirPDF
            this.btnImprimirPDF = new Button { Text = "IMPRIMIR PDF VIATICOS", Location = new Point(leftColX, 260), Size = new Size(220, 35), Font = new Font("Segoe UI", 9F), FlatStyle = FlatStyle.Flat, BackColor = Color.LightBlue };
            
            // btnImprimirMasivo
            this.btnImprimirMasivo = new Button { Text = "GENERAR ARCHIVO MASIVO", Location = new Point(leftColX, 260), Size = new Size(220, 35), Font = new Font("Segoe UI", 9F), FlatStyle = FlatStyle.Flat, BackColor = Color.LightBlue, Visible = false };
            
            // btnImprimirCumplido
            this.btnImprimirCumplido = new Button { Text = "GENERAR CUMPLIDOS E INFORMES", Location = new Point(leftColX, 260), Size = new Size(220, 35), Font = new Font("Segoe UI", 9F), FlatStyle = FlatStyle.Flat, BackColor = Color.LightBlue, Visible = false };
            
            // btnEditarDependencias
            this.btnEditarDependencias = new Button { Text = "EDITAR DEPENDENCIAS", Location = new Point(15, 330), Size = new Size(235, 35), Font = new Font("Segoe UI", 9F), FlatStyle = FlatStyle.System };

            // btnEditarConfig
            this.btnEditarConfig = new Button { Text = "CONFIGURACION", Location = new Point(rightColX, 330), Size = new Size(110, 35), Font = new Font("Segoe UI", 8F), FlatStyle = FlatStyle.System };
            
            // btnEditarDestinos
            this.btnEditarDestinos = new Button { Text = "DESTINOS", Location = new Point(rightColX + 115, 330), Size = new Size(110, 35), Font = new Font("Segoe UI", 8F), FlatStyle = FlatStyle.System };

            // btnEditarCategoria1
            this.btnEditarCategoria1 = new Button { Text = "M. CATEGORIA UNO", Location = new Point(rightColX + 230, 330), Size = new Size(115, 35), Font = new Font("Segoe UI", 8F), FlatStyle = FlatStyle.System };

            // Calendar Panel
            this.pnlCalendar = new Panel { Location = new Point(rightColX, 15), Size = new Size(345, 280), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            
            // Calendar Header Background
            Panel pnlCalHeaderBg = new Panel { Location = new Point(0, 0), Size = new Size(345, 60), BackColor = Color.DarkBlue };
            this.pnlCalendar.Controls.Add(pnlCalHeaderBg);

            // Month Year Label
            this.lblMonthYear = new Label { Text = "MES AÑO", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 16F, FontStyle.Regular), ForeColor = Color.White, BackColor = Color.Transparent };
            pnlCalHeaderBg.Controls.Add(this.lblMonthYear);

            // Prev Next Buttons
            this.btnPrevMonth = new Button { Text = "◀", Location = new Point(275, 15), Size = new Size(25, 25), FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.DarkBlue, Cursor = Cursors.Hand };
            this.btnPrevMonth.FlatAppearance.BorderSize = 0;
            this.btnNextMonth = new Button { Text = "▶", Location = new Point(305, 15), Size = new Size(25, 25), FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.DarkBlue, Cursor = Cursors.Hand };
            this.btnNextMonth.FlatAppearance.BorderSize = 0;
            pnlCalHeaderBg.Controls.Add(this.btnPrevMonth);
            pnlCalHeaderBg.Controls.Add(this.btnNextMonth);

            // Days Header
            this.pnlDaysHeader = new Panel { Location = new Point(0, 60), Size = new Size(345, 25), BackColor = Color.DarkBlue };
            string[] days = { "DOM", "LUN", "MAR", "MIE", "JUE", "VIE", "SAB" };
            for (int i = 0; i < 7; i++)
            {
                Label lblDay = new Label { Text = days[i], Location = new Point(i * 49, 5), Size = new Size(49, 15), Font = new Font("Segoe UI", 8F), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
                this.pnlDaysHeader.Controls.Add(lblDay);
            }
            this.pnlCalendar.Controls.Add(this.pnlDaysHeader);

            // Days Grid
            this.pnlDaysGrid = new Panel { Location = new Point(0, 85), Size = new Size(345, 190), BackColor = Color.White };
            this.pnlCalendar.Controls.Add(this.pnlDaysGrid);

            // Main Form Settings
            this.ClientSize = new Size(625, 385);
            this.Controls.Add(this.lblFechaInicio);
            this.Controls.Add(this.txtFechaInicio);
            this.Controls.Add(this.btnCalInicio);
            this.Controls.Add(this.lblFechaFinal);
            this.Controls.Add(this.txtFechaFinal);
            this.Controls.Add(this.btnCalFinal);
            this.Controls.Add(this.lblCambiarCDP);
            this.Controls.Add(this.cmbCDP);
            this.Controls.Add(this.lblCDPHeader);
            this.Controls.Add(this.lblCDPValue);
            this.Controls.Add(this.lblCDPDependencia);
            this.Controls.Add(this.lblCDPDependenciaDescripcion);
            this.Controls.Add(this.btnImprimirPDF);
            this.Controls.Add(this.btnImprimirMasivo);
            this.Controls.Add(this.btnImprimirCumplido);
            this.Controls.Add(this.btnEditarDependencias);
            this.Controls.Add(this.btnEditarConfig);
            this.Controls.Add(this.btnEditarDestinos);
            this.Controls.Add(this.btnEditarCategoria1);
            this.Controls.Add(this.pnlCalendar);

            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormProgramador";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "PROGRAMACION DE VIATICOS";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}