using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Services;
namespace ConversorPDF.XLL.FormProgramador
{
    public partial class FormGestionDocumental : Form
    {
        private dynamic _activeWorkbook;
        private readonly IViaticosService _viaticosService;

        public FormGestionDocumental(dynamic activeWorkbook, IViaticosService viaticosService)
        {
            InitializeComponent();
            _activeWorkbook = activeWorkbook;
            _viaticosService = viaticosService;

            // Numeric validation
            txtViaticoInicial.KeyPress += TxtNumeric_KeyPress;
            txtViaticoFinal.KeyPress += TxtNumeric_KeyPress;

            btnRenombrar.Click += BtnRenombrar_Click;
            btnUnificar.Click += BtnUnificar_Click;
        }

        private void TxtNumeric_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void BtnRenombrar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtViaticoInicial.Text) || string.IsNullOrWhiteSpace(txtViaticoFinal.Text))
            {
                MessageBox.Show("Por favor, ingrese el rango de viáticos (Inicial y Final).", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int inicio = int.Parse(txtViaticoInicial.Text);
            int fin = int.Parse(txtViaticoFinal.Text);

            _viaticosService.RenombrarSolicitudes(inicio, fin);
        }

        private void BtnUnificar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtViaticoInicial.Text) || string.IsNullOrWhiteSpace(txtViaticoFinal.Text))
            {
                MessageBox.Show("Por favor, ingrese el rango de viáticos (Inicial y Final).", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int inicio = int.Parse(txtViaticoInicial.Text);
            int fin = int.Parse(txtViaticoFinal.Text);

            _viaticosService.UnificarSolicitudes(inicio, fin);
        }
    }
}