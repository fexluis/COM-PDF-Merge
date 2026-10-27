using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL;
namespace ConversorPDF.XLL.FormProgramador
{
    public partial class FormConfig : Form
    {
        private readonly IViaticosService _viaticosService;

        public FormConfig(IViaticosService viaticosService)
        {
            _viaticosService = viaticosService;
            InitializeComponent();

            this.Load += FormConfig_Load;

            // Numeric validation (Integers only)

            txtSalario.KeyPress += TextBox_KeyPress_Numeric;
            txtViaticoA.KeyPress += TextBox_KeyPress_Numeric;
            txtViaticoB.KeyPress += TextBox_KeyPress_Numeric;
            txtCedulaGerente.KeyPress += TextBox_KeyPress_Numeric;

            // Decimal validation (Regional)
            txtPorcA.KeyPress += TextBox_KeyPress_Decimal;
            txtPorcB.KeyPress += TextBox_KeyPress_Decimal;
        }

        private void TextBox_KeyPress_Numeric(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TextBox_KeyPress_Decimal(object sender, KeyPressEventArgs e)
        {
            char decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];


            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != decimalSeparator)
            {
                e.Handled = true;
            }

            // Only one decimal separator
            if (e.KeyChar == decimalSeparator && (sender as TextBox).Text.IndexOf(decimalSeparator) > -1)
            {
                e.Handled = true;
            }
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            try
            {
                object[,] data = _viaticosService.GetAdvancedConfigRows();
                if (data != null && data.GetLength(0) >= 13)
                {
                    txtSalario.Text       = data[0, 0]?.ToString() ?? "";
                    txtViaticoA.Text      = data[1, 0]?.ToString() ?? "";
                    txtViaticoB.Text      = data[2, 0]?.ToString() ?? "";
                    txtPorcA.Text         = data[3, 0]?.ToString() ?? "";
                    txtPorcB.Text         = data[4, 0]?.ToString() ?? "";
                    txtNombreDT.Text      = data[5, 0]?.ToString() ?? "";
                    txtCargoDT.Text       = data[6, 0]?.ToString() ?? "";
                    txtCedulaGerente.Text = data[7, 0]?.ToString() ?? "";
                    txtGerenteSustituto.Text = data[8, 0]?.ToString() ?? "";
                    txtCargoSustituto.Text   = data[9, 0]?.ToString() ?? "";
                    txtDeptoBase.Text     = data[10, 0]?.ToString() ?? "";
                    txtCiudadBase.Text    = data[11, 0]?.ToString() ?? "";
                    txtAeropuertoBase.Text = data[12, 0]?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error cargando configuración en FormConfig", ex);
            }
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                object[,] data = new object[13, 1];
                data[0, 0] = txtSalario.Text;
                data[1, 0] = txtViaticoA.Text;
                data[2, 0] = txtViaticoB.Text;
                data[3, 0] = txtPorcA.Text;
                data[4, 0] = txtPorcB.Text;
                data[5, 0] = txtNombreDT.Text;
                data[6, 0] = txtCargoDT.Text;
                data[7, 0] = txtCedulaGerente.Text;
                data[8, 0] = txtGerenteSustituto.Text;
                data[9, 0] = txtCargoSustituto.Text;
                data[10, 0] = txtDeptoBase.Text;
                data[11, 0] = txtCiudadBase.Text;
                data[12, 0] = txtAeropuertoBase.Text;

                _viaticosService.SaveAdvancedConfigRows(data);
                this.Close();

            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error guardando configuración en FormConfig", ex);
                MessageBox.Show("Ocurrió un error al guardar la configuración.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}