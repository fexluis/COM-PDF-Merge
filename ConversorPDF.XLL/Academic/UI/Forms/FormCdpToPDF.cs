using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormCdpToPDF : Form
    {
        private readonly ConfigRepository _configRepo;
        private readonly CdpDependenciaRepository _cdpRepo;
        private Config _currentConfig;
        
        public FormCdpToPDF()
        {
            InitializeComponent();
            var context = SqliteHelper.GetContext();
            _configRepo = new ConfigRepository(context);
            _cdpRepo = new CdpDependenciaRepository(context.Open());
        }

        private void FormCdpToPDF_Load(object sender, EventArgs e)
        {
            try
            {
                picLogo.Image = GetImageResource("ConversorPDF.XLL.img.esap.png");

                var configResult = _configRepo.GetConfig();
                if (configResult.IsFailure)
                {
                    MessageBox.Show(configResult.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    return;
                }
                
                _currentConfig = configResult.Value;

                var cdpResult = _cdpRepo.GetAll();
                if (cdpResult.IsFailure)
                {
                    MessageBox.Show(cdpResult.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    return;
                }

                var activos = cdpResult.Value.Where(c => c.Activo == 1).ToList();
                
                var dataSource = activos.Select(c => new 
                { 
                    Text = string.IsNullOrEmpty(c.Dependencia) ? c.CdpNum : $"{c.CdpNum}-{c.Dependencia}", 
                    Value = c.CdpNum 
                }).ToList();

                cmbCdp.DisplayMember = "Text";
                cmbCdp.ValueMember = "Value";
                cmbCdp.DataSource = dataSource;

                if (!string.IsNullOrEmpty(_currentConfig.Cdp))
                {
                    cmbCdp.SelectedValue = _currentConfig.Cdp;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormCdpToPDF] Error al cargar CDPs", ex);
                MessageBox.Show("Ocurrió un error al cargar los datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool GuardarConfiguracion()
        {
            if (cmbCdp.SelectedValue == null)
            {
                MessageBox.Show("Por favor seleccione un CDP.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            _currentConfig.Cdp = cmbCdp.SelectedValue.ToString();
            
            var result = _configRepo.UpdateConfig(_currentConfig);
            if (result.IsFailure)
            {
                MessageBox.Show(result.Error, "Error al guardar configuración", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
            return true;
        }

        private void btnPdfViaticos_Click(object sender, EventArgs e)
        {
            if (GuardarConfiguracion())
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void btnMasivo_Click(object sender, EventArgs e)
        {
            if (GuardarConfiguracion())
            {
                this.DialogResult = DialogResult.Yes;
                this.Close();
            }
        }

        private Image GetImageResource(string name)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(name))
                {
                    if (stream == null) return null;
                    
                    using (var originalImage = Image.FromStream(stream))
                    {
                        return new Bitmap(originalImage);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"No se pudo cargar el recurso de imagen {name}: {ex.Message}");
                return null;
            }
        }
    }
}
