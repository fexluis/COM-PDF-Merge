using System;
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
    public partial class FormInfo : Form
    {
        public FormInfo()
        {
            ConversorPDF.XLL.Core.Helpers.DatabaseBootstrapper.EnsureInitialized();
            InitializeComponent();
        }

        private async void btnSync_Click(object sender, EventArgs e)
        {
            try
            {
                // Fase 64a FIX: GetContext() retorna singleton — usar using lo dispone para todos.
                // GetContextForPath() retorna instancia nueva, segura para using.
                using (var ctx = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContextForPath())
                {
                    string sub = Dapper.SqlMapper.QueryFirstOrDefault<string>(ctx.Open(), "SELECT subunidad FROM config LIMIT 1");
                    ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info($"[FormInfo] btnSync_Click - config.subunidad actual es: '{sub}'");
                }
            }
            catch (Exception ex)
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Warning($"[FormInfo] No se pudo leer config.subunidad: {ex.Message}");
            }

            if (ConversorPDF.XLL.Core.Sync.SyncContext.IsInitializing)
            {
                MessageBox.Show(
                    "El sistema de sincronización se está inicializando en segundo plano.\nPor favor, espere unos segundos e intente nuevamente.",
                    "Sincronización Iniciando",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (!ConversorPDF.XLL.Core.Sync.SyncContext.IsEnabled)
            {
                MessageBox.Show(
                    "La sincronización no está habilitada.\nRegístrese en una subunidad para activarla.",
                    "Sync Deshabilitado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            this.btnSync.Enabled = false;
            this.btnSync.Text = "Sync...";

            try
            {
                await ConversorPDF.XLL.Core.Sync.SyncContext.ForcePullAsync();

                MessageBox.Show(
                    $"Datos actualizados.\nÚltima actualización: {DateTime.Now:dd/MM/yyyy HH:mm:ss}",
                    "Sincronización Completada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error durante la sincronización: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.btnSync.Enabled = true;
                this.btnSync.Text = "SYNC";
            }
        }
    }
}