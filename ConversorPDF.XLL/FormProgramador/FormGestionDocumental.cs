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
            this.Load += FormGestionDocumental_Load;
        }

        /// <summary>
        /// Prellena el rango con el primer y último registro de la columna SOLICITUD (B) de CONSULTAS.
        /// </summary>
        private void FormGestionDocumental_Load(object sender, EventArgs e)
        {
            try
            {
                RangoSolicitudes rango = _viaticosService.ObtenerRangoSolicitudes();
                if (!rango.HayDatos) return;

                txtViaticoInicial.Text = rango.Inicio.ToString();
                txtViaticoFinal.Text = rango.Fin.ToString();
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"FormGestionDocumental: no se pudo prellenar el rango: {ex.Message}");
            }
        }

        private void TxtNumeric_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// Lee y valida el rango de viáticos ingresado en el formulario.
        /// </summary>
        /// <param name="inicio">Límite inicial del rango.</param>
        /// <param name="fin">Límite final del rango.</param>
        /// <returns>True si el rango es válido; de lo contrario false con mensaje al usuario.</returns>
        private bool TryLeerRango(out int inicio, out int fin)
        {
            inicio = 0;
            fin = 0;

            string tIni = txtViaticoInicial.Text.Trim();
            string tFin = txtViaticoFinal.Text.Trim();

            if (string.IsNullOrWhiteSpace(tIni) || string.IsNullOrWhiteSpace(tFin))
            {
                MessageBox.Show("Por favor, ingrese el rango de viáticos (Inicial y Final).", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!int.TryParse(tIni, out inicio) || !int.TryParse(tFin, out fin))
            {
                MessageBox.Show("Ingrese números válidos (ej. 126 y 14526).", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (inicio > fin)
            {
                MessageBox.Show("El viático inicial no puede ser mayor que el final.", App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void BtnRenombrar_Click(object sender, EventArgs e)
        {
            if (!TryLeerRango(out int inicio, out int fin))
            {
                return;
            }

            var r = _viaticosService.RenombrarSolicitudesConDetalle(inicio, fin);

            if (r.HayError)
            {
                MessageBox.Show($"Ocurrió un error al renombrar solicitudes: {r.Error}",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ResultadoVerificacionSolicitudes v = null;
            if (r.ComisionesEnRango > 0)
            {
                v = _viaticosService.VerificarSolicitudesPresentes(inicio, fin);
            }

            string mensaje = $"Renombrado: {r.Renombrados} · duplicados: {r.Duplicados} · ya correctos: {r.YaCorrectos}\r\n" +
                             $"Sin comisión: {r.SinComision} · sin datos: {r.SinDatos} · colisiones: {r.Colisiones} · " +
                             $"ambiguos: {r.Ambiguos} · prefijo inválido: {r.PrefijoInvalido}\r\n" +
                             $"PDF examinados: {r.PdfsExaminados} · comisiones en rango: {r.ComisionesEnRango}";

            bool conCambios = r.Renombrados > 0 || r.Duplicados > 0;
            MessageBoxIcon icono = conCambios ? MessageBoxIcon.Information : MessageBoxIcon.Warning;

            if (v != null)
            {
                if (v.HayError)
                {
                    mensaje += $"\r\n\r\nAuditoría de solicitudes: ERROR — {v.Error}";
                    icono = MessageBoxIcon.Warning;
                }
                else
                {
                    mensaje += $"\r\n\r\nAuditoría de solicitudes (columna INFORME):\r\n" +
                               $"Marcas escritas: {v.CeldasEscritas} · marcas corregidas: {v.CeldasLimpiadas} · " +
                               $"sin archivo: {v.SinArchivo}\r\n" +
                               $"Protegidas (VERIFICADO/REINTEGRO): {v.OmitidasProtegidas} · otras marcas: {v.OmitidasOtrasMarcas} · " +
                               $"ya marcadas: {v.YaMarcadas}\r\n" +
                               $"Guardado: {(v.LibroGuardado ? "sí" : "no")}" +
                               (string.IsNullOrEmpty(v.ErrorGuardado) ? string.Empty : $" ({v.ErrorGuardado})");

                    if (v.CeldasEscritas + v.CeldasLimpiadas > 0) icono = MessageBoxIcon.Information;
                }
            }

            if (r.SinComision > 0)
            {
                mensaje += $"\r\n\r\nHay {r.SinComision} comisión(es) del rango sin registro en CONSULTAS: actualice la hoja para poder renombrarlas.";
            }

            MessageBox.Show(mensaje, App.APP_NAME, MessageBoxButtons.OK, icono);
        }

        private void BtnUnificar_Click(object sender, EventArgs e)
        {
            if (!TryLeerRango(out int inicio, out int fin))
            {
                return;
            }

            int n = _viaticosService.UnificarSolicitudes(inicio, fin);
            MessageBox.Show(
                n > 0 ? $"Unificación completada: {n} expediente(s)."
                      : n == 0 ? "No se encontraron expedientes para el rango indicado."
                      : "Ocurrió un error al unificar solicitudes. Consulte el registro de eventos.",
                App.APP_NAME,
                MessageBoxButtons.OK,
                n >= 0 ? (n > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning) : MessageBoxIcon.Error);
        }
    }
}
