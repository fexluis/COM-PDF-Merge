using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.UI;

namespace ConversorPDF.XLL.FormProgramador
{
    /// <summary>
    /// Formulario del botón "Buscar Informes": renombra los PDF de &lt;libro&gt;\Informes al
    /// estándar, unifica parejas INF + CUM y audita la columna INFORME de CONSULTAS.
    /// </summary>
    public partial class FormBuscarInformes : Form
    {
        private readonly IViaticosService _viaticosService;
        private string _carpeta;

        /// <summary>
        /// Inicializa el formulario.
        /// </summary>
        /// <param name="activeWorkbook">Libro activo (reservado para futuras validaciones).</param>
        /// <param name="viaticosService">Servicio de viáticos con el proceso de informes.</param>
        public FormBuscarInformes(dynamic activeWorkbook, IViaticosService viaticosService)
        {
            InitializeComponent();
            _viaticosService = viaticosService;

            txtViaticoInicial.KeyPress += TxtNumeric_KeyPress;
            txtViaticoFinal.KeyPress += TxtNumeric_KeyPress;
            btnExaminar.Click += BtnExaminar_Click;
            btnProcesar.Click += BtnProcesar_Click;
            btnCerrar.Click += BtnCerrar_Click;
            this.Load += FormBuscarInformes_Load;
        }

        private void FormBuscarInformes_Load(object sender, EventArgs e)
        {
            try
            {
                _carpeta = _viaticosService.ObtenerCarpetaInformes();
                txtCarpeta.Text = string.IsNullOrEmpty(_carpeta) ? "(carpeta no resuelta)" : _carpeta;

                RangoSolicitudes rango = _viaticosService.ObtenerRangoSolicitudes();
                if (rango != null && rango.HayDatos)
                {
                    txtViaticoInicial.Text = rango.Inicio.ToString();
                    txtViaticoFinal.Text = rango.Fin.ToString();
                }

                AgregarLog("Carpeta de informes: " + (string.IsNullOrEmpty(_carpeta) ? "(no resuelta)" : _carpeta));
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"FormBuscarInformes: no se pudo inicializar: {ex.Message}");
                AgregarLog("ADVERTENCIA: " + ex.Message);
            }
        }

        /// <summary> Permite solo dígitos en los campos de rango. </summary>
        private void TxtNumeric_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /// <summary> Selecciona otra carpeta de informes distinta a la del libro. </summary>
        private void BtnExaminar_Click(object sender, EventArgs e)
        {
            using (var dialogo = new FolderBrowserDialog())
            {
                dialogo.Description = "Seleccione la carpeta de informes";
                dialogo.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(_carpeta) && Directory.Exists(_carpeta))
                {
                    dialogo.SelectedPath = _carpeta;
                }

                if (dialogo.ShowDialog(this) != DialogResult.OK) return;

                _carpeta = dialogo.SelectedPath;
                txtCarpeta.Text = _carpeta;
                AgregarLog("Carpeta cambiada manualmente: " + _carpeta);
            }
        }

        private void BtnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary> Lee y valida el rango de solicitudes ingresado. </summary>
        private bool TryLeerRango(out int inicio, out int fin)
        {
            inicio = 0;
            fin = 0;

            string textoInicio = txtViaticoInicial.Text.Trim();
            string textoFin = txtViaticoFinal.Text.Trim();

            if (string.IsNullOrWhiteSpace(textoInicio) || string.IsNullOrWhiteSpace(textoFin))
            {
                MessageBox.Show("Ingrese el rango de solicitudes (inicial y final).",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!int.TryParse(textoInicio, out inicio) || !int.TryParse(textoFin, out fin))
            {
                MessageBox.Show("Ingrese números válidos (ej. 126 y 14526).",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (inicio <= 0 || fin <= 0)
            {
                MessageBox.Show("Los límites del rango deben ser mayores que cero.",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (inicio > fin)
            {
                MessageBox.Show("La solicitud inicial no puede ser mayor que la final.",
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        /// <summary> Ejecuta el proceso con barra de progreso y muestra el resumen final. </summary>
        private void BtnProcesar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_carpeta) || !Directory.Exists(_carpeta))
            {
                MessageBox.Show("La carpeta de informes no existe:\n" + (_carpeta ?? "(no definida)"),
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int inicio, fin;
            if (!TryLeerRango(out inicio, out fin)) return;

            string modo = chkSimular.Checked
                ? Environment.NewLine + Environment.NewLine + "MODO SIMULACIÓN: no se modificará ningún archivo ni celda."
                : string.Empty;

            string confirmacion =
                "Se procesarán los informes de:" + Environment.NewLine + _carpeta + Environment.NewLine + Environment.NewLine +
                "Rango de solicitudes: " + inicio + " - " + fin + Environment.NewLine +
                "Unificar parejas INF + CUM: " + (chkUnificar.Checked ? "Sí" : "No") + Environment.NewLine +
                "Actualizar columna INFORME: " + (chkAuditar.Checked ? "Sí" : "No") + modo +
                Environment.NewLine + Environment.NewLine + "¿Continuar?";

            if (MessageBox.Show(confirmacion, App.APP_NAME, MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            {
                return;
            }

            EstablecerControlesHabilitados(false);
            txtLog.Clear();
            AgregarLog("Inicio: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            try
            {
                Cursor = Cursors.WaitCursor;

                ResultadoBusquedaInformes resultado;
                using (var progreso = new ProgressManager("Buscando informes", 100))
                {
                    resultado = _viaticosService.ProcesarInformes(
                        inicio, fin, chkUnificar.Checked, chkAuditar.Checked, chkSimular.Checked, _carpeta,
                        (actual, total) => progreso.ReportProgress(actual));
                }

                MostrarDetalle(resultado);
                MostrarResumen(resultado);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("ERROR en FormBuscarInformes.BtnProcesar", ex);
                MessageBox.Show("Error al buscar informes: " + ex.Message,
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                EstablecerControlesHabilitados(true);
                AgregarLog("Fin: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }
        }

        /// <summary> Vuelca en el log el detalle por archivo/fila devuelto por el proceso. </summary>
        private void MostrarDetalle(ResultadoBusquedaInformes resultado)
        {
            if (resultado == null || resultado.Detalle == null || resultado.Detalle.Count == 0)
            {
                AgregarLog("Sin novedades que reportar.");
                return;
            }

            AgregarLog("Detalle (" + resultado.Detalle.Count + " registro(s)):");

            int mostrados = 0;
            foreach (var entrada in resultado.Detalle.OrderBy(d => d.Key))
            {
                AgregarLog("  " + entrada.Key + " -> " + entrada.Value);
                mostrados++;
                if (mostrados >= 500)
                {
                    AgregarLog("  ... (detalle truncado; consulte el log de la aplicación)");
                    break;
                }
            }
        }

        /// <summary> Muestra el resumen tipo MessageBox con los contadores del proceso. </summary>
        private void MostrarResumen(ResultadoBusquedaInformes resultado)
        {
            if (resultado == null) return;

            if (resultado.HayError)
            {
                AgregarLog("ERROR: " + resultado.Error);
                MessageBox.Show("Ocurrió un error: " + resultado.Error,
                    App.APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var mensaje = new StringBuilder();
            if (resultado.Simulado) mensaje.AppendLine("SIMULACIÓN (no se modificó nada)").AppendLine();

            mensaje.AppendLine("Renombrado:");
            mensaje.AppendLine("  PDF examinados: " + resultado.PdfsExaminados);
            mensaje.AppendLine("  Renombrados: " + resultado.Renombrados + " · ya correctos: " + resultado.YaCorrectos);
            mensaje.AppendLine("  Duplicados: " + resultado.Duplicados + " · colisiones: " + resultado.Colisiones);
            mensaje.AppendLine("  Unificados (UNI): " + resultado.Unificados + " · fallos de unión: " + resultado.MergeFallidos);
            mensaje.AppendLine("  Sin texto: " + resultado.SinTexto + " · sin comisión: " + resultado.SinComision +
                               " · sin datos: " + resultado.SinDatos);
            mensaje.AppendLine();

            mensaje.AppendLine("Auditoría INFORME:");
            mensaje.AppendLine("  Filas revisadas: " + resultado.FilasRevisadas + " · protegidas: " + resultado.OmitidasProtegidas);
            mensaje.AppendLine("  VERIFICADO: " + resultado.Verificados + " · FALTA CUM: " + resultado.FaltaCum +
                               " · FALTA INF: " + resultado.FaltaInf);
            mensaje.AppendLine("  ERROR: " + resultado.ConError + " · SIN INFORME: " + resultado.SinInforme +
                               " · PAR OK: " + resultado.ParOk);
            mensaje.AppendLine("  Celdas escritas: " + resultado.CeldasEscritas +
                               " · marcas limpiadas: " + resultado.CeldasLimpiadas);
            mensaje.AppendLine("  Libro guardado: " + (resultado.LibroGuardado ? "sí" : "no") +
                               (string.IsNullOrEmpty(resultado.ErrorGuardado) ? string.Empty : " (" + resultado.ErrorGuardado + ")"));
            mensaje.AppendLine();

            mensaje.AppendLine("Errores de E/S: " + resultado.Errores);
            mensaje.AppendLine("Tiempo: " + (resultado.TiempoMs / 1000.0).ToString("0.00") + " s");

            string texto = mensaje.ToString();
            AgregarLog(texto);

            MessageBoxIcon icono = resultado.Errores > 0 || resultado.MergeFallidos > 0
                ? MessageBoxIcon.Warning
                : MessageBoxIcon.Information;

            MessageBox.Show(texto, App.APP_NAME, MessageBoxButtons.OK, icono);
        }

        /// <summary> Habilita o deshabilita los controles durante el proceso. </summary>
        private void EstablecerControlesHabilitados(bool habilitado)
        {
            btnProcesar.Enabled = habilitado;
            btnExaminar.Enabled = habilitado;
            chkUnificar.Enabled = habilitado;
            chkAuditar.Enabled = habilitado;
            chkSimular.Enabled = habilitado;
            txtViaticoInicial.Enabled = habilitado;
            txtViaticoFinal.Enabled = habilitado;
            Application.DoEvents();
        }

        /// <summary> Añade una línea con marca de tiempo al log del formulario. </summary>
        private void AgregarLog(string linea)
        {
            if (txtLog == null) return;

            txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + (linea ?? string.Empty) + Environment.NewLine);
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
        }
    }
}
