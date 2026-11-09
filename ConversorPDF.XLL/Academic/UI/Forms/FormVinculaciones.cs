using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.UI;
using CSharpFunctionalExtensions;
using System.Data;
using Dapper;

using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Formulario para la gestión de vinculaciones académicas.
    /// Sigue el principio de separación de responsabilidades: Cero lógica de negocio en el code-behind.
    /// </summary>
    public partial class FormVinculaciones : Form
    {
        private readonly IVinculacionListService _listService;
        private readonly ISedeLookupRepository _sedeLookup;
        private readonly ITerceroLookupRepository _terceroLookup;
        private readonly Func<IDbConnection> _dbFactory;
        private readonly IDbConnection _currentConnection;
        private readonly dynamic _excelApp;
        private readonly AgModoDatos _modo;
        private readonly long? _idCorridaInicial;
        private ComboBox cmbCorrida;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly BindingSource _bsSesiones = new BindingSource();
        private IReadOnlyCollection<LookupItem> _allSedes;
        private IReadOnlyCollection<LookupItem> _allTerceros;
        private bool _isUpdatingFilters = false;
        private bool _isSyncing = false;
        private string _syncStatusText = "";

        public FormVinculaciones(
            IVinculacionListService listService,
            ISedeLookupRepository sedeLookup,
            ITerceroLookupRepository terceroLookup,
            Func<IDbConnection> dbFactory,
            dynamic excelApp,
            AgModoDatos modo = null,
            long? idCorridaInicial = null)
        {
            InitializeComponent();
            _listService = listService;
            _sedeLookup = sedeLookup;
            _terceroLookup = terceroLookup;
            _dbFactory = dbFactory;
            _excelApp = excelApp;
            _modo = modo ?? AgModoDatos.Determinista;
            _idCorridaInicial = idCorridaInicial;

            if (_modo.EsModoAg)
            {
                this.Text = $"{this.Text} [MODO AG - STAGING]";
            }

            // La conexión principal se evalúa dinámicamente vía factoría para admitir auto-recuperación
            if (_dbFactory == null)
            {
                LoggerHelper.Error("[FormVinculaciones] dbFactory provisto es null.");
                throw new ArgumentNullException(nameof(dbFactory));
            }
            
            _currentConnection = _dbFactory.Invoke();
            
            try
            {
                var conn = _currentConnection;
                if (conn == null)
                {
                    throw new InvalidOperationException("No se pudo obtener conexión a BD.");
                }
                LoggerHelper.Info($"[FormVinculaciones] Conexión principal de factoría preparada. Estado: {conn.State}");
            }
            catch (ObjectDisposedException)
            {
                LoggerHelper.Error("[FormVinculaciones] Conexión ya dispuesta al obtener State en constructor.");
            }
            
            // ✅ HOTFIX 1.2.51.25: Esta conexión es prestada. 
            // Los formularios hijos (FormViaticos, FormSesionesAcademicas) 
            // deben usar su propio contexto o respetar el ciclo de vida de esta conexión.

            // Actualizar etiqueta de usuario activo con la territorial configurada
            RefrescarEtiquetaUsuario();

            // Configuración inicial de UI
            ConfigurarGrid();
            this.Load += async (s, e) => {
                lblModo.Text = _modo.EsModoAg ? "MODO AG (staging)" : "PRODUCCIÓN";
                lblModo.ForeColor = _modo.EsModoAg ? Color.OrangeRed : Color.DimGray;
                lblModo.Font = new Font(this.Font, FontStyle.Bold);
                ReposicionarEtiquetaModo();
                this.Text = _modo.EsModoAg ? "Vinculaciones Académicas — MODO AG" : "Vinculaciones Académicas";

                // Fase 1492: el selector de corrida vive en el Designer (fila 3, bajo Semestre y Tercero)
                ConfigurarSelectorCorrida();

                SetupSyncIndicatorUI();

                // Fase 1492: catálogos → corridas → una sola consulta inicial (con todos los filtros ya resueltos)
                await CargarCatalogosAsync();
                await CargarCorridasAsync();
                await EjecutarConsultaSeguraAsync();
            };

            // La etiqueta de modo se mantiene anclada a la derecha al redimensionar
            pnlFilters.Resize += (s, e) => ReposicionarEtiquetaModo();
            
            cmbTerritoriales.SelectedIndexChanged += (s, e) => {
                if (!_isUpdatingFilters) FiltrarSedesYTerceros();
            };
            
            cmbSede.SelectedIndexChanged += (s, e) => {
                if (!_isUpdatingFilters) FiltrarTerceros();
            };
        }

        /// <summary>
        /// Configura el selector de corridas (fila 3 del panel de filtros). En PRODUCCIÓN queda oculto
        /// y el panel recupera su altura original (Fases 1492/1494).
        /// </summary>
        private void ConfigurarSelectorCorrida()
        {
            bool esModoAg = _modo.EsModoAg;

            pnlFilters.Height = esModoAg ? 104 : 80;
            lblCorrida.Visible = esModoAg;
            cmbCorrida.Visible = esModoAg;
            btnAjustarFiltrosCorrida.Visible = esModoAg;
            btnEliminarCorrida.Visible = esModoAg;

            if (!esModoAg) return;

            cmbCorrida.DisplayMember = nameof(CorridaStagingItem.Etiqueta);
            cmbCorrida.ValueMember = nameof(CorridaStagingItem.IdCorrida);

            cmbCorrida.SelectedIndexChanged += (s, e) =>
            {
                if (!_isUpdatingFilters) _ = EjecutarConsultaSeguraAsync();
            };

            // Refresco sin reabrir el formulario (sin async void sobre el hilo STA de Excel)
            cmbCorrida.DropDown += (s, e) => _ = CargarCorridasSafeAsync();

            btnAjustarFiltrosCorrida.Click += (s, e) => _ = AjustarFiltrosALaCorridaAsync();
            btnEliminarCorrida.Click += (s, e) => _ = EliminarCorridaSeleccionadaAsync();
        }

        /// <summary>
        /// Ubica la etiqueta de modo arriba a la derecha (Producción) o en la fila 3 (MODO AG),
        /// manteniéndola anclada al borde derecho del panel de filtros.
        /// </summary>
        private void ReposicionarEtiquetaModo()
        {
            if (lblModo == null || pnlFilters == null) return;

            int margen = 12;
            int y = _modo.EsModoAg ? 78 : 10;
            int x = Math.Max(margen, pnlFilters.ClientSize.Width - lblModo.Width - margen);
            lblModo.Location = new Point(x, y);
        }

        /// <summary>
        /// Elimina el staging de la corrida seleccionada (solo MODO AG), previa confirmación explícita.
        /// Borra vinculaciones, sesiones, genes, individuos, bitácoras, violaciones, promociones,
        /// salvedades y la cabecera en ag_corrida (Fases 1516/1517).
        /// </summary>
        private async Task EliminarCorridaSeleccionadaAsync()
        {
            if (!_modo.EsModoAg) return;

            if (!(cmbCorrida?.SelectedItem is CorridaStagingItem corrida) || corrida.IdCorrida <= 0)
            {
                lblStatus.Text = $"{_syncStatusText}Seleccione una corrida concreta (no «Todas las corridas») para eliminarla.";
                return;
            }

            string advertenciaExtra = corrida.IdCorrida == _idCorridaInicial
                ? "\n\nATENCIÓN: es la corrida cargada en el Módulo R; sus resultados en pantalla dejarán de estar disponibles."
                : string.Empty;

            string mensaje =
                $"¿Eliminar la corrida #{corrida.IdCorrida}?\n\n" +
                $"{corrida.PeriodoNombre} · {corrida.ProgramaNombre} · {corrida.Subunidad}\n" +
                $"Se borrarán {corrida.TotalVinculaciones} vinculaciones y {corrida.TotalSesiones} sesiones de staging." +
                advertenciaExtra +
                "\n\nEsta acción no se puede deshacer.";

            if (MessageBox.Show(this, mensaje, "Eliminar corrida de staging",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            {
                return;
            }

            btnEliminarCorrida.Enabled = false;
            try
            {
                var resultado = await _listService.EliminarCorridaStagingAsync(corrida.IdCorrida, _cts.Token);
                if (resultado.IsFailure)
                {
                    LoggerHelper.Error($"[FormVinculaciones] No se pudo eliminar la corrida #{corrida.IdCorrida}: {resultado.Error}");
                    MessageBox.Show(this, $"No se pudo eliminar la corrida #{corrida.IdCorrida}:\n{resultado.Error}",
                        "Eliminar corrida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                LoggerHelper.Info(
                    $"[FormVinculaciones] Corrida #{corrida.IdCorrida} eliminada del staging " +
                    $"({resultado.Value.Sesiones} sesiones / {resultado.Value.Vinculaciones} vinculaciones).");

                await CargarCorridasAsync();
                await EjecutarConsultaSeguraAsync();

                lblStatus.Text = $"{_syncStatusText}Corrida #{corrida.IdCorrida} eliminada ({resultado.Value.Sesiones} sesiones / {resultado.Value.Vinculaciones} vinculaciones).";
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[FormVinculaciones] Error al eliminar la corrida #{corrida.IdCorrida}", ex);
                MessageBox.Show(this, $"Error al eliminar la corrida:\n{ex.Message}", "Eliminar corrida",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnEliminarCorrida.Enabled = true;
            }
        }

        /// <summary>
        /// Carga (o recarga) el selector de corridas desde el servicio de aplicación,
        /// preservando la selección actual cuando corresponde (Fase 1492).
        /// </summary>
        private async Task CargarCorridasAsync(bool preservarSeleccion = false)
        {
            if (!_modo.EsModoAg || cmbCorrida == null) return;

            long idActual = (cmbCorrida.SelectedItem as CorridaStagingItem)?.IdCorrida ?? 0;

            var resultado = await _listService.GetCorridasStagingAsync(null, null, null, _cts.Token);
            if (resultado.IsFailure)
            {
                LoggerHelper.Warning($"[FormVinculaciones] No se pudieron cargar las corridas AG: {resultado.Error}");
                return;
            }

            var items = resultado.Value.ToList();
            var seleccion = preservarSeleccion && idActual > 0
                ? items.FirstOrDefault(c => c.IdCorrida == idActual)
                : null;

            _isUpdatingFilters = true;
            try
            {
                cmbCorrida.Items.Clear();
                foreach (var item in items) cmbCorrida.Items.Add(item);

                var elegido = seleccion ?? AgCorridaSelector.ResolverSeleccionInicial(items, _idCorridaInicial);
                var indice = elegido != null ? items.IndexOf(elegido) : -1;
                cmbCorrida.SelectedIndex = indice >= 0 ? indice : (cmbCorrida.Items.Count > 0 ? 0 : -1);
            }
            finally
            {
                _isUpdatingFilters = false;
            }
        }

        private async Task CargarCorridasSafeAsync()
        {
            try { await CargarCorridasAsync(preservarSeleccion: true); }
            catch (OperationCanceledException) { }
            catch (Exception ex) { LoggerHelper.Warning($"[FormVinculaciones] Error al refrescar corridas: {ex.Message}"); }
        }

        private async Task EjecutarConsultaSeguraAsync()
        {
            try { await EjecutarConsultaAsync(); }
            catch (OperationCanceledException) { }
            catch (Exception ex) { LoggerHelper.Warning($"[FormVinculaciones] Error al ejecutar la consulta: {ex.Message}"); }
        }

        /// <summary>
        /// Fase 1493: sincroniza los filtros visibles (período, territorial y programa)
        /// con los valores reales de la corrida seleccionada y relanza la consulta.
        /// </summary>
        private async Task AjustarFiltrosALaCorridaAsync()
        {
            if (!(cmbCorrida?.SelectedItem is CorridaStagingItem corrida) || corrida.IdCorrida <= 0)
            {
                lblStatus.Text = $"{_syncStatusText}Seleccione una corrida para ajustar los filtros.";
                return;
            }

            _isUpdatingFilters = true;
            try
            {
                SeleccionarEnCombo(cmbPeriodo, i => i.Id == corrida.IdPeriodo);
                SeleccionarEnCombo(cmbTerritoriales, i => string.Equals(
                    (i.Subunidad ?? string.Empty).Trim(),
                    (corrida.Subunidad ?? string.Empty).Trim(),
                    StringComparison.OrdinalIgnoreCase));

                FiltrarSedesYTerceros();

                SeleccionarEnCombo(cmbProgramas, i => i.Id == corrida.IdPrograma);
            }
            finally
            {
                _isUpdatingFilters = false;
            }

            await EjecutarConsultaSeguraAsync();
        }

        private static void SeleccionarEnCombo(ComboBox combo, Func<LookupItem, bool> criterio)
        {
            if (combo == null || criterio == null) return;

            foreach (var item in combo.Items)
            {
                var lookup = item as LookupItem;
                if (lookup != null && criterio(lookup))
                {
                    combo.SelectedItem = lookup;
                    return;
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _cts?.Cancel(); } catch (ObjectDisposedException) { }
                _cts?.Dispose();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void SetupSyncIndicatorUI()
        {
            try
            {
                var resolver = new ConversorPDF.XLL.Core.Sync.SyncPathResolver();
                string path = resolver.GetSyncRootFolder();
                _isSyncing = true;
                _syncStatusText = "Sincronizando | ";
            }
            catch (Exception)
            {
                _isSyncing = false;
                _syncStatusText = "Sync Inactivo | ";
            }

            this.statusStrip1.Paint += (s, ev) =>
            {
                using (var brush = new System.Drawing.SolidBrush(_isSyncing ? System.Drawing.Color.LimeGreen : System.Drawing.Color.Red))
                using (var pen = new System.Drawing.Pen(System.Drawing.Color.DarkGray))
                {
                    ev.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    // Y centrado: (22/2) - 6 = 5 (StatusStrip usualmente tiene Height=22)
                    ev.Graphics.FillEllipse(brush, 6, 5, 12, 12);
                    ev.Graphics.DrawEllipse(pen, 6, 5, 12, 12);
                }
            };
            
            // Margen izquierdo para dar espacio al círculo dibujado
            lblStatus.Margin = new Padding(22, 3, 0, 2);
        }

        private void ConfigurarGrid()
        {
            dgvVinculaciones.AutoGenerateColumns = false;
            dgvVinculaciones.ReadOnly = true;
            dgvVinculaciones.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvVinculaciones.MultiSelect = false;
            dgvVinculaciones.AllowUserToAddRows = false;

            _bsSesiones.DataSource = typeof(List<SesionListItem>);
            dgvVinculaciones.DataSource = _bsSesiones;

            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ProgramaDetalle", HeaderText = "Programa", Width = 80 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DocenteNombre", HeaderText = "Docente", Width = 250 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Semestre", HeaderText = "Sem", Width = 30 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "AsignaturaNombre", HeaderText = "Asignatura", Width = 250 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "SedeCiudad", HeaderText = "Sede", Width = 100 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DiaSemana", HeaderText = "Día Semana", Width = 50 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Horas", HeaderText = "Horas", Width = 40 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "RangoFechas", HeaderText = "Fecha Sesión", Width = 80 });
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "HorarioSesion", HeaderText = "Horario", Width = 110 });
            dgvVinculaciones.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Viatica", HeaderText = "Viática", Width = 60 });
            
            // Columna oculta con IdVinculacion para el botón Revisar
            dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                DataPropertyName = "IdVinculacion", 
                Name = "ColIdVinculacion",
                Visible = false 
            });

            // Fase 461 / Fase 471: Columnas de trazabilidad staging en MODO AG
            if (_modo.EsModoAg)
            {
                dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "IdCorrida", HeaderText = "Corrida", Width = 55 });
                dgvVinculaciones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "IdGen", HeaderText = "Gen", Width = 45 });
            }
        }

        private async Task CargarCatalogosAsync()
        {
            if (_cts.Token.IsCancellationRequested) return;

            try
            {
                var result = await _listService.GetCatalogosFiltroAsync(_cts.Token);
                if (result.IsFailure)
                {
                    MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Poblar Periodos
                cmbPeriodo.Items.Clear();
                cmbPeriodo.Items.Add(new LookupItem { Id = 0, DisplayName = "-- Todos los Períodos --" });
                foreach (var periodo in result.Value.Periodos) cmbPeriodo.Items.Add(periodo);
                cmbPeriodo.SelectedIndex = 0;

                // Poblar Programas
                cmbProgramas.Items.Clear();
                cmbProgramas.Items.Add(new LookupItem { Id = 0, DisplayName = "-- Todos los Programas --" });
                cmbProgramas.Items.Add(new LookupItem { Id = -1, DisplayName = "PREGRADO" });
                cmbProgramas.Items.Add(new LookupItem { Id = -2, DisplayName = "POSGRADO" });
                foreach (var programa in result.Value.Programas) cmbProgramas.Items.Add(programa);
                cmbProgramas.SelectedIndex = 0;

                // Poblar Territoriales y Semestres
                _isUpdatingFilters = true;
                
                try
                {
                    try
                    {
                        if (_cts.Token.IsCancellationRequested) return;
                        var repo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.AcademicRepository(_currentConnection);
                        var subResult = repo.GetSubunidades();
                        cmbTerritoriales.Items.Clear();
                        cmbTerritoriales.Items.Add(new LookupItem { Id = 0, DisplayName = "-- Todas las Territoriales --", Subunidad = null });
                        if (subResult.IsSuccess)
                        {
                            int index = 1;
                            foreach (var t in subResult.Value) 
                            {
                                cmbTerritoriales.Items.Add(new LookupItem { Id = index++, DisplayName = t, Subunidad = t });
                            }
                        }
                        else
                        {
                            LoggerHelper.Warning($"[FormVinculaciones] Error cargando territoriales: {subResult.Error}");
                            cmbTerritoriales.Items.Clear();
                            cmbTerritoriales.Items.Add(new LookupItem { Id = 0, DisplayName = "(Error al cargar - recargue el formulario)", Subunidad = null });
                        }
                        cmbTerritoriales.SelectedIndex = 0;
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error("[FormVinculaciones] Error al cargar territoriales", ex);
                        cmbTerritoriales.Items.Clear();
                        cmbTerritoriales.Items.Add(new LookupItem { Id = 0, DisplayName = "(Error al cargar - recargue el formulario)", Subunidad = null });
                        cmbTerritoriales.SelectedIndex = 0;
                    }

                    cmbSemestres.Items.Clear();
                    cmbSemestres.Items.AddRange(new string[] { "-- Todos --", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" });
                    cmbSemestres.SelectedIndex = 0;

                    _allSedes = result.Value.Sedes;
                    _allTerceros = result.Value.Terceros;
                }
                finally
                {
                    _isUpdatingFilters = false;
                }

                // Fase 7: El filtrado inicial se posterga al bloque de auto-selección para que
                // se ejecute una sola vez con los valores de configuración ya cargados.
                // Auto-selección inicial:
                // Se lee la configuración persistida (id=1) para preseleccionar la Subunidad y el Periodo.
                // Fase 8: Se usa _currentConnection directamente para evitar instanciar ConfigRepository
                // con un EncryptedSQLiteContext externo no gestionado desde un contexto async.
                // Fase 7: Guard _isUpdatingFilters activo durante la selección para evitar disparos
                // prematuros de eventos en cascada; FiltrarSedesYTerceros() se llama una sola vez al final.
                try
                {
                    // Fase 1491: en MODO AG la corrida es la fuente de verdad; los filtros ocultos de
                    // config (período/territorial) no deben aplicarse o esconden sesiones de otras vigencias.
                    bool aplicarPeriodo = AgFiltrosPolicy.AplicarPeriodoConfig(_modo);
                    bool aplicarSubunidad = AgFiltrosPolicy.AplicarSubunidadConfig(_modo);

                    _isUpdatingFilters = true;
                    try
                    {
                        if (aplicarPeriodo || aplicarSubunidad)
                        {
                            object targetIdObj = aplicarPeriodo
                                ? _currentConnection.ExecuteScalar("SELECT id_periodo FROM config WHERE id = 1")
                                : null;

                            string subunidadConfig = aplicarSubunidad
                                ? _currentConnection.QueryFirstOrDefault<string>("SELECT subunidad FROM config WHERE id = 1")
                                : null;

                            if (targetIdObj != null && targetIdObj != DBNull.Value)
                            {
                                int targetId = Convert.ToInt32(targetIdObj);
                                if (targetId > 0)
                                {
                                    foreach (var item in cmbPeriodo.Items)
                                    {
                                        if (item is LookupItem pItem && pItem.Id == targetId)
                                        {
                                            cmbPeriodo.SelectedItem = pItem;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (!string.IsNullOrEmpty(subunidadConfig))
                            {
                                foreach (var item in cmbTerritoriales.Items)
                                {
                                    if (item is LookupItem tItem && string.Equals(
                                            (tItem.Subunidad ?? string.Empty).Trim(),
                                            subunidadConfig.Trim(),
                                            StringComparison.OrdinalIgnoreCase))
                                    {
                                        cmbTerritoriales.SelectedItem = tItem;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        _isUpdatingFilters = false;
                    }

                    // Repoblar sedes/terceros (con o sin valores preseleccionados de config)
                    FiltrarSedesYTerceros();
                }
                catch (Exception exConfig)
                {
                    LoggerHelper.Warning($"[FormVinculaciones] Error al cargar configuración predeterminada: {exConfig.Message}");
                }
            }
            catch (Exception ex)
            {
                // Auditoría de ensamblados para diagnóstico de empaquetado (Excel-DNA)
                var sb = new StringBuilder();
                sb.AppendLine($"ERROR CRÍTICO en CargarCatalogosAsync: {ex.Message}");
                if (ex is FileNotFoundException fnf) sb.AppendLine($"Archivo faltante: {fnf.FileName}");
                
                sb.AppendLine("\n--- AUDITORÍA DE ENSAMBLADOS (AppDomain) ---");
                try
                {
                    var assemblies = AppDomain.CurrentDomain.GetAssemblies()
                        .OrderBy(a => a.FullName)
                        .Select(a => $"- {a.FullName} (Location: {(a.IsDynamic ? "Dynamic" : a.Location)})");
                    foreach (var asm in assemblies) sb.AppendLine(asm);
                }
                catch (Exception auditEx) { sb.AppendLine($"No se pudo completar la auditoría: {auditEx.Message}"); }
                
                LoggerHelper.Error(sb.ToString(), ex);
                MessageBox.Show($"Error al inicializar el catálogo académico. Consulte los logs para más detalles.\n\n{ex.Message}", 
                    "Error de Dependencias", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private async void btnFiltrar_Click(object sender, EventArgs e)
        {
            await EjecutarConsultaAsync();
        }

        private async Task EjecutarConsultaAsync()
        {
            if (_cts.Token.IsCancellationRequested) return;

            var programaSeleccionado = cmbProgramas.SelectedItem as LookupItem;
            int? idPrograma = null;
            string modalidadPrograma = null;

            if (programaSeleccionado != null)
            {
                if (programaSeleccionado.Id == -1) modalidadPrograma = "PREGRADO";
                else if (programaSeleccionado.Id == -2) modalidadPrograma = "POSGRADO";
                else if (programaSeleccionado.Id > 0) idPrograma = programaSeleccionado.Id;
            }

            // [Evaluación UX] Prevenir filtros ignorados silenciosamente
            // Si el control está en modo DropDown, el usuario puede teclear un texto que no existe.
            // Si hay texto pero el SelectedItem es nulo, avisamos al usuario en vez de traer toda la base de datos.
            if (!string.IsNullOrWhiteSpace(cmbTercero.Text) && cmbTercero.SelectedItem == null)
            {
                MessageBox.Show("El texto ingresado en Tercero no corresponde a una opción válida. Por favor, seleccione uno de la lista autocompletada o borre el texto.", "Validación de Filtro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            long? idCorridaFiltro = null;
            if (_modo.EsModoAg && cmbCorrida?.SelectedItem is CorridaStagingItem corridaItem && corridaItem.IdCorrida > 0)
            {
                idCorridaFiltro = corridaItem.IdCorrida;
            }

            var filtros = new VinculacionFilterDto
            {
                IdCorrida = idCorridaFiltro,
                IdPeriodo = (cmbPeriodo.SelectedItem as LookupItem)?.Id > 0 ? (int?)(cmbPeriodo.SelectedItem as LookupItem).Id : null,
                IdSede = (cmbSede.SelectedItem as LookupItem)?.Id > 0 ? (int?)(cmbSede.SelectedItem as LookupItem).Id : null,
                IdTercero = (cmbTercero.SelectedItem as LookupItem)?.Id > 0 ? (int?)(cmbTercero.SelectedItem as LookupItem).Id : null,
                IdPrograma = idPrograma,
                ModalidadPrograma = modalidadPrograma,
                Subunidad = cmbTerritoriales.SelectedIndex > 0 ? (cmbTerritoriales.SelectedItem as LookupItem)?.Subunidad : null,
                Semestre = cmbSemestres.SelectedIndex > 0 ? cmbSemestres.SelectedItem?.ToString() : null,
                SoloViaticas = chkSoloViaticas.Checked,
                FechaDesde = _modo.EsModoAg ? null : (DateTime?)dtpFechaInicio.Value.Date,
                FechaHasta = _modo.EsModoAg ? null : (DateTime?)dtpFechaFinal.Value.Date
            };

            lblStatus.Text = $"{_syncStatusText}Cargando sesiones...";
            btnFiltrar.Enabled = false;

            try
            {
                var result = await _listService.GetListaSesionesAsync(filtros, _cts.Token);
                
                if (result.IsFailure)
                {
                    MessageBox.Show(result.Error, "Error de Consulta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _bsSesiones.DataSource = new List<SesionListItem>(result.Value);
                _bsSesiones.ResetBindings(false);

                // Fase 1493: diagnóstico accionable del staging (misma fuente de verdad que la grilla)
                if (_modo.EsModoAg)
                {
                    long idCorridaVal = idCorridaFiltro ?? 0;

                    if (idCorridaVal <= 0)
                    {
                        lblStatus.Text = $"{_syncStatusText}Mostrando todas las corridas ({result.Value.Count} sesiones). Seleccione una corrida para ver su diagnóstico.";
                    }
                    else
                    {
                        var resumenRes = await _listService.GetResumenStagingAsync(idCorridaVal, filtros, _cts.Token);

                        if (resumenRes.IsFailure)
                        {
                            LoggerHelper.Warning($"[FormVinculaciones] No se pudo construir el resumen de staging: {resumenRes.Error}");
                            lblStatus.Text = $"{_syncStatusText}Se encontraron {result.Value.Count} sesiones (corrida #{idCorridaVal}).";
                        }
                        else
                        {
                            var r = resumenRes.Value;
                            string periodoUi = (cmbPeriodo.SelectedItem as LookupItem)?.DisplayName ?? "todos";
                            string subunidadUi = (cmbTerritoriales.SelectedItem as LookupItem)?.Subunidad ?? "todas";

                            if (r.SesionesCorrida > 0 && r.SesionesConFiltros == 0)
                            {
                                LoggerHelper.Warning(
                                    $"[FormVinculaciones] Corrida #{idCorridaVal} con {r.SesionesCorrida} sesiones en staging " +
                                    $"({r.PeriodosCorrida} / {r.SubunidadesCorrida}) pero 0 coinciden con los filtros activos " +
                                    $"(periodo {periodoUi}, territorial {subunidadUi}).");

                                lblStatus.Text =
                                    $"{_syncStatusText}La corrida #{idCorridaVal} tiene {r.SesionesCorrida} sesiones del período " +
                                    $"{r.PeriodosCorrida} ({r.SubunidadesCorrida}), pero los filtros activos (período {periodoUi}, " +
                                    $"territorial {subunidadUi}) dejan 0. Use «Ajustar filtros».";
                            }
                            else if (r.SesionesCorrida == 0)
                            {
                                lblStatus.Text = $"{_syncStatusText}La corrida #{idCorridaVal} no tiene sesiones en staging. Ejecute el AG o elija otra corrida.";
                            }
                            else
                            {
                                lblStatus.Text =
                                    $"{_syncStatusText}Se encontraron {r.VincCorrida} vinculaciones / {result.Value.Count} sesiones " +
                                    $"(corrida #{idCorridaVal} · {r.PeriodosCorrida} · {r.SubunidadesCorrida}).";
                            }
                        }
                    }
                }
                else
                {
                    lblStatus.Text = $"{_syncStatusText}Se encontraron {result.Value.Count} sesiones.";
                }
            }
            finally
            {
                btnFiltrar.Enabled = true;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // _cts.Cancel() se maneja ahora en Dispose
            
            // Forzar persistencia del Vault ANTES de que
            // RibbonController destruya la sesión y la conexión compartida.
            var scheduler = ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler;
            if (scheduler != null)
            {
                var sharedCtx = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext();
                if (sharedCtx != null)
                {
                    try
                    {
                        scheduler.ForcePersistVaultAsync(sharedCtx).GetAwaiter().GetResult();
                    }
                    catch (Exception ex)
                    {
                        ConversorPDF.XLL.Core.Helpers.LoggerHelper.Error("[FormVinculaciones] Error al persistir el Vault en el cierre", ex);
                    }
                }

                // Forzar el PUSH a OneDrive de los cambios recientes (fire-and-forget)
                scheduler.Flush();
            }

            base.OnFormClosing(e);
        }

        private async void btnLimpiar_Click(object sender, EventArgs e)
        {
            cmbPeriodo.SelectedIndex = 0;
            cmbTerritoriales.SelectedIndex = 0;
            cmbSemestres.SelectedIndex = 0;
            // Note: Setting cmbTerritoriales will trigger cascade and reset Sede and Tercero
            cmbProgramas.SelectedIndex = 0;
            chkSoloViaticas.Checked = false;
            dtpFechaInicio.Value = new DateTime(2026, 1, 1);
            dtpFechaFinal.Value = new DateTime(2026, 12, 31);
            await EjecutarConsultaAsync();
        }

        private async void btnRevisar_Click(object sender, EventArgs e)
        {
            if (dgvVinculaciones.SelectedRows.Count == 0)
            {
                MessageBox.Show("Seleccione una sesión para revisar.", "Revisar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var row = dgvVinculaciones.SelectedRows[0];
            var item = row.DataBoundItem as SesionListItem;
            
            // Fase 23: Defensa contra columna inexistente en la grilla para no arrojar NRE
            int fallbackId = 0;
            if (dgvVinculaciones.Columns.Contains("ColIdVinculacion"))
            {
                fallbackId = Convert.ToInt32(row.Cells["ColIdVinculacion"].Value ?? 0);
            }

            // Fase 8: Validar que IdVinculacion > 0 antes de usar el valor de DataBoundItem para evitar cargar ID=0
            int idVinculacion = (item != null && item.IdVinculacion > 0) 
                ? item.IdVinculacion 
                : fallbackId;
            
            using (var progress = new ProgressManager("Abriendo Sesiones Académicas...", 100))
            {
                using (var form = new FormSesionesAcademicas(_dbFactory, progress, idVinculacion, _modo))
                {
                    progress.Close();
                    form.ShowDialog(this);
                    
                    if (form.DatosGuardados)
                    {
                        LoggerHelper.Info("[FormVinculaciones] Refrescando grid por cambios detectados.");
                        await EjecutarConsultaAsync();
                    }
                }
            }
        }

        private async void btnAdicionar_Click(object sender, EventArgs e)
        {
            using (var progress = new ProgressManager("Abriendo Sesiones Académicas...", 100))
            {
                using (var form = new FormSesionesAcademicas(_dbFactory, progress, 0, _modo))
                {
                    progress.Close();
                    form.ShowDialog(this);
                    
                    if (form.DatosGuardados)
                    {
                        LoggerHelper.Info("[FormVinculaciones] Refrescando grid tras inserción de nueva vinculación.");
                        await EjecutarConsultaAsync();
                    }
                }
            }
        }

        private async void btnTerceros_Click(object sender, EventArgs e)
        {
            using (var form = new FormTerceros(_dbFactory))
            {
                form.ShowDialog(this);
                // Refrescar catálogo de terceros tras cerrar
                await CargarCatalogosAsync();
                await EjecutarConsultaAsync();
            }
        }

        private async void btnSedes_Click(object sender, EventArgs e)
        {
            using (var form = new FormSedes(_dbFactory))
            {
                form.ShowDialog(this);
                await CargarCatalogosAsync();
                await EjecutarConsultaAsync();
            }
        }

        private void btnGastosTerrestres_Click(object sender, EventArgs e)
        {
            LoggerHelper.Info("[FormVinculaciones] Abriendo FormGastosTerrestres.");
            using (var form = new FormGastosTerrestres(_dbFactory))
            {
                form.ShowDialog(this);
            }
        }

        private void btnVerVinculaciones_Click(object sender, EventArgs e)
        {
            var items = (_bsSesiones.DataSource as List<SesionListItem>)
                ?.ToList() ?? new List<SesionListItem>();

            if (!items.Any())
            {
                MessageBox.Show("No hay sesiones en el filtro actual.",
                    "Vinculaciones", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            LoggerHelper.Info("[FormVinculaciones] Abriendo FormVerVinculaciones con todas las sesiones.");
            // IMPORTANTE: _currentConnection es propiedad de FormVinculaciones.
            // FormVerVinculaciones NO debe llamar Dispose sobre esta conexión.
            using (var form = new FormVerVinculaciones(items, _currentConnection, _excelApp, _modo))
            {
                form.ShowDialog(this);
            }
        }

        private async void btnFestivos_Click(object sender, EventArgs e)
        {
            LoggerHelper.Info("[FormVinculaciones] Abriendo FormFestivos.");
            using (var form = new FormFestivos(_dbFactory))
            {
                form.ShowDialog(this);
                await CargarCatalogosAsync();
                await EjecutarConsultaAsync();
            }
        }

        private async void btnPeriodos_Click(object sender, EventArgs e)
        {
            LoggerHelper.Info("[FormVinculaciones] Abriendo FormPeriodos.");
            using (var form = new FormPeriodos(_dbFactory))
            {
                form.ShowDialog(this);
                await CargarCatalogosAsync();
                await EjecutarConsultaAsync();
            }
        }



        private void btnCdp_Click(object sender, EventArgs e)
        {
            using (var frm = new FormCdp(_dbFactory))
            {
                frm.ShowDialog(this);
            }
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            LoggerHelper.Info("[FormVinculaciones] Abriendo FormConfig.");
            using (var form = new FormConfig())
            {
                form.ShowDialog(this);
            }
        }

        private void btnTerminales_Click(object sender, EventArgs e)
        {
            LoggerHelper.Info("[FormVinculaciones] Abriendo FormTerminales.");
            using (var form = new FormTerminales())
            {
                form.ShowDialog(this);
            }
        }

        private void FiltrarSedesYTerceros()
        {
            if (_allSedes == null || _allTerceros == null) return;
            _isUpdatingFilters = true;

            try
            {
                string selectedTerritorial = cmbTerritoriales.SelectedIndex > 0 ? (cmbTerritoriales.SelectedItem as LookupItem)?.Subunidad : null;

                var sedesFiltradas = _allSedes.Where(s => 
                    string.IsNullOrEmpty(selectedTerritorial) || 
                    s.Subunidad == selectedTerritorial ||
                    string.IsNullOrEmpty(s.Subunidad) // Fase 14 FIX: Incluir sedes globales (compartidas)
                ).ToList();
                
                cmbSede.Items.Clear();
                cmbSede.Items.Add(new LookupItem { Id = 0, DisplayName = "-- Todas las Sedes --" });
                foreach (var sede in sedesFiltradas) cmbSede.Items.Add(sede);
                cmbSede.SelectedIndex = 0;
            }
            finally
            {
                _isUpdatingFilters = false;
            }

            FiltrarTerceros();
        }

        private void FiltrarTerceros()
        {
            if (_allSedes == null || _allTerceros == null) return;
            _isUpdatingFilters = true;

            try
            {
                string selectedTerritorial = cmbTerritoriales.SelectedIndex > 0 ? (cmbTerritoriales.SelectedItem as LookupItem)?.Subunidad : null;

                // Para terceros, sólo filtramos por Territorial si la territorial está seleccionada. 
                // Si la sede está seleccionada, no podemos filtrar directamente el catálogo de terceros por Sede, 
                // porque el maestro de terceros sólo tiene 'Subunidad', no 'IdSede'.
                // Fase 15 FIX: Soportar subunidades múltiples separadas por comas en Terceros
                var tercerosFiltrados = _allTerceros.Where(t => 
                    (string.IsNullOrEmpty(selectedTerritorial) || 
                    (t.Subunidad ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(s => s.Trim())
                                       .Any(s => string.Equals(s, selectedTerritorial, StringComparison.OrdinalIgnoreCase)))
                    && !string.Equals(t.DisplayName?.Trim(), "VACANTE", StringComparison.OrdinalIgnoreCase)
                    && t.Cedula != "1111111111"
                ).ToList();
                
                // [Fase 38 Corrección] Se evita el uso de Items.Add en bucle y se utiliza Data Binding.
                // Esto garantiza que WinForms respete DisplayMember para el autocompletado y no dependa
                // exclusivamente de la sobrescritura de ToString() de la clase de dominio LookupItem.
                var dataSourceTerceros = new List<LookupItem> { new LookupItem { Id = 0, DisplayName = "-- Todos los Terceros --" } };
                dataSourceTerceros.AddRange(tercerosFiltrados);
                
                cmbTercero.DataSource = null; // Limpiar binding anterior
                cmbTercero.DisplayMember = "DisplayName";
                cmbTercero.ValueMember = "Id";
                cmbTercero.DataSource = dataSourceTerceros;
            }
            finally
            {
                _isUpdatingFilters = false;
            }
        }

        public async void RecargarCatalogosUI()
        {
            if (this.IsDisposed) return;
            await CargarCatalogosAsync();
        }

        private void RefrescarEtiquetaUsuario()
        {
            string subunidad = "";
            try 
            {
                subunidad = _currentConnection.QueryFirstOrDefault<string>("SELECT subunidad FROM config LIMIT 1") ?? "";
            } 
            catch (Exception) { /* Ignorar error inicial de config */ }
            
            if (lblUsuarioActivo != null)
            {
                lblUsuarioActivo.Text = $"Usuario: {Environment.UserName} | Territorial: {subunidad}";
            }
        }
    }
}
