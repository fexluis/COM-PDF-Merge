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
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
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
        private Label lblModo;
        private Label lblCorrida;
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
            AgModoDatos modo = null)
        {
            InitializeComponent();
            _listService = listService;
            _sedeLookup = sedeLookup;
            _terceroLookup = terceroLookup;
            _dbFactory = dbFactory;
            _excelApp = excelApp;
            _modo = modo ?? AgModoDatos.Determinista;

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
                if (lblModo == null)
                {
                    lblModo = new Label
                    {
                        AutoSize = true,
                        Location = new Point(pnlFilters.Width - 180, 10),
                        Anchor = AnchorStyles.Top | AnchorStyles.Right
                    };
                    pnlFilters.Controls.Add(lblModo);
                }
                lblModo.Text = _modo.EsModoAg ? "MODO AG (staging)" : "PRODUCCIÓN";
                lblModo.ForeColor = _modo.EsModoAg ? Color.OrangeRed : Color.DimGray;
                lblModo.Font = new Font(this.Font, FontStyle.Bold);
                this.Text = _modo.EsModoAg ? "Vinculaciones Académicas — MODO AG" : "Vinculaciones Académicas";

                // Fase 457: En Modo AG, agregar selector dinámico de corrida para filtrar staging fielmente
                if (_modo.EsModoAg && cmbCorrida == null)
                {
                    lblCorrida = new Label
                    {
                        Text = "Corrida:",
                        AutoSize = true,
                        Location = new Point(pnlFilters.Width - 365, 10),
                        Anchor = AnchorStyles.Top | AnchorStyles.Right
                    };
                    cmbCorrida = new ComboBox
                    {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Width = 175,
                        Location = new Point(pnlFilters.Width - 310, 7),
                        Anchor = AnchorStyles.Top | AnchorStyles.Right
                    };
                    pnlFilters.Controls.Add(lblCorrida);
                    pnlFilters.Controls.Add(cmbCorrida);

                    cmbCorrida.SelectedIndexChanged += async (sCorr, eCorr) =>
                    {
                        if (!_isUpdatingFilters) await EjecutarConsultaAsync();
                    };

                    try
                    {
                        var corridas = (await _currentConnection.QueryAsync<LookupItem>(
                            "SELECT id AS Id, ('Corrida #' || id || ' (' || COALESCE(fecha_inicio_corrida, '') || ')') AS DisplayName FROM ag_corrida ORDER BY id DESC")).ToList();
                        cmbCorrida.Items.Clear();
                        cmbCorrida.Items.Add(new LookupItem { Id = 0, DisplayName = "-- Todas las corridas --" });
                        foreach (var c in corridas) cmbCorrida.Items.Add(c);
                        if (cmbCorrida.Items.Count > 1) cmbCorrida.SelectedIndex = 1; // Default: última corrida
                        else cmbCorrida.SelectedIndex = 0;
                    }
                    catch (Exception exCorr)
                    {
                        LoggerHelper.Warning($"[FormVinculaciones] Error al cargar corridas de AG: {exCorr.Message}");
                    }
                }

                SetupSyncIndicatorUI();
                await CargarCatalogosAsync();
            };
            
            cmbTerritoriales.SelectedIndexChanged += (s, e) => {
                if (!_isUpdatingFilters) FiltrarSedesYTerceros();
            };
            
            cmbSede.SelectedIndexChanged += (s, e) => {
                if (!_isUpdatingFilters) FiltrarTerceros();
            };
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
                    var targetIdObj = _currentConnection.ExecuteScalar("SELECT id_periodo FROM config WHERE id = 1");
                    var subunidadConfig = _currentConnection.QueryFirstOrDefault<string>("SELECT subunidad FROM config WHERE id = 1");

                    _isUpdatingFilters = true;
                    try
                    {
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
                                if (item is LookupItem tItem && tItem.Subunidad == subunidadConfig)
                                {
                                    cmbTerritoriales.SelectedItem = tItem;
                                    break;
                                }
                            }
                        }
                    }
                    finally
                    {
                        _isUpdatingFilters = false;
                    }

                    // Refiltrar una sola vez con ambos valores ya establecidos
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
            if (_modo.EsModoAg && cmbCorrida?.SelectedItem is LookupItem corridaItem && corridaItem.Id > 0)
            {
                idCorridaFiltro = corridaItem.Id;
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

                // Fase 457 / Fase 469: En Modo AG, reportar conteo conjunto de vinculaciones y sesiones de la corrida
                if (_modo.EsModoAg)
                {
                    long idCorridaVal = idCorridaFiltro ?? 0;
                    int vincCount = await _currentConnection.ExecuteScalarAsync<int>(
                        "SELECT COUNT(*) FROM ag_vinculaciones WHERE (@idCorrida = 0 OR id_corrida = @idCorrida) AND estado = 0",
                        new { idCorrida = idCorridaVal });

                    if (result.Value.Count == 0 && vincCount > 0)
                    {
                        int totalSesStaging = await _currentConnection.ExecuteScalarAsync<int>(
                            "SELECT COUNT(*) FROM ag_sesiones s JOIN ag_vinculaciones v ON s.idvinculacion = v.id WHERE (@idCorrida = 0 OR v.id_corrida = @idCorrida) AND v.estado = 0",
                            new { idCorrida = idCorridaVal });
                        if (totalSesStaging > 0)
                        {
                            LoggerHelper.Warning($"[FormVinculaciones] Modo AG: Existen {totalSesStaging} sesiones en staging pero la consulta devolvió 0 filas. Verifique los filtros activos o la vista vw_AgSesionesDetalle.");
                            lblStatus.Text = $"{_syncStatusText}Aviso: Hay {totalSesStaging} sesiones en staging pero 0 coinciden con filtros (corrida #{idCorridaVal}).";
                            return;
                        }
                    }

                    lblStatus.Text = $"{_syncStatusText}Se encontraron {vincCount} vinculaciones / {result.Value.Count} sesiones (corrida #{idCorridaVal}).";
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
