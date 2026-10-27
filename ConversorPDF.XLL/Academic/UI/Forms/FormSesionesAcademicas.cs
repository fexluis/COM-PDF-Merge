using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Core.Repositories;
using System.Runtime.InteropServices;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.UI;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Formulario de gestión de sesiones académicas.
    /// </summary>
    public partial class FormSesionesAcademicas : Form
    {
        private readonly GeneradorSesionesRecurrentes _generador;
        private List<Sesion> _sesionesGeneradas;

        // Fase 26: El default correcto de ID para indicar nuevo es 0, evitando inconsistencias iniciales.
        private int _currentVinculacionId = 0;
        private CancellationTokenSource _guardadoCts;

        /// <summary>
        /// Indica si se han guardado cambios durante la sesión actual del formulario.
        /// </summary>
        public bool DatosGuardados { get; private set; } = false;

        private readonly Func<IDbConnection> _connectionFactory;
        private readonly IDbConnection _sharedConn;
        private readonly AgModoDatos _modo;
        private Label lblModo;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (lblModo == null)
            {
                lblModo = new Label
                {
                    AutoSize = true,
                    Location = new Point(620, 12),
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                this.Controls.Add(lblModo);
            }
            lblModo.Text = _modo.EsModoAg ? "MODO AG (staging)" : "PRODUCCIÓN";
            lblModo.ForeColor = _modo.EsModoAg ? Color.OrangeRed : Color.DimGray;
            lblModo.Font = new Font(this.Font, FontStyle.Bold);
            this.Text = _modo.EsModoAg ? "Sesiones Académicas [MODO AG - STAGING]" : "Sesiones Académicas";
        }

        /// <summary>
        /// Guarda un "Hash" con el estado de los controles que determinan la generación
        /// (Tercero, Horas y Distribución de días). Se fija en [Generar] y tras guardar/sincronizar.
        /// Fase 699 / Fase 702: los segmentos FI/FF se refrescan aparte al editar las fechas maestras
        /// (SincronizarFechasEnEstadoGenerado), preservando T/H/D.
        /// </summary>
        private string _ultimoEstadoGenerado = string.Empty;
        private readonly IAcademicDateSyncService _dateSyncService;

        // FASE 237 / F1.5: Bandera para evitar disparos recursivos en CellValueChanged durante sincronización interna
        private bool _sincronizandoGrid = false;
        // Fase 699: evita re-entrada recursiva al auto-ampliar la ventana maestra desde sus propios handlers
        private bool _ajustandoFechaMaestra = false;
        private static readonly CultureInfo CulturaEs = new CultureInfo("es-ES");
        private List<Festivo> _rangosFestivos = new List<Festivo>();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        private const uint PBM_SETBARCOLOR = 0x0409;

        private class TerceroDisplayDto
        {
            public int Id { get; set; }
            public string Nombre { get; set; }
            public string Residencia { get; set; }
            public string DisplayName => string.IsNullOrEmpty(Residencia) ? Nombre : $"{Nombre} ({Residencia})";
            public override string ToString() => DisplayName;
        }

        private static string NormalizeCityName(string city)
        {
            if (string.IsNullOrEmpty(city)) return city;
            var normalized = city.ToUpperInvariant().Replace(".", "").Trim();
            normalized = new string(normalized
                .Normalize(System.Text.NormalizationForm.FormD)
                .Where(c => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                .ToArray());
            while (normalized.Contains("  ")) normalized = normalized.Replace("  ", " ");
            return normalized;
        }

        public FormSesionesAcademicas(Func<IDbConnection> connectionFactory, ProgressManager progress = null, int? idVinculacion = null, AgModoDatos modo = null)
        {
            _connectionFactory = connectionFactory;
            _modo = modo ?? AgModoDatos.Determinista;
            _sharedConn = _connectionFactory?.Invoke();
            if (_sharedConn != null && _sharedConn.State != ConnectionState.Open)
            {
                try { _sharedConn.Open(); } catch { }
            }

            _generador = new GeneradorSesionesRecurrentes();
            _sesionesGeneradas = new List<Sesion>();
            _dateSyncService = new AcademicDateSyncService();

            if (idVinculacion.HasValue)
            {
                if (idVinculacion.Value == 0)
                {
                    _currentVinculacionId = 0; // Modo NUEVO REGISTRO
                    LoggerHelper.Info("[FormSesionesAcademicas] Modo Nuevo: Creando vinculación nueva");
                }
                else
                {
                    _currentVinculacionId = idVinculacion.Value; // Modo Revisar
                    LoggerHelper.Info($"[FormSesionesAcademicas] Modo Revisar: Cargando Vinculación Id={_currentVinculacionId}");
                }
            }
            else
            {
                _currentVinculacionId = 0; // Modo Default (Ribbon)
                LoggerHelper.Info("[FormSesionesAcademicas] Modo Default: Creando vinculación nueva (Id=0)");
            }

            progress?.AddProgress(10); // 10% - Inicio

            // 1. Asegurar motor (Nivel 1)
            SQLiteEngineBootstrap.EnsureInitialized();
            progress?.AddProgress(10); // 20% - Motor listo

            InitializeComponent();
            this.Text = _modo.EsModoAg ? "Sesiones Académicas [MODO AG - STAGING]" : "Sesiones Académicas";
            pnlLoading.BringToFront(); // Asegurar que cubra todo al inicio
            SetupGrid();

            // Sincronizar fecha final si la nueva fecha de inicio es mayor
            dtpInicio.ValueChanged += DtpInicio_ValueChanged;
            dtpFinal.ValueChanged += DtpFinal_ValueChanged;

            // Restringir ingreso en Gasto Terminal Aérea a solo números enteros
            txtGastoTerminalAerea.KeyPress += (s, e) =>
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                }
            };

            txtGastoTerminalAerea.TextChanged += (s, e) =>
            {
                var tb = (TextBox)s;
                int selStart = tb.SelectionStart;
                string cleaned = new string(tb.Text.Where(char.IsDigit).ToArray());

                if (!string.IsNullOrEmpty(cleaned) && double.TryParse(cleaned, out double valorDigitado))
                {
                    var terminal = cmbTerminalesAereas.SelectedItem as Terminal;
                    // Fase 25: Evitar bypass. Si no hay terminal válido, no se puede ingresar valor.
                    if (terminal == null)
                    {
                        cleaned = string.Empty;
                    }
                    else if (valorDigitado > terminal.MaxValor)
                    {
                        cleaned = terminal.MaxValor.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                }

                if (tb.Text != cleaned)
                {
                    tb.Text = cleaned;
                    tb.SelectionStart = Math.Min(selStart, tb.Text.Length);
                }
            };
        }

        /// <summary>
        /// Wrapper efímero para interactuar con la base de datos. 
        /// Si el contexto compartido está abierto (OnShown), lo reutiliza para evitar I/O redundante.
        /// </summary>
        private void InvokeRepository(Action<IAcademicRepository> repoAction)
        {
            if (_sharedConn != null && _sharedConn.State == ConnectionState.Open)
            {
                repoAction(new AcademicRepository(_sharedConn, modo: _modo));
                return;
            }

            LoggerHelper.Warning("[FormSesionesAcademicas] InvokeRepository: Contexto compartido no disponible. Usando fallback efímero (No recomendado en v1.2.50.0).");
            using (var ctx = SqliteHelper.GetContextForPath())
            {
                ctx.MarkAsReadOnly();
                using (var conn = ctx.Open())
                {
                    repoAction(new AcademicRepository(conn, modo: _modo));
                }
            }
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);

            try
            {
                // Mostramos el loader explícitamente y forzamos el repintado
                pnlLoading.Visible = true;
                pnlLoading.BringToFront();

                // Forzar color azul en la barra de progreso (ESAP Blue: 0, 120, 212)
                int blueColor = ColorTranslator.ToWin32(Color.FromArgb(0, 120, 212));
                if (pgbLoading.IsHandleCreated)
                    SendMessage(pgbLoading.Handle, PBM_SETBARCOLOR, IntPtr.Zero, (IntPtr)blueColor);
                if (pgbGuardar.IsHandleCreated)
                    SendMessage(pgbGuardar.Handle, PBM_SETBARCOLOR, IntPtr.Zero, (IntPtr)blueColor);

                System.Windows.Forms.Application.DoEvents();


                // B.1 - Obtener Conexión Compartida (Dinámica)
                LoggerHelper.Info($"[FormSesionesAcademicas] OnShown: Conexión dinámica preparada.");

                // 1. Inyectar datos iniciales si es necesario (Catálogos)
                // Fase 10: Evitar Task.Run para evitar posibles SQLITE_BUSY con SQLite sobre _sharedConn en OnShown
                ExecuteSeedData();

                // 2. Cargar datos iniciales en combos
                // Nota: LoadInitialData y CargarEstadoPrevio tocan la UI, 
                // así que internamente deben manejar sus propios Task.Run para la parte de IO
                LoadInitialData(null);

                // 3. Restaurar estado guardado (Vinculación y Sesiones previas)
                CargarEstadoPrevio();

                await VerificarPerfilAdmin(); // Fase 16: Await the async Task

                // Pequeña pausa para asegurar que el usuario vea que terminó
                await Task.Delay(500);

                // 4. Ocultar loader tras carga exitosa
                pnlLoading.Visible = false;
                pnlLoading.SendToBack();

                this.BeginInvoke((Action)(() =>
                {
                    this.Activate();
                    dgvSesiones.Focus();
                }));
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en OnShown al cargar datos iniciales", ex);
                pnlLoading.Visible = false;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _guardadoCts?.Cancel();

            // Liberar conexión compartida (gestionada por la sesión principal)
            LoggerHelper.Info("[FormSesionesAcademicas] OnFormClosing disparado...");

            var scheduler = ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler;
            if (this.DatosGuardados && scheduler != null)
            {
                LoggerHelper.Info("[FormSesionesAcademicas] Esperando finalización de tareas asíncronas pendientes...");
                scheduler.WaitForPendingSync(TimeSpan.FromSeconds(5));
            }

            // Forzar el PUSH a OneDrive de los cambios recientes (fire-and-forget)
            scheduler?.Flush();

            base.OnFormClosing(e);
        }

        private void SetupGrid()
        {
            dgvSesiones.AutoGenerateColumns = false;
            dgvSesiones.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvSesiones.MultiSelect = true;
            dgvSesiones.RowHeadersVisible = false;
            dgvSesiones.BackgroundColor = Color.White;
            dgvSesiones.BorderStyle = BorderStyle.None;
            dgvSesiones.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            dgvSesiones.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(243, 244, 246);
            dgvSesiones.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvSesiones.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvSesiones.EnableHeadersVisualStyles = false;

            // FASE 239 / F1.10: Deshabilitar adición de filas por el usuario (elimina la fila vacía "16 de 17")
            dgvSesiones.AllowUserToAddRows = false;

            // FASE 239: Eliminar DataPropertyName en grilla no enlazada (un-bound) para evitar bugs de WinForms
            // y marcar columnas de solo lectura (ColTercero, ColDias)
            dgvSesiones.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ColChecked", HeaderText = "", Width = 30 });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColTercero", HeaderText = "Tercero", Width = 230, ReadOnly = true });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDias", HeaderText = "Día Semana", Width = 120, ReadOnly = true });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColHoras", HeaderText = "Horas", Width = 60 });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColHorario", HeaderText = "Horario Sesión", Width = 120 });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColFechaSesion", HeaderText = "Fecha Sesión", Width = 100 });
            dgvSesiones.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ColViatica", HeaderText = "Viática", Width = 60 });

            // FASE 239 / F1.11: Deshabilitar ordenamiento automático en todas las columnas para evitar orden lexicográfico
            // sobre cadenas de fecha y preservar la correspondencia fila <-> Tag
            foreach (DataGridViewColumn c in dgvSesiones.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void LoadInitialData(ProgressManager progress = null)
        {
            try
            {
                LoggerHelper.Info("[FormSesionesAcademicas] LoadInitialData: Iniciando carga de catálogos...");
                var deptoBase = GetDepartamentoBase();
                if (!string.IsNullOrEmpty(deptoBase))
                {
                    LoggerHelper.Info($"[FormSesionesAcademicas] Aplicando filtro de Departamento Base: {deptoBase}");
                }

                progress?.AddProgress(10); // 30% - Desencriptando...

                // FASE 2: Inyección obligatoria del contexto compartido
                LoggerHelper.Info("[FormSesionesAcademicas] Asegurando caché de catálogos v1.2.50.0 (Shared Connection Optimized)...");
                if (_sharedConn != null)
                {
                    CatalogCacheService.EnsureLoaded(_sharedConn);
                }
                else
                {
                    throw new InvalidOperationException("FALLO DE ARQUITECTURA: Shared Context no inicializado en OnShown.");
                }

                progress?.AddProgress(10); // 40%

                // 1. Periodos
                if (CatalogCacheService.Periodos.Any())
                {
                    cmbPeriodos.DisplayMember = "PeriodoNombre";
                    cmbPeriodos.ValueMember = "IdPeriodo";
                    cmbPeriodos.DataSource = CatalogCacheService.Periodos;
                }
                else
                {
                    LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando periodos: Caché vacío.");
                    // Fase 24: Feedback en UI cuando un catálogo no carga
                    MessageBox.Show("No se pudieron cargar los períodos académicos. Verifique la conexión a la base de datos.", "Advertencia de Catálogo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                progress?.AddProgress(5); // 45%

                // 1.5 Sedes
                try
                {
                    var lookupRepo = new SedeLookupRepository(() => _sharedConn);
                    // Fase 11: Prevenir deadlock de .GetAwaiter().GetResult() encapsulando en Task.Run para cambiar al ThreadPool
                    var lookupResult = Task.Run(() => lookupRepo.GetSedesLookupAsync(System.Threading.CancellationToken.None)).GetAwaiter().GetResult();
                    if (lookupResult.IsSuccess && lookupResult.Value.Any())
                    {
                        cmbSedes.DisplayMember = "DisplayName";
                        cmbSedes.ValueMember = "Id";
                        cmbSedes.DataSource = lookupResult.Value.ToList();
                    }
                    else
                    {
                        LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando sedes: Sin datos de lookup.");
                        // Fase 25: Feedback en UI
                        MessageBox.Show("No se pudieron cargar las Sedes.", "Advertencia de Catálogo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[FormSesionesAcademicas] Error cargando sedes lookup", ex);
                }
                progress?.AddProgress(10); // 50%

                // 2. Programas
                LoggerHelper.Info("[FormSesionesAcademicas] Aplicando Programas desde Caché...");
                if (CatalogCacheService.Programas.Any())
                {
                    cmbProgramas.DisplayMember = "Abreviatura";
                    cmbProgramas.ValueMember = "Id";
                    cmbProgramas.DataSource = CatalogCacheService.Programas;
                }
                else
                {
                    LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando programas: Caché vacío.");
                    // Fase 26: Feedback en UI
                    MessageBox.Show("No se pudieron cargar los Programas.", "Advertencia de Catálogo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                progress?.AddProgress(10); // 60%

                // 2.5 Grupos
                cmbGrupos.Items.Clear();
                cmbGrupos.Items.AddRange(new string[] { "--", "A", "B", "C", "D" });
                cmbGrupos.SelectedIndex = 0;

                // Tipo Tutoría
                cmbTipoTutoria.Items.Clear();
                cmbTipoTutoria.Items.AddRange(ModalidadTutoriaCatalog.Opciones.Cast<object>().ToArray());
                cmbTipoTutoria.SelectedIndex = 0;

                // 3. Semestres
                cmbSemestres.Items.Clear();
                cmbSemestres.Items.AddRange(new string[] { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" });
                cmbSemestres.SelectedIndex = 0;

                // 4. Terceros
                LoggerHelper.Info("[FormSesionesAcademicas] Aplicando Terceros desde Caché...");
                if (CatalogCacheService.Terceros.Any())
                {
                    var tercerosList = CatalogCacheService.Terceros;

                    if (!string.IsNullOrEmpty(deptoBase))
                    {
                        var filtered = tercerosList.Where(t => t.Departamento != null && t.Departamento.ToUpper().Trim() == deptoBase).ToList();
                        if (filtered.Any()) tercerosList = filtered;
                    }

                    var displayList = tercerosList
                        .Where(t => !string.Equals(t.Nombre?.Trim(), "VACANTE", StringComparison.OrdinalIgnoreCase) && t.Cedula != "1111111111")
                        .OrderBy(t => t.Nombre)
                        .Select(t => new TerceroDisplayDto
                        {
                            Id = t.Id,
                            Nombre = t.Nombre,
                            Residencia = t.Residencia
                        }).ToList();

                    cmbTerceros.DisplayMember = "DisplayName";
                    cmbTerceros.ValueMember = "Id";
                    cmbTerceros.DataSource = displayList;
                }
                else LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando terceros: Caché vacío.");
                progress?.AddProgress(20); // 80%

                // 5. Terminales Aéreos
                try
                {
                    InvokeRepository(repo =>
                    {
                        var termResult = repo.GetTerminales();
                        if (termResult.IsSuccess && termResult.Value.Any())
                        {
                            cmbTerminalesAereas.DisplayMember = "TerminalName";
                            cmbTerminalesAereas.ValueMember = "Id";
                            cmbTerminalesAereas.DataSource = termResult.Value.ToList();
                        }
                    });
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[FormSesionesAcademicas] Error cargando terminales", ex);
                }

                // Las fechas se restaurarán desde la DB en CargarEstadoPrevio si hay estado previo.
                // Si no, usamos los valores por defecto. Asignar FechaFin primero evita que ValueChanged de FechaInicio la sobrescriba (Fase 1)
                dtpFinal.Value = DateTime.Today.AddMonths(1);
                dtpInicio.Value = DateTime.Today;

                progress?.AddProgress(20); // 100% - Completo
                LoggerHelper.Info("[FormSesionesAcademicas] LoadInitialData: Carga completada con éxito.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en LoadInitialData", ex);
                MessageBox.Show($"Error cargando datos: {ex.Message}", "Error Diagnostics", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DtpInicio_ValueChanged(object sender, EventArgs e)
        {
            // Invariante existente: Fecha Final nunca anterior a Fecha Inicio
            if (dtpFinal.Value < dtpInicio.Value)
            {
                dtpFinal.Value = dtpInicio.Value;
            }

            AmpliarVentanaMaestraAlGrid();          // Fase 699
            SincronizarFechasEnEstadoGenerado();    // Fase 699
        }

        private void DtpFinal_ValueChanged(object sender, EventArgs e)
        {
            AmpliarVentanaMaestraAlGrid();          // Fase 699 — "comprueba la última fecha"
            SincronizarFechasEnEstadoGenerado();    // Fase 699
        }

        private void cmbProgramas_SelectedIndexChanged(object sender, EventArgs e) => UpdateAsignaturas();
        private void cmbSemestres_SelectedIndexChanged(object sender, EventArgs e) => UpdateAsignaturas();

        private void UpdateAsignaturas()
        {
            int idProg = Convert.ToInt32(cmbProgramas.SelectedValue ?? 0);
            string sem = cmbSemestres.SelectedItem?.ToString();

            if (idProg == 0 || string.IsNullOrWhiteSpace(sem))
            {
                LoggerHelper.Info("[FormSesionesAcademicas] UpdateAsignaturas: Saltando (selección incompleta).");
                return;
            }

            try
            {
                LoggerHelper.Info($"[FormSesionesAcademicas] UpdateAsignaturas: Filtrando para ProgId={idProg}, Semestre={sem}...");
                InvokeRepository(repo =>
                {
                    var result = repo.GetAsignaturasByPrograma(idProg, sem);
                    if (result.IsSuccess)
                    {
                        var filtered = result.Value.ToList();

                        cmbAsignaturas.DisplayMember = "Nombre";
                        cmbAsignaturas.ValueMember = "Id";
                        cmbAsignaturas.DataSource = filtered;

                        LoggerHelper.Info($"[FormSesionesAcademicas] Asignaturas filtradas vía SQL: {filtered.Count}");
                    }
                    else
                    {
                        LoggerHelper.Warning($"[FormSesionesAcademicas] Error obteniendo asignaturas: {result.Error}");
                    }
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en UpdateAsignaturas", ex);
            }
        }

        private void cmbAsignaturas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbAsignaturas.SelectedItem is Asignatura asig)
            {
                lblIntHoraria.Text = $"     Total: {asig.IntHoraria}h";
            }
            else
            {
                lblIntHoraria.Text = "";
            }
        }

        /// <summary>
        /// Popula catálogos en la base de datos temporal (MEMORIA) para el uso actual del formulario.
        /// B.4 y B.5: Implementa DDL idempotente y Seed Unificado vía Recurso Embebido.
        /// </summary>
        private void ExecuteSeedData()
        {
            try
            {
                InvokeRepository(repo =>
                {
                    // 1. Verificación de Esquema Idempotente (B.4)
                    // Usamos una conexión directa a través del repo (Dapper) para verificar columnas
                    bool needsMigration = true;
                    try
                    {
                        // Fase 15: Evitar NRE si _sharedConn es nulo (fallback). Usar reflexión segura si la interfaz no expone Connection.
                        var activeConn = _sharedConn;
                        if (activeConn == null)
                        {
                            var field = repo.GetType().GetField("_connection", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                            activeConn = field?.GetValue(repo) as IDbConnection;
                        }

                        if (activeConn != null)
                        {
                            var cols = activeConn.Query<dynamic>("PRAGMA table_info(vinculaciones);");
                            if (cols.Any(c => ((string)c.name).ToLower() == "idtercero"))
                            {
                                needsMigration = false;
                            }
                        }
                    }
                    catch { /* Fallback a True si falla el PRAGMA */ }

                    // La migración de esquema está ahora delegada al SqliteHelper.
                    // Ya no rastreamos variables locales DDL aquí.
                    if (needsMigration)
                    {
                        LoggerHelper.Info("[Seed] Migración de esquema detectada como pendiente por UI, delegando a Infrastructure.");
                    }

                    // 2. Poblado de Catálogos Unificado (B.5)
                    // Auditoría forense de ruta de lectura (Diagnostico C-02)
                    string vaultPath = SqliteHelper.DbPathEnc;
                    LoggerHelper.Info($"[AUDIT-VAULT] Iniciando formulario académico (ExecuteSeedData). Vault Lectura: {vaultPath}");

                    LoggerHelper.Info("[FormSesionesAcademicas] Ejecutando Seed Unificado desde Recurso Embebido...");
                    var conn = _sharedConn;
                    var unwrapped = conn is NonDisposableDbConnection nd
                        ? nd.Unwrap()
                        : conn;
                    if (unwrapped is System.Data.SQLite.SQLiteConnection sqliteConn)
                    {
                        ConversorPDF.XLL.Core.Sync.CrSqliteLoader.EnableOnConnection(sqliteConn);
                    }
                    repo.SeedCatalogosTransaccional();
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en ExecuteSeedData (Phase B Optimization)", ex);
            }
        }

        private bool _estaGenerando = false;

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            if (_estaGenerando)
            {
                MessageBox.Show("El proceso de generación está en curso. Por favor espere.", "Proceso Activo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _estaGenerando = true;
            btnGenerar.Enabled = false;
            btnGenerar.Cursor = Cursors.WaitCursor;

            try
            {
                LoggerHelper.Info("[FormSesionesAcademicas] btnGenerar: Iniciando proceso de generación...");

                // D.2 - Validación de coherencia de fechas
                if (dtpFinal.Value.Date < dtpInicio.Value.Date)
                {
                    LoggerHelper.Warning("[FormSesionesAcademicas] btnGenerar: Fecha final es menor a la inicial.");
                    MessageBox.Show("La fecha final no puede ser anterior a la fecha inicial.", "Rango de Fechas Inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 1. Validaciones Iniciales
                if (!double.TryParse(txtTotalHoras.Text, out double totalHorasObjetivo) || totalHorasObjetivo <= 0)
                {
                    LoggerHelper.Warning($"[FormSesionesAcademicas] btnGenerar: Valor de horas totales inválido: {txtTotalHoras.Text}");
                    MessageBox.Show("Por favor, ingrese un número válido de horas totales.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (txtDiasHoras.SelectedDays == null || !txtDiasHoras.SelectedDays.Any())
                {
                    LoggerHelper.Warning("[FormSesionesAcademicas] btnGenerar: No se han seleccionado días/horas.");
                    MessageBox.Show("Debe seleccionar al menos un día y asignar horas.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 2. Aviso de sobrescritura si ya hay datos
                if (dgvSesiones.Rows.Count > 0)
                {
                    MessageBox.Show("Se sobrescribirán las sesiones actuales con los nuevos parámetros.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoggerHelper.Info("[FormSesionesAcademicas] btnGenerar: Se sobrescribirá la grilla (Aceptado por defecto).");
                }

                dgvSesiones.Rows.Clear();
                _sesionesGeneradas.Clear();

                // 3. Preparación de Variables
                DateTime fechaActual = dtpInicio.Value.Date;
                DateTime fechaLimite = dtpFinal.Value.Date;

                int anioActual = DateTime.Now.Year;
                DateTime limiteDiciembre = new DateTime(anioActual, 12, 15);
                if (fechaLimite > limiteDiciembre)
                {
                    var respuesta = MessageBox.Show($"La fecha final supera el 15 de diciembre del año {anioActual}. ¿Desea continuar de todos modos?",
                                                    "Fecha Límite Excedida", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (respuesta != DialogResult.Yes) return;
                }
                double horasAcumuladas = 0;
                var diasInfo = txtDiasHoras.SelectedDays;
                var patronSemanas = ParsearSelectedDays(diasInfo);
                int safetyGuard = 0;
                int cycleIndex = 0;

                string patronLog = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value.Horas}|{k.Value.Horario}"));
                LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Configuraciones leídas. Objetivo: {totalHorasObjetivo}h. Patrón: {patronLog}");

                // Obtener festivos del periodo para validación de PRESENCIAL/FESTIVO
                int idPeriodo = Convert.ToInt32(cmbPeriodos.SelectedValue ?? 0);
                if (idPeriodo <= 0)
                {
                    MessageBox.Show("Debe seleccionar un período académico válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // [Fase 39 Corrección] Validación robusta del SelectedItem
                // En modo DropDown, SelectedValue puede ser nulo u omitir validación si el texto no fue completado.
                // Evaluamos que el SelectedItem sea efectivamente un objeto válido del catálogo.
                if (!(cmbTerceros.SelectedItem is TerceroDisplayDto dtoTercero) || dtoTercero.Id <= 0)
                {
                    MessageBox.Show("Debe seleccionar un docente (Tercero) válido de la lista autocompletada.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int idTercero = dtoTercero.Id;

                // Determinar modalidad ANTES de la consulta de festivos para poder usarla
                // en el lambda de InvokeRepository (las lambdas capturan variables por referencia;
                // una variable declarada después del lambda no está en su ámbito de captura).
                string tipoTutoria = cmbTipoTutoria.SelectedItem?.ToString();
                string modalidad = ModalidadTutoriaCatalog.NormalizarTutoria(tipoTutoria);
                bool esPlataforma = ModalidadTutoriaCatalog.EsPlataforma(modalidad);
                bool esMixta = ModalidadTutoriaCatalog.EsMixta(modalidad);

                string residenciaTercero = "";
                if (cmbTerceros.SelectedItem is TerceroDisplayDto dto)
                {
                    residenciaTercero = dto.Residencia ?? "";
                }

                string sedeDisplay = cmbSedes.SelectedItem is LookupItem li
                    ? li.DisplayName
                    : (cmbSedes.Text ?? "");
                string sedeCiudad = sedeDisplay.Split('(')[0].Trim();

                bool esSedeIgualResidencia = !string.IsNullOrEmpty(residenciaTercero) &&
                                              !string.IsNullOrEmpty(sedeCiudad) &&
                                              NormalizeCityName(residenciaTercero).Equals(NormalizeCityName(sedeCiudad), StringComparison.OrdinalIgnoreCase);

                // ─────────────────────────────────────────────────────────────────────────────
                // Carga de Festivos y Rangos PRESENCIAL
                //
                // Regla de negocio:
                //   • PRESENCIAL: rango de fechas del calendario semestral en que el docente
                //     se desplaza físicamente a la sede. Son GLOBALES a todos los períodos;
                //     no se filtran por idPeriodo porque un mismo calendario semestral puede
                //     compartirse entre varias vinculaciones de distintos períodos.
                //   • FESTIVO:    día feriado específico del período académico en curso.
                //     Se filtran por idPeriodo para respetar el calendario propio de cada cohorte.
                //
                // Fix v1.2.54: El bug anterior filtraba PRESENCIAL por idPeriodo, lo que causaba
                // que las sesiones dentro de rangos PRESENCIAL se generaran con Viatica=false
                // cuando el idPeriodo del combo difería del idPeriodo con que se guardaron los rangos.
                // ─────────────────────────────────────────────────────────────────────────────
                var rangosPresenciales = new List<Festivo>();
                var rangosFestivos = new List<Festivo>();
                bool consultaFestivosExitosa = false;
                string errorFestivos = null;

                InvokeRepository(repo =>
                {
                    var result = repo.GetFestivos();
                    if (result.IsSuccess)
                    {
                        consultaFestivosExitosa = true;
                        var todosLosFestivos = result.Value.ToList();
                        _rangosFestivos = todosLosFestivos;

                        // PRESENCIAL: se cargan de TODOS los períodos (calendario semestral global).
                        rangosPresenciales = todosLosFestivos
                            .Where(f => string.Equals((f.TipoFestivo ?? "").Trim(), "PRESENCIAL", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        // FESTIVO: se filtran por el período académico seleccionado.
                        rangosFestivos = todosLosFestivos
                            .Where(f => string.Equals((f.TipoFestivo ?? "").Trim(), "FESTIVO", StringComparison.OrdinalIgnoreCase)
                                     && f.IdPeriodo == idPeriodo)
                            .ToList();

                        LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Festivos cargados. " +
                                          $"Total BD: {todosLosFestivos.Count}. " +
                                          $"PRESENCIAL (global): {rangosPresenciales.Count}. " +
                                          $"FESTIVO (período {idPeriodo}): {rangosFestivos.Count}. " +
                                          $"Modalidad: {modalidad}.");

                        foreach (var r in rangosPresenciales)
                        {
                            LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Rango PRESENCIAL: " +
                                              $"{r.FechaFestivoInicio:dd/MM/yyyy} → {r.FechaFestivoFinal:dd/MM/yyyy} | {r.DetalleFestivo}");
                        }
                    }
                    else
                    {
                        errorFestivos = result.Error;
                        LoggerHelper.Error($"[FormSesionesAcademicas] Fallo al consultar festivos: {result.Error}");
                    }
                });

                if (!consultaFestivosExitosa)
                {
                    MessageBox.Show($"No se pudo consultar la configuración de festivos. Error: {errorFestivos}",
                                    "Error de Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Advertencia preventiva: vinculación MIXTA sin rangos configurados en BD.
                // El usuario puede continuar, pero todas las sesiones quedarán con Viatica=false.
                if (esMixta && !esSedeIgualResidencia && rangosPresenciales.Count == 0)
                {
                    LoggerHelper.Warning($"[FormSesionesAcademicas] btnGenerar: Modalidad MIXTA pero no hay rangos PRESENCIAL en BD. Viatica=false para todas las sesiones.");
                    var respuesta = MessageBox.Show(
                        "No se encontraron rangos PRESENCIAL configurados en la base de datos.\n\n" +
                        "Con la modalidad Mixta, las sesiones se generarán sin viáticas asignadas (Viática = false para todas).\n\n" +
                        "Para corregirlo: configure rangos tipo PRESENCIAL en el módulo de Festivos antes de generar.\n\n" +
                        "¿Desea continuar de todas formas?",
                        "Sin Rangos Presencial",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (respuesta != DialogResult.Yes)
                    {
                        return;
                    }
                }

                // ─────────────────────────────────────────────────────────────────────────────
                // Bucle de Generación por Horas Exactas
                // Continúa iterando hasta cumplir totalHorasObjetivo o el límite de seguridad.
                // ─────────────────────────────────────────────────────────────────────────────
                int safetyGuardMax = 2000;

                while (horasAcumuladas < totalHorasObjetivo && safetyGuard < safetyGuardMax)
                {
                    // Validación de Filtro Estricto PRESENCIAL y Exclusión FESTIVO
                    bool estaEnRangoPresencial = rangosPresenciales.Any(f =>
                        fechaActual >= f.FechaFestivoInicio.Date && fechaActual <= f.FechaFestivoFinal.Date);

                    bool esFestivo = rangosFestivos.Any(f =>
                        fechaActual >= f.FechaFestivoInicio.Date && fechaActual <= f.FechaFestivoFinal.Date);

                    // Siempre se saltan los festivos, sin importar la modalidad
                    if (esFestivo)
                    {
                        fechaActual = fechaActual.AddDays(1);
                        // Se omite incrementar el safetyGuard aquí para que solo cuente iteraciones productivas
                        continue;
                    }

                    // Determinación de Viáticos desacoplada vía CalculadoraViaticaSesion
                    bool asignarViatica = CalculadoraViaticaSesion.DeterminarViatica(
                        fechaActual,
                        modalidad,
                        sedeCiudad,
                        residenciaTercero,
                        rangosPresenciales.Select(f => (f.FechaFestivoInicio.Date, f.FechaFestivoFinal.Date)));

                    if (patronSemanas.ContainsKey(fechaActual.DayOfWeek))
                    {
                        double horasBase = patronSemanas[fechaActual.DayOfWeek];
                        double horasFaltantes = totalHorasObjetivo - horasAcumuladas;
                        double horasAsignar = Math.Min(horasBase, horasFaltantes);

                        LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Ciclo {cycleIndex}. Fecha: {fechaActual:dd/MM/yyyy}. " +
                                         $"EstaEnPresencial: {estaEnRangoPresencial}. EsPlataforma: {esPlataforma}. EsSedeIgualResidencia: {esSedeIgualResidencia}. " +
                                         $"Faltan: {horasFaltantes}h. Asignando: {horasAsignar}h. Viatica: {asignarViatica}");

                        // Encontrar la información del horario para el día actual
                        string dayNameES = fechaActual.ToString("dddd", new CultureInfo("es-ES")).ToUpper();
                        string horarioSesion = "8:00 AM A 12:00 M"; // Default razonable
                        bool found = false;

                        // Buscar el DaySessionInfo correspondiente normalizando nombres
                        foreach (var kvp in diasInfo)
                        {
                            string nomNormalizado = DayPatternParser.TranslateDay(kvp.Key);
                            string nomFechaNormalizado = DayPatternParser.TranslateDay(dayNameES);
                            if (nomNormalizado == nomFechaNormalizado && !string.IsNullOrEmpty(kvp.Value.Horario))
                            {
                                horarioSesion = kvp.Value.Horario;
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LoggerHelper.Warning($"[FormSesionesAcademicas] No se encontró horario en UI para {dayNameES}. Usando default: {horarioSesion}");
                        }

                        // El bucle solo construye el modelo de dominio. La presentación
                        // visual se delega a DisplaySesiones() al terminar el cálculo,
                        // lo que garantiza un Refresh() atómico y elimina la duplicación
                        // de lógica de pintado que existía previamente en este bucle.
                        var sesion = new Sesion
                        {
                            IdVinculacion = _currentVinculacionId,
                            IdTercero = idTercero,
                            FechaSesion = fechaActual,
                            HorarioSesion = horarioSesion,
                            Horas = horasAsignar,
                            Viatica = asignarViatica
                        };

                        _sesionesGeneradas.Add(sesion);
                        horasAcumuladas += horasAsignar;
                        safetyGuard++; // Incrementamos solo cuando la fecha es válida y corresponde a un día en el patrón
                        cycleIndex++;
                    }
                    fechaActual = fechaActual.AddDays(1);
                }

                // 5. Informe Final y Auto-Ajuste de Fechas
                // Poblar el grid desde _sesionesGeneradas (fuente de verdad del modelo).
                // DisplaySesiones() asigna el Tag, llama a UpdateSelectionSummary() y
                // fuerza un Refresh() para que DataGridViewCheckBoxColumn renderice correctamente.
                DisplaySesiones();
                int sesionesRealesGeneradas = dgvSesiones.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Generación completada. Sesiones: {sesionesRealesGeneradas}. Total: {horasAcumuladas}h.");

                bool generacionCompleta = horasAcumuladas >= totalHorasObjetivo;
                string msgBase;
                string titulo;
                MessageBoxIcon icono;

                if (generacionCompleta)
                {
                    msgBase = $"Generación completada: {horasAcumuladas} horas repartidas en {sesionesRealesGeneradas} sesiones.";
                    titulo = "Generación Exitosa";
                    icono = MessageBoxIcon.Information;
                }
                else
                {
                    msgBase = $"Generación INCOMPLETA: solo se pudieron generar {horasAcumuladas} de {totalHorasObjetivo} horas ({sesionesRealesGeneradas} sesiones).\n\n" +
                              $"Se alcanzó el límite de seguridad de {safetyGuardMax} sesiones. " +
                              $"Verifique que haya suficientes días hábiles dentro del patrón semanal seleccionado.";
                    titulo = "Generación Incompleta";
                    icono = MessageBoxIcon.Warning;
                }

                // Guard clause defensivo ante futuros refactors de DI (actualmente se inicializa en el constructor).
                // El return aborta la generación correctamente y delega en el bloque 'finally' la liberación de la UI.
                if (_dateSyncService == null)
                {
                    LoggerHelper.Error("[FormSesionesAcademicas] _dateSyncService es null. No se puede calcular ajuste.");
                    MessageBox.Show("Error de inicialización del servicio de sincronización.", "Error Interno",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var adjustment = _dateSyncService.CalculateAdjustment(_sesionesGeneradas, dtpInicio.Value, dtpFinal.Value);

                if (adjustment.NeedsAdjustment)
                {
                    // Actualizar controles (Fase 8 / Fase 699: Evitar que el ValueChanged altere los valores prematuramente)
                    dtpInicio.ValueChanged -= DtpInicio_ValueChanged;
                    dtpFinal.ValueChanged -= DtpFinal_ValueChanged;
                    try
                    {
                        dtpFinal.Value = adjustment.CalculatedEnd;
                        dtpInicio.Value = adjustment.CalculatedStart;
                    }
                    finally
                    {
                        dtpInicio.ValueChanged += DtpInicio_ValueChanged;
                        dtpFinal.ValueChanged += DtpFinal_ValueChanged;
                    }

                    string detalles = string.Join("\n", adjustment.Messages);
                    MessageBox.Show(msgBase + "\n\nSe detectaron ajustes automáticos en las fechas:\n" + detalles,
                                    titulo + " - Fechas Auto-Ajustadas", MessageBoxButtons.OK,
                                    adjustment.ExceededOriginalLimit ? MessageBoxIcon.Warning : icono);
                }
                else
                {
                    MessageBox.Show(msgBase + (generacionCompleta ? $"\n\nPeríodo abarcado: {dtpInicio.Value:dd/MM/yyyy} al {dtpFinal.Value:dd/MM/yyyy}." : ""),
                                    titulo, MessageBoxButtons.OK, icono);
                }

                // Guardar el estado validado temporal
                _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en btnGenerar_Click", ex);
                MessageBox.Show($"Error durante la generación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _estaGenerando = false;
                btnGenerar.Enabled = true;
                btnGenerar.Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// Asegura que la lista en memoria de festivos del calendario esté cargada desde BD si aún no se ha obtenido.
        /// </summary>
        private void CargarFestivosSiEsNecesario()
        {
            if (_rangosFestivos == null || _rangosFestivos.Count == 0)
            {
                InvokeRepository(repo =>
                {
                    var result = repo.GetFestivos();
                    if (result.IsSuccess && result.Value != null)
                    {
                        _rangosFestivos = result.Value.ToList();
                    }
                });
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            // C.1 - Corregido: La validación ahora tiene sentido lógico
            if (ParametrosEstanDesincronizados())
            {
                MessageBox.Show(
                    "Has modificado los parámetros maestros (Docente, Horas o Distribución de días).\n\n" +
                    "Debes volver a hacer click en [Generar Sesiones] para sincronizar el calendario antes de Guardar.",
                    "Falta Generar Sesiones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgvSesiones.Rows.Count == 0)
            {
                MessageBox.Show("No hay sesiones generadas para guardar.", "Advertencia",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // F2.3: Asegurar festivos en memoria para validaciones interactivas
            CargarFestivosSiEsNecesario();

            // Extraer sesiones marcadas preliminarmente para validaciones interactivas en UI Thread
            List<Sesion> sesionesParaValidar;
            try
            {
                sesionesParaValidar = PrepararSesionesDesdeGrid().ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en los datos de las sesiones: {ex.Message}", "Datos Inválidos",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (sesionesParaValidar.Count == 0)
            {
                MessageBox.Show("No hay sesiones marcadas para guardar.", "Advertencia",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 690 (plan 666) / F2.3-bis: Validación de rango de fechas del período (Advertencia confirmable)
            var fueraDeRango = sesionesParaValidar
                .Where(s => s.FechaSesion.Date < dtpInicio.Value.Date || s.FechaSesion.Date > dtpFinal.Value.Date)
                .OrderBy(s => s.FechaSesion)
                .ToList();

            bool ampliarRangoConfirmado = false;

            if (fueraDeRango.Any())
            {
                string fechasFuera = string.Join(", ", fueraDeRango
                    .Select(s => s.FechaSesion.ToString("dd/MM/yyyy"))
                    .Distinct());

                LoggerHelper.Warning(
                    $"[Guardado] {fueraDeRango.Count} sesión(es) fuera del rango configurado " +
                    $"[{dtpInicio.Value:dd/MM/yyyy} - {dtpFinal.Value:dd/MM/yyyy}]: {fechasFuera}");

                var rtaRango = MessageBox.Show(
                    $"Existen sesiones con fechas fuera de la ventana de la vinculación ({dtpInicio.Value:dd/MM/yyyy} a {dtpFinal.Value:dd/MM/yyyy}):\n" +
                    $"{fechasFuera}\n\n" +
                    "Si continúa, el rango [Fecha Inicio] - [Fecha Final] se ampliará automáticamente " +
                    "para cubrir estas fechas y se guardará tal como está configurado.\n\n" +
                    "Si selecciona 'No', ajuste las fechas del grid o modifique [Fecha Inicio]/[Fecha Final] antes de guardar.\n\n" +
                    "¿Desea continuar y guardar de todas formas?",
                    "Fechas Fuera de la Ventana",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (rtaRango != DialogResult.Yes)
                {
                    LoggerHelper.Info("[Guardado] Cancelado por el usuario: fechas fuera de la ventana de la vinculación.");
                    return;
                }

                ampliarRangoConfirmado = true;
            }

            // F2.3: Advertencia interactiva de festivos (solo para sesiones marcadas a guardar)
            int idPeriodoSel = Convert.ToInt32(cmbPeriodos.SelectedValue ?? 0);
            var enFestivo = sesionesParaValidar
                .Where(s => _rangosFestivos != null && _rangosFestivos.Any(f =>
                    string.Equals((f.TipoFestivo ?? "").Trim(), "FESTIVO", StringComparison.OrdinalIgnoreCase) &&
                    (f.IdPeriodo == idPeriodoSel || f.IdPeriodo <= 0) &&
                    s.FechaSesion.Date >= f.FechaFestivoInicio.Date &&
                    s.FechaSesion.Date <= f.FechaFestivoFinal.Date))
                .ToList();

            if (enFestivo.Any())
            {
                string fechasFestivas = string.Join(", ", enFestivo.Select(s => s.FechaSesion.ToString("dd/MM/yyyy")).Distinct());
                var rtaFestivo = MessageBox.Show(
                    $"Las siguientes sesiones coinciden con días festivos o no laborales del período:\n{fechasFestivas}\n\n¿Desea continuar y guardarlas de todas formas?",
                    "Sesiones en Días Festivos",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (rtaFestivo != DialogResult.Yes)
                {
                    return;
                }
            }

            // F2.5: Advertencia interactiva si existen sesiones con más de 8 horas diarias
            var jornadasLargas = sesionesParaValidar.Where(s => s.Horas > 8).ToList();
            if (jornadasLargas.Any())
            {
                string filasInfo = string.Join(", ", jornadasLargas.Select(s => $"{s.FechaSesion:dd/MM/yyyy} ({s.Horas}h)"));
                var rtaHoras = MessageBox.Show(
                    $"Las siguientes sesiones superan la jornada estándar de 8 horas diarias:\n{filasInfo}\n\n¿Desea continuar con estos valores?",
                    "Jornada Superior a 8 Horas",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (rtaHoras != DialogResult.Yes)
                {
                    return;
                }
            }

            // F2.4: Preguntar al usuario si la suma de horas del grid difiere de txtTotalHoras
            double sumaGrid = sesionesParaValidar.Sum(s => s.Horas);
            int objetivoHoras = int.TryParse(txtTotalHoras.Text, out int th) ? th : 0;
            if (objetivoHoras > 0 && Math.Abs(sumaGrid - objetivoHoras) > 0.0001)
            {
                var rtaSuma = MessageBox.Show(
                    $"La suma de horas de las sesiones seleccionadas ({sumaGrid}h) no coincide con el total de horas de la vinculación ({objetivoHoras}h).\n\n" +
                    $"¿Desea actualizar el total de horas de la vinculación a {sumaGrid}h y continuar guardando?\n\n" +
                    "(Seleccione 'Sí' para actualizar el total, 'No' para mantener {objetivoHoras}h, o 'Cancelar' para revisar)",
                    "Discrepancia en Total de Horas",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (rtaSuma == DialogResult.Cancel)
                {
                    return;
                }
                else if (rtaSuma == DialogResult.Yes)
                {
                    txtTotalHoras.Text = ((int)Math.Round(sumaGrid)).ToString();
                }
            }

            // Fase 691 (plan 666.b): Si el usuario aceptó fechas fuera de rango, ampliar la ventana maestra ANTES de persistir
            if (ampliarRangoConfirmado && sesionesParaValidar.Any())
            {
                var minFecha = sesionesParaValidar.Min(s => s.FechaSesion).Date;
                var maxFecha = sesionesParaValidar.Max(s => s.FechaSesion).Date;

                dtpInicio.ValueChanged -= DtpInicio_ValueChanged;
                dtpFinal.ValueChanged -= DtpFinal_ValueChanged;
                try
                {
                    if (dtpInicio.Value.Date > minFecha) dtpInicio.Value = minFecha;
                    if (dtpFinal.Value.Date < maxFecha) dtpFinal.Value = maxFecha;
                }
                finally
                {
                    dtpInicio.ValueChanged += DtpInicio_ValueChanged;
                    dtpFinal.ValueChanged += DtpFinal_ValueChanged;
                }

                SincronizarFechasEnEstadoGenerado();   // Fase 704: mantener el snapshot coherente aunque el guardado falle o se cancele

                LoggerHelper.Info($"[Guardado] Rango maestro ampliado por confirmación: {dtpInicio.Value:dd/MM/yyyy} → {dtpFinal.Value:dd/MM/yyyy}");
            }

            // Fase 705: consolidar el snapshot con los valores que efectivamente se van a persistir
            // (F2.4 pudo actualizar TotalHoras y F691 pudo ampliar FI/FF). Evita falsos
            // "Falta Generar Sesiones" si la persistencia falla o se cancela.
            _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();

            _guardadoCts = new CancellationTokenSource();

            using (var dlg = new ConversorPDF.XLL.UI.ProgressDialog("Guardando Cambios..."))
            {
                try
                {
                    btnGuardar.Enabled = false;
                    btnGuardar.BackColor = Color.Gray;
                    this.Cursor = Cursors.WaitCursor;

                    dlg.Shown += async (s, ev) =>
                    {
                        try
                        {
                            await GuardarAsync(dlg, _guardadoCts.Token);

                            this.Invoke((Action)(() =>
                            {
                                MessageBox.Show("Sesiones académicas guardadas exitosamente.", "Éxito",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                                this.DatosGuardados = true; // Activar flag de refresco para el padre
                                _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();
                                CargarEstadoPrevio(); // Refrescar vista local
                                dlg.Close();
                            }));
                        }
                        catch (OperationCanceledException)
                        {
                            LoggerHelper.Info("[Guardado] Operación cancelada por el usuario.");
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Error("[Guardado] Error crítico durante guardado", ex);
                            this.Invoke((Action)(() =>
                            {
                                MessageBox.Show($"Error al guardar: {ex.Message}", "Fallo de Persistencia",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                dlg.Close();
                            }));
                        }
                        finally
                        {
                            this.Invoke((Action)(() =>
                            {
                                btnGuardar.Enabled = true;
                                btnGuardar.BackColor = Color.FromArgb(0, 120, 212);
                                this.Cursor = Cursors.Default;
                            }));
                        }
                    };

                    dlg.ShowDialog(this);
                }
                finally
                {
                    // Fase 12: Asegurar la correcta limpieza y liberación del CancellationToken 
                    // independientemente de si el evento Shown de ProgressDialog llegó a dispararse
                    _guardadoCts?.Cancel();
                    _guardadoCts?.Dispose();
                    _guardadoCts = null;
                }
            }
        }

        private async Task GuardarAsync(ProgressDialog dlg, CancellationToken ct)
        {
            var sesiones = PrepararSesionesDesdeGrid().ToList();

            // FASE 241 / F1.14, F2.2, F2.5: Validaciones de dominio previas a la persistencia
            if (sesiones.Any())
            {
                var repetidas = sesiones.GroupBy(s => s.FechaSesion.Date)
                                        .Where(g => g.Count() > 1)
                                        .Select(g => g.Key)
                                        .ToList();

                if (repetidas.Any())
                {
                    throw new InvalidOperationException(
                        $"Existen fechas duplicadas en las sesiones de la grilla: {string.Join(", ", repetidas.Select(d => d.ToString("dd/MM/yyyy")))}. Cada sesión debe tener una fecha única.");
                }

                // F2.2: Detallar números de fila específicos sin horario
                var sinHorario = sesiones.Select((s, i) => new { Sesion = s, Fila = i + 1 })
                                         .Where(x => string.IsNullOrWhiteSpace(x.Sesion.HorarioSesion))
                                         .Select(x => x.Fila)
                                         .ToList();

                if (sinHorario.Any())
                {
                    throw new InvalidOperationException(
                        $"Sesiones sin horario en las filas: {string.Join(", ", sinHorario)}. Corrija la columna [Horario Sesión].");
                }

                // F2.5: Techo de jornada por sesión
                if (sesiones.Any(s => s.Horas <= 0 || s.Horas > 24))
                {
                    throw new InvalidOperationException("Existen sesiones en la grilla con horas no válidas (deben ser mayores a cero y menores o iguales a 24).");
                }
            }

            var vinculoActual = PrepararVinculacionDesdeForm();

            string tempDbPath = string.Empty;

            // FASE 1 y 2: Guardado Atómico usando Contexto Compartido
            await Task.Run(() =>
            {
                // Usamos la conexión ya abierta (Shared Context)
                if (_sharedConn == null || _sharedConn.State != ConnectionState.Open)
                    throw new InvalidOperationException("Contexto de base de datos no disponible para guardado.");

                // Transacción real a nivel de base de datos
                try
                {
                    // Fase 10 FIX / Fase 13 FIX: Esperar a que no haya PUSH activo para evitar 'database is locked' ANTES de iniciar la transacción
                    var scheduler = ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler;
                    if (scheduler != null && scheduler.IsSyncing)
                    {
                        dlg.Progress.Report("Esperando a que finalice la sincronización en segundo plano...");
                        int waited = 0;
                        while (scheduler.IsSyncing && waited < 30000) // 30s max
                        {
                            System.Threading.Thread.Sleep(200);
                            waited += 200;
                        }
                    }

                    using (var transaction = _sharedConn.BeginTransaction())
                    {
                        try
                        {
                            var repo = new AcademicRepository(_sharedConn, transaction, suppressMarkDirty: true, modo: _modo);

                            ct.ThrowIfCancellationRequested();
                            // dlg.Progress.Report("Sincronizando catálogos (v1.2.50.0)...");
                            // repo.SeedCatalogosTransaccional(); // FIX-1.2: Innecesario en cada guardado

                            ct.ThrowIfCancellationRequested();
                            dlg.Progress.Report($"Guardando vinculación y {sesiones.Count} sesiones...");

                            int newVincId = vinculoActual.Id;

                            if (vinculoActual.Id > 0)
                            {
                                var resUpd = repo.UpdateVinculacion(vinculoActual);
                                if (resUpd.IsFailure)
                                    throw new Exception($"Error al actualizar vinculación: {resUpd.Error}");

                                var resDel = repo.DeleteSesionesByVinculacion(vinculoActual.Id);
                                if (resDel.IsFailure)
                                    throw new Exception($"Error al eliminar sesiones previas: {resDel.Error}");
                            }
                            else
                            {
                                var resVinc = repo.InsertVinculacion(vinculoActual);
                                if (resVinc.IsFailure || resVinc.Value <= 0)
                                    throw new Exception($"Error al insertar vinculación: {resVinc.Error ?? "Id generado inválido"}");

                                newVincId = resVinc.Value;
                                foreach (var s in sesiones) s.IdVinculacion = newVincId;
                            }

                            var resSesiones = repo.InsertSesiones(sesiones);
                            if (resSesiones.IsFailure)
                                throw new Exception($"Error al insertar sesiones: {resSesiones.Error}");

                            // Hook H1 (Fase 99/109): Upsert en pool docente-asignatura dentro de la transacción (sólo en producción)
                            if (!_modo.EsModoAg)
                            {
                                try
                                {
                                    var tasRepoTx = new TerceroAsignaturaRepository(_sharedConn, transaction, suppressMarkDirty: true);
                                    int conteo = tasRepoTx.ContarAsignaciones(vinculoActual.IdTercero, vinculoActual.IdPrograma, vinculoActual.IdAsignatura, vinculoActual.Semestre);
                                    var tasEntity = TerceroAsignaturaMapper.FromVinculacion(vinculoActual, conteo);
                                    tasRepoTx.Upsert(tasEntity, transaction);
                                }
                                catch (Exception exH1)
                                {
                                    LoggerHelper.Warning($"[FormSesionesAcademicas] Hook H1 (pool docente): {exH1.Message}");
                                }
                            }

                            ct.ThrowIfCancellationRequested();
                            dlg.Progress.Report("Confirmando transacción...");
                            transaction.Commit();

                            // Fase 14: Actualizar _currentVinculacionId solo después del Commit exitoso
                            if (vinculoActual.Id <= 0)
                            {
                                vinculoActual.Id = newVincId;
                                _currentVinculacionId = newVincId;
                            }

                            ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();

                            try
                            {
                                var poolRepo = new TerceroAsignaturaRepository(_connectionFactory);
                                var syncService = new TerceroAsignaturaSyncService(_connectionFactory, poolRepo);
                                syncService.SincronizarDocenteAsignaturaAsync(
                                    vinculoActual.IdTercero,
                                    vinculoActual.IdPrograma,
                                    vinculoActual.IdAsignatura,
                                    vinculoActual.Semestre,
                                    vinculoActual.Subunidad,
                                    ct).GetAwaiter().GetResult();
                            }
                            catch (Exception exSync)
                            {
                                LoggerHelper.Warning($"[FormSesionesAcademicas] Error no fatal al sincronizar pool docente-asignatura: {exSync.Message}");
                            }

                            LoggerHelper.Info($"[PERF] Transacción v1.2.50.3 completada. VincID={vinculoActual.Id}");
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("Transaction"))
                {
                    LoggerHelper.Error("[GuardarAsync] Transacción huérfana detectada", ex);
                    throw new InvalidOperationException("Se detectó una transacción activa bloqueando la base de datos. Por favor cierre y vuelva a abrir el formulario.", ex);
                }

                // Nota: La persistencia al Vault ahora la gestiona el SyncScheduler de forma asíncrona.
            });

            // FASE 240 / F1.13: Sincronización post-commit de selectores dtpInicio y dtpFinal
            // Se ejecuta únicamente tras la confirmación exitosa de la transacción en base de datos.
            if (sesiones.Any())
            {
                var minFecha = sesiones.Min(s => s.FechaSesion).Date;
                var maxFecha = sesiones.Max(s => s.FechaSesion).Date;

                dtpInicio.ValueChanged -= DtpInicio_ValueChanged;
                dtpFinal.ValueChanged -= DtpFinal_ValueChanged;
                try
                {
                    if (dtpInicio.Value.Date > minFecha) dtpInicio.Value = minFecha;
                    if (dtpFinal.Value.Date < maxFecha) dtpFinal.Value = maxFecha;
                }
                finally
                {
                    dtpInicio.ValueChanged += DtpInicio_ValueChanged;
                    dtpFinal.ValueChanged += DtpFinal_ValueChanged;
                }

                LoggerHelper.Info($"[Guardado] Rango maestro de fechas sincronizado post-commit: {dtpInicio.Value:yyyy-MM-dd} → {dtpFinal.Value:yyyy-MM-dd}");
            }
        }

        private Dictionary<DayOfWeek, double> ParsearSelectedDays(Dictionary<string, DaySessionInfo> input)
        {
            var tmp = input.ToDictionary(k => k.Key, v => (int)v.Value.Horas);
            return DayPatternParser.Parse(tmp);
        }

        /// <summary>
        /// Escanea los controles visuales críticos que afectan matemáticamente el
        /// algoritmo generador de sesiones.
        /// </summary>
        private string ObtenerEstadoParametrosCriticos()
        {
            string diasHorasStr = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value.Horas}|{k.Value.Horario}"));
            int idTercero = Convert.ToInt32(cmbTerceros.SelectedValue ?? 0);
            string totalHoras = txtTotalHoras.Text.Trim();
            string fi = dtpInicio.Value.Date.ToString("yyyyMMdd");
            string ff = dtpFinal.Value.Date.ToString("yyyyMMdd");

            return $"T={idTercero};H={totalHoras};D={diasHorasStr};FI={fi};FF={ff}";
        }

        /// <summary>
        /// Fase 699 / Fase 701: obtiene [primera, última] fecha de sesión del grid para filas marcadas.
        /// </summary>
        private bool TryObtenerRangoSesionesGrid(out DateTime primera, out DateTime ultima)
        {
            primera = DateTime.MinValue;
            ultima = DateTime.MinValue;
            bool hay = false;

            foreach (DataGridViewRow row in dgvSesiones.Rows)
            {
                if (row.IsNewRow) continue;
                if (!SafeCellBool(row.Cells["ColChecked"].Value)) continue;   // Fase 701: alinear con el criterio de guardado
                if (!GridCellValueParser.TryParseFecha(row.Cells["ColFechaSesion"].Value, out DateTime f)) continue;

                if (!hay) { primera = ultima = f.Date; hay = true; }
                else
                {
                    if (f.Date < primera) primera = f.Date;
                    if (f.Date > ultima) ultima = f.Date;
                }
            }

            return hay;
        }

        /// <summary>
        /// Fase 699: amplía silenciosamente [Fecha Inicio, Fecha Final] para envolver las sesiones del grid
        /// (aplica "comprueba la última fecha" y su análogo en la primera).
        /// </summary>
        private void AmpliarVentanaMaestraAlGrid()
        {
            if (_ajustandoFechaMaestra) return;
            if (!TryObtenerRangoSesionesGrid(out DateTime primera, out DateTime ultima)) return;

            bool ampliada = false;
            _ajustandoFechaMaestra = true;
            dtpInicio.ValueChanged -= DtpInicio_ValueChanged;
            dtpFinal.ValueChanged -= DtpFinal_ValueChanged;
            try
            {
                // Fase 700: acotar contra MinDate y MaxDate del control para evitar ArgumentOutOfRangeException
                if (primera < dtpInicio.MinDate) primera = dtpInicio.MinDate;
                if (ultima > dtpFinal.MaxDate) ultima = dtpFinal.MaxDate;

                if (dtpInicio.Value.Date > primera) { dtpInicio.Value = primera; ampliada = true; }   // análogo: primera fecha
                if (dtpFinal.Value.Date < ultima) { dtpFinal.Value = ultima; ampliada = true; }   // "comprueba la última fecha"
            }
            finally
            {
                dtpInicio.ValueChanged += DtpInicio_ValueChanged;
                dtpFinal.ValueChanged += DtpFinal_ValueChanged;
                _ajustandoFechaMaestra = false;
            }

            if (ampliada)
            {
                LoggerHelper.Info(
                    $"[Sincronizacion] Ventana maestra auto-ampliada al rango de sesiones: {dtpInicio.Value:dd/MM/yyyy} → {dtpFinal.Value:dd/MM/yyyy}.");
            }
        }

        /// <summary>
        /// Fase 699: refresca únicamente los segmentos FI/FF del snapshot, preservando T/H/D
        /// para no enmascarar cambios reales de generación. No-op si no existe snapshot.
        /// </summary>
        private void SincronizarFechasEnEstadoGenerado()
        {
            if (string.IsNullOrEmpty(_ultimoEstadoGenerado)) return;

            string fi = dtpInicio.Value.Date.ToString("yyyyMMdd");
            string ff = dtpFinal.Value.Date.ToString("yyyyMMdd");

            var partes = _ultimoEstadoGenerado.Split(';');
            for (int i = 0; i < partes.Length; i++)
            {
                if (partes[i].StartsWith("FI=")) partes[i] = "FI=" + fi;
                else if (partes[i].StartsWith("FF=")) partes[i] = "FF=" + ff;
            }

            string nuevo = string.Join(";", partes);
            if (nuevo == _ultimoEstadoGenerado) return;   // Fase 703: evita logs duplicados si no hubo cambio real

            _ultimoEstadoGenerado = nuevo;
            LoggerHelper.Info($"[Sincronizacion] Snapshot de fechas actualizado por edición manual: FI={fi}, FF={ff}.");
        }

        /// <summary>
        /// Comprueba si el estado actual de los parámetros coincide con el último 
        /// estado usado para generar el grid.
        /// </summary>
        private bool ParametrosEstanDesincronizados()
        {
            if (string.IsNullOrEmpty(_ultimoEstadoGenerado)) return true;

            string estadoActual = ObtenerEstadoParametrosCriticos();
            return estadoActual != _ultimoEstadoGenerado;
        }

        private Vinculacion PrepararVinculacionDesdeForm()
        {
            int idTercero = Convert.ToInt32(cmbTerceros.SelectedValue ?? 0);

            // FASE 27: Por defecto idTerminal = 1 (NO APLICA) y gasto = 0 cuando no es aéreo
            double gastoTerminal = 0;
            int idTerminal = 1;
            int isTerminalAereo = chkTerminalAereo.Checked ? 1 : 0;

            if (chkTerminalAereo.Checked)
            {
                idTerminal = cmbTerminalesAereas.SelectedValue is IConvertible cv
                    ? cv.ToInt32(null)
                    : (int.TryParse(cmbTerminalesAereas.SelectedValue?.ToString(), out int parsed) ? parsed : 0);

                if (idTerminal <= 1)
                {
                    cmbTerminalesAereas.BackColor = Color.LightYellow;
                    cmbTerminalesAereas.Focus();
                    throw new InvalidOperationException("Debe seleccionar una terminal aérea.");
                }
                cmbTerminalesAereas.BackColor = SystemColors.Window;

                if (!double.TryParse(txtGastoTerminalAerea.Text, NumberStyles.Any, CultureInfo.CurrentCulture, out gastoTerminal))
                    throw new InvalidOperationException("El gasto de terminal aérea no es un valor numérico válido.");

                if (gastoTerminal <= 0)
                    throw new InvalidOperationException("El gasto de terminal aérea debe ser mayor a cero.");

                if (cmbTerminalesAereas.SelectedItem is Terminal terminal && gastoTerminal > terminal.MaxValor)
                    throw new InvalidOperationException($"El gasto de terminal aérea ({gastoTerminal:C0}) supera el máximo permitido ({terminal.MaxValor:C0}).");
            }
            else
            {
                cmbTerminalesAereas.BackColor = SystemColors.Window;
                idTerminal = 1; // NO APLICA
                gastoTerminal = 0;
            }

            // FASE 1: Validación Defensiva de Tercero (Cross-Validation)
            var vinculacion = new Vinculacion
            {
                Id = _currentVinculacionId,
                IdPeriodo = Convert.ToInt32(cmbPeriodos.SelectedValue ?? 0),
                IdSede = Convert.ToInt32(cmbSedes.SelectedValue ?? 0),
                IdTercero = idTercero,
                IdPrograma = Convert.ToInt32(cmbProgramas.SelectedValue ?? 0),
                // Calidad: Capturamos el valor para evitar NRE si SelectedItem es null durante el binding
                Grupo = (cmbGrupos.SelectedItem?.ToString() == null || cmbGrupos.SelectedItem.ToString() == "--") ? null : cmbGrupos.SelectedItem.ToString(),
                Semestre = cmbSemestres.SelectedItem?.ToString() ?? "I",
                IdAsignatura = Convert.ToInt32(cmbAsignaturas.SelectedValue ?? 0),
                Modalidad = ModalidadTutoriaCatalog.NormalizarTutoria(cmbTipoTutoria.SelectedItem?.ToString()),
                TotalHoras = int.TryParse(txtTotalHoras.Text, out int th) ? th : 0,
                DiaHorasSemanales = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value.Horas}|{k.Value.Horario}")),
                FechaInicio = dtpInicio.Value.Date,
                FechaFin = dtpFinal.Value.Date,
                IsPernoctaUltimoDia = chkPernoctaUltimoDia.Checked ? 1 : 0,
                IsTerminalAereo = isTerminalAereo,
                IdTerminales = idTerminal,
                GastoTerminalAerea = gastoTerminal
            };

            // Validar que el objeto esté completo
            if (vinculacion.IdPeriodo == 0 || vinculacion.IdTercero == 0 || vinculacion.IdSede == 0 || vinculacion.IdAsignatura == 0)
            {
                throw new InvalidOperationException("Faltan parámetros maestros por seleccionar (Período, Tercero, Sede o Asignatura).");
            }

            return vinculacion;
        }

        private void chkTerminalAereo_CheckedChanged(object sender, EventArgs e)
        {
            pnlTransporteAereo.Visible = chkTerminalAereo.Checked;
            if (!chkTerminalAereo.Checked)
            {
                cmbTerminalesAereas.SelectedIndex = -1;
                txtGastoTerminalAerea.Text = string.Empty;
                lblMaxValor.Text = string.Empty;
            }
        }

        private void cmbTerminalesAereas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTerminalesAereas.SelectedItem is Terminal terminal)
            {
                lblMaxValor.Text = $"(Máximo: {terminal.MaxValor.ToString("C0", CultureInfo.CurrentCulture)})";
                txtGastoTerminalAerea.Text = terminal.MaxValor.ToString(CultureInfo.CurrentCulture);
            }
            else
            {
                lblMaxValor.Text = string.Empty;
                txtGastoTerminalAerea.Text = string.Empty;
            }
        }

        private static bool SafeCellBool(object value)
        {
            if (value == null || value == DBNull.Value) return false;
            if (value is bool b) return b;
            if (value is int i) return i != 0;
            if (value is string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return false;
                if (bool.TryParse(s, out bool parsed)) return parsed;
                if (int.TryParse(s, out int intVal)) return intVal != 0;
                return false;
            }
            try { return Convert.ToBoolean(value); }
            catch { return false; }
        }

        /// <summary>
        /// Prepara y extrae la lista de sesiones seleccionadas desde el DataGridView,
        /// validando el contenido de las filas marcadas y aplicando reparación defensiva
        /// desde el modelo en <see cref="DataGridViewRow.Tag"/> si es necesario.
        /// (Fase 235 / F1.1 - F1.9, Fase 2 / F2.1, F2.8).
        /// </summary>
        /// <returns>Lista de sesiones validadas y listas para persistir.</returns>
        private List<Sesion> PrepararSesionesDesdeGrid()
        {
            // FASE UI: Forzar commit global de manera segura
            if (dgvSesiones.IsCurrentCellDirty) dgvSesiones.CommitEdit(DataGridViewDataErrorContexts.Commit);
            dgvSesiones.EndEdit();

            // Construir DTOs abstractos desde WinForms
            var rows = new List<GridRowDto>();
            foreach (DataGridViewRow row in dgvSesiones.Rows)
            {
                if (row.IsNewRow) continue;

                bool isChecked = SafeCellBool(row.Cells["ColChecked"].Value);

                // F2.1: No validar filas desmarcadas/excluidas (evita bloquear datos legacy o ediciones canceladas)
                if (!isChecked)
                {
                    var tagSkip = row.Tag as Sesion;
                    rows.Add(new GridRowDto
                    {
                        IsChecked = false,
                        IsNewRow = false,
                        Tag = tagSkip,
                        Horas = tagSkip?.Horas ?? 0,
                        FechaSesion = tagSkip?.FechaSesion ?? DateTime.MinValue,
                        HorarioSesion = tagSkip?.HorarioSesion ?? string.Empty,
                        Viatica = SafeCellBool(row.Cells["ColViatica"].Value)
                    });
                    continue;
                }

                // F2.1: Para filas marcadas, parsear con reparación defensiva desde Tag en caso de texto inválido
                if (!GridCellValueParser.TryParseHoras(row.Cells["ColHoras"].Value, out double horas))
                {
                    var tagRepair = row.Tag as Sesion;
                    if (tagRepair != null && tagRepair.Horas > 0 && tagRepair.Horas <= 24)
                    {
                        horas = tagRepair.Horas;
                        row.Cells["ColHoras"].Value = horas;
                        LoggerHelper.Warning($"[Grid] Fila {row.Index + 1}: horas ilegibles; restauradas desde modelo ({horas}).");
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Fila {row.Index + 1}: Cantidad de horas inválida '{row.Cells["ColHoras"].Value}'. Ingrese un número mayor a cero y menor o igual a 24.");
                    }
                }

                string horario = row.Cells["ColHorario"].Value?.ToString()?.Trim() ?? string.Empty;

                if (!GridCellValueParser.TryParseFecha(row.Cells["ColFechaSesion"].Value, out DateTime fechaSesion))
                {
                    var tagRepair = row.Tag as Sesion;
                    if (tagRepair != null && tagRepair.FechaSesion != default(DateTime))
                    {
                        fechaSesion = tagRepair.FechaSesion;
                        row.Cells["ColFechaSesion"].Value = fechaSesion.ToString("dd/MM/yyyy");
                        LoggerHelper.Warning($"[Grid] Fila {row.Index + 1}: fecha ilegible; restaurada desde modelo ({fechaSesion:dd/MM/yyyy}).");
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Fila {row.Index + 1}: Formato de fecha inválido '{row.Cells["ColFechaSesion"].Value}'. Use el formato dd/MM/yyyy.");
                    }
                }

                if (row.Tag is Sesion tagModel)
                {
                    tagModel.Horas = horas;
                    tagModel.FechaSesion = fechaSesion;
                    tagModel.HorarioSesion = horario;
                    tagModel.Viatica = SafeCellBool(row.Cells["ColViatica"].Value);
                }

                rows.Add(new GridRowDto
                {
                    IsChecked = true,
                    IsNewRow = false,
                    Tag = row.Tag,
                    Horas = horas,
                    FechaSesion = fechaSesion,
                    HorarioSesion = horario,
                    // F2.8: Viatica fiel al estado de la celda post sincronización CellValueChanged
                    Viatica = SafeCellBool(row.Cells["ColViatica"].Value)
                });
            }

            // FASE 238 / F1.9: Liberar selección activa después de construir DTOs para no cancelar ediciones pendientes
            dgvSesiones.CurrentCell = null;

            int idTerceroMaster = Convert.ToInt32(cmbTerceros.SelectedValue ?? 0);

            // DELEGAR Lógica Pura al Mapper (Permite testeo unitario sin WinForms)
            return GridToSesionMapper.Extract(rows, idTerceroMaster, _currentVinculacionId);
        }
        /// <summary>
        /// Restaura los controles del formulario a partir del último estado persistido
        /// en la Vinculación Id=1 y sus Sesiones asociadas.
        /// </summary>
        private void CargarEstadoPrevio()
        {
            try
            {
                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Buscando estado guardado para VincID={_currentVinculacionId}...");
                int idVinculacion = _currentVinculacionId;

                InvokeRepository(repo =>
                {
                    var vincResult = repo.GetVinculacionById(idVinculacion);
                    if (vincResult.IsFailure)
                    {
                        LoggerHelper.Info("[FormSesionesAcademicas] CargarEstadoPrevio: No existe vinculación previa. Usando valores por defecto.");
                        try
                        {
                            if (IsHandleCreated)
                            {
                                this.Invoke((Action)(() =>
                                {
                                    chkTerminalAereo.Checked = false;
                                    chkPernoctaUltimoDia.Checked = false;
                                    pnlTransporteAereo.Visible = false;
                                }));
                            }

                            if (_sharedConn != null && _sharedConn.State == ConnectionState.Open)
                            {
                                var targetIdObj = _sharedConn.ExecuteScalar("SELECT id_periodo FROM config WHERE id = 1");
                                if (targetIdObj != null && targetIdObj != DBNull.Value)
                                {
                                    int defaultIdPeriodo = Convert.ToInt32(targetIdObj);
                                    if (defaultIdPeriodo > 0)
                                    {
                                        this.Invoke((Action)(() =>
                                        {
                                            if (cmbPeriodos.DataSource != null)
                                            {
                                                cmbPeriodos.SelectedValue = defaultIdPeriodo;
                                                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Período por defecto asignado desde config: {defaultIdPeriodo}");
                                            }
                                        }));
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[FormSesionesAcademicas] Error al cargar valores por defecto de config: {ex.Message}");
                        }
                        return;
                    }

                    var vinc = vincResult.Value;
                    string fechaInicioStr = vinc.FechaInicio.HasValue
                        ? vinc.FechaInicio.Value.ToString("yyyy-MM-dd HH:mm:ss")
                        : "NULL";
                    string fechaFinStr = vinc.FechaFin.HasValue
                        ? vinc.FechaFin.Value.ToString("yyyy-MM-dd HH:mm:ss")
                        : "NULL";

                    LoggerHelper.Info($"[DIAGNOSTICO DB] Valores recuperados para vinculación ID {idVinculacion}:\r\n" +
                                      $"  - Id: {vinc.Id}\r\n" +
                                      $"  - IdPeriodo: {vinc.IdPeriodo}\r\n" +
                                      $"  - IdSede: {vinc.IdSede}\r\n" +
                                      $"  - IdTercero: {vinc.IdTercero}\r\n" +
                                      $"  - IdPrograma: {vinc.IdPrograma}\r\n" +
                                      $"  - Semestre: '{vinc.Semestre}'\r\n" +
                                      $"  - IdAsignatura: {vinc.IdAsignatura}\r\n" +
                                      $"  - Grupo: '{vinc.Grupo}'\r\n" +
                                      $"  - Modalidad: '{vinc.Modalidad}'\r\n" +
                                      $"  - TotalHoras: {vinc.TotalHoras}\r\n" +
                                      $"  - DiaHorasSemanales: '{vinc.DiaHorasSemanales}'\r\n" +
                                      $"  - FechaInicio: {fechaInicioStr}\r\n" +
                                      $"  - FechaFin: {fechaFinStr}");

                    // ── Restaurar combos de Vinculación ──
                    this.Invoke((Action)(() =>
                    {
                        if (vinc.IdPeriodo > 0)
                        {
                            // Fase 19: Verificar que DataSource no sea nulo antes de asignar el SelectedValue
                            if (cmbPeriodos.DataSource != null)
                                cmbPeriodos.SelectedValue = vinc.IdPeriodo;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Período restaurado Id={vinc.IdPeriodo}");
                        }

                        if (vinc.IdTercero > 0)
                        {
                            var tercerosDto = cmbTerceros.DataSource as List<TerceroDisplayDto>;
                            var terceroExists = tercerosDto != null && tercerosDto.Any(t => t.Id == vinc.IdTercero);
                            if (terceroExists)
                            {
                                // Fase 18: Verificar que DataSource no sea nulo antes de asignar el SelectedValue
                                if (cmbTerceros.DataSource != null)
                                    cmbTerceros.SelectedValue = vinc.IdTercero;

                                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Tercero restaurado Id={vinc.IdTercero}");
                            }
                            else
                            {
                                if (tercerosDto != null)
                                {
                                    try
                                    {
                                        // Fase 14: Protección contra ObjectDisposedException
                                        if (_sharedConn == null) throw new ObjectDisposedException(nameof(_sharedConn));

                                        // Fase 9: Verificar el estado de la conexión antes de consultar (para evitar InvalidOperationException en Dapper)
                                        if (_sharedConn.State != System.Data.ConnectionState.Open)
                                        {
                                            if (_sharedConn.State == System.Data.ConnectionState.Broken)
                                                _sharedConn.Close();
                                            _sharedConn.Open();
                                        }

                                        var t = _sharedConn.QueryFirstOrDefault<Tercero>("SELECT id as Id, nombre as Nombre, residencia as Residencia FROM terceros WHERE id = @id", new { id = vinc.IdTercero });
                                        if (t != null)
                                        {
                                            tercerosDto.Add(new TerceroDisplayDto { Id = t.Id, Nombre = t.Nombre + " [Inactivo/Filtrado]", Residencia = t.Residencia });

                                            // Fase 10: Asignar DisplayMember y ValueMember antes de asignar el DataSource
                                            cmbTerceros.DataSource = null;
                                            cmbTerceros.DisplayMember = "DisplayName";
                                            cmbTerceros.ValueMember = "Id";
                                            cmbTerceros.DataSource = tercerosDto;

                                            cmbTerceros.SelectedValue = vinc.IdTercero;
                                            LoggerHelper.Warning($"[FormSesionesAcademicas] CargarEstadoPrevio: Tercero ID={vinc.IdTercero} inactivo o filtrado. Agregado temporalmente.");
                                        }
                                        else
                                        {
                                            LoggerHelper.Warning($"[CargarEstadoPrevio] Tercero ID={vinc.IdTercero} no disponible tras filtro y no hallado en BD.");
                                            // Fase 11: Notificar al usuario si el tercero crítico no existe en BD
                                            MessageBox.Show($"El docente asignado (ID: {vinc.IdTercero}) no pudo ser recuperado de la base de datos. Verifique su existencia.", "Advertencia de Datos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                    catch (ObjectDisposedException)
                                    {
                                        LoggerHelper.Error("[CargarEstadoPrevio] _sharedConn dispuesto.");
                                        MessageBox.Show("La conexión fue cerrada inesperadamente. No se pudo recuperar el docente.", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    catch (Exception ex)
                                    {
                                        LoggerHelper.Error($"[CargarEstadoPrevio] Error consultando tercero fallback ID={vinc.IdTercero}", ex);
                                        // Fase 11: Notificar al usuario sobre el fallo en la consulta Dapper
                                        MessageBox.Show($"Ocurrió un error al intentar recuperar el docente inactivo/filtrado. Detalle: {ex.Message}", "Error de Carga", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                        }

                        if (vinc.IdSede > 0)
                        {
                            // Fase 17: Verificar que DataSource no sea nulo antes de asignar el SelectedValue
                            if (cmbSedes.DataSource != null)
                                cmbSedes.SelectedValue = vinc.IdSede;

                            // Defensa: si la sede no está en el dropdown (ej. está inactiva), 
                            // agregarla temporalmente para preservar la integridad de la vinculación
                            // Fase 13: Validación segura de tipo para SelectedValue
                            if (cmbSedes.SelectedValue == null || !(cmbSedes.SelectedValue is int selectedSedeId) || selectedSedeId != vinc.IdSede)
                            {
                                try
                                {
                                    var sedesList = cmbSedes.DataSource as List<LookupItem>;
                                    if (sedesList != null)
                                    {
                                        // Buscar el nombre de la sede desde la conexión compartida
                                        // Fase 15: Instanciar con suppressMarkDirty: true para evitar marcas innecesarias en operaciones de lectura
                                        var sedeRepo = new AcademicRepository(_sharedConn, suppressMarkDirty: true, modo: _modo);
                                        var sedeResult = sedeRepo.GetSedeById(vinc.IdSede);
                                        if (sedeResult.IsSuccess)
                                        {
                                            var sede = sedeResult.Value;
                                            string display = string.IsNullOrEmpty(sede.Subsede)
                                                ? sede.Ciudad
                                                : $"{sede.Ciudad} ({sede.Subsede})";

                                            sedesList.Add(new LookupItem { Id = vinc.IdSede, DisplayName = display + " [Inactiva]" });

                                            // Fase 12: Asignar DisplayMember/ValueMember antes del DataSource para evitar parpadeos visuales
                                            cmbSedes.DataSource = null;
                                            cmbSedes.DisplayMember = "DisplayName";
                                            cmbSedes.ValueMember = "Id";
                                            cmbSedes.DataSource = sedesList;

                                            cmbSedes.SelectedValue = vinc.IdSede;

                                            LoggerHelper.Warning($"[FormSesionesAcademicas] CargarEstadoPrevio: Sede ID={vinc.IdSede} está inactiva. Agregada temporalmente al dropdown con marca [Inactiva].");
                                        }
                                    }
                                }
                                catch (ObjectDisposedException)
                                {
                                    // Fase 16: Control de ObjectDisposedException en Sede fallback
                                    LoggerHelper.Error("[CargarEstadoPrevio] _sharedConn dispuesto en restauración sede.");
                                    MessageBox.Show("Conexión cerrada inesperadamente. No se pudo recuperar la sede inactiva.", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                                catch (Exception ex)
                                {
                                    // Fase 16: Control global de errores en Sede fallback
                                    LoggerHelper.Error($"[CargarEstadoPrevio] Error restaurando sede ID={vinc.IdSede}", ex);
                                }
                            }
                            else
                            {
                                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Sede restaurada Id={vinc.IdSede}");
                            }
                        }

                        // Desactivar event handlers temporalmente para evitar llamadas desincronizadas a UpdateAsignaturas
                        cmbProgramas.SelectedIndexChanged -= cmbProgramas_SelectedIndexChanged;
                        cmbSemestres.SelectedIndexChanged -= cmbSemestres_SelectedIndexChanged;

                        try
                        {
                            if (vinc.IdPrograma > 0)
                            {
                                // Fase 20: Verificar que DataSource no sea nulo antes de asignar el SelectedValue
                                if (cmbProgramas.DataSource != null)
                                    cmbProgramas.SelectedValue = vinc.IdPrograma;
                                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Programa restaurado Id={vinc.IdPrograma}. cmbProgramas.SelectedValue={cmbProgramas.SelectedValue}");
                            }

                            if (!string.IsNullOrWhiteSpace(vinc.Semestre))
                            {
                                cmbSemestres.SelectedItem = vinc.Semestre;
                                LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Semestre restaurado '{vinc.Semestre}'");
                            }

                            // Forzar una única actualización manual y controlada de las asignaturas
                            UpdateAsignaturas();
                        }
                        finally
                        {
                            // Reactivar event handlers
                            cmbProgramas.SelectedIndexChanged += cmbProgramas_SelectedIndexChanged;
                            cmbSemestres.SelectedIndexChanged += cmbSemestres_SelectedIndexChanged;
                        }

                        if (!string.IsNullOrWhiteSpace(vinc.Grupo))
                        {
                            cmbGrupos.SelectedItem = vinc.Grupo;
                        }
                        else
                        {
                            cmbGrupos.SelectedIndex = 0;
                        }

                        if (!string.IsNullOrWhiteSpace(vinc.Modalidad))
                        {
                            cmbTipoTutoria.SelectedItem = ModalidadTutoriaCatalog.NormalizarTutoria(vinc.Modalidad);
                        }

                        // Fase 21: Protección para cmbAsignaturas contra DataSources nulos
                        if (vinc.IdAsignatura > 0)
                        {
                            if (cmbAsignaturas.DataSource != null)
                                cmbAsignaturas.SelectedValue = vinc.IdAsignatura;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Asignatura restaurada Id={vinc.IdAsignatura}");
                        }

                        // Restaurar Horas
                        txtTotalHoras.Text = vinc.TotalHoras.ToString();

                        // Restaurar Fechas (Fase 2: Asignar FechaFin primero para no disparar ajuste de dtpInicio)
                        // Fase 18 / Fase 699: Suprimir ValueChanged para que dtpInicio/dtpFinal no disparen re-evaluaciones al restaurar
                        dtpInicio.ValueChanged -= DtpInicio_ValueChanged;
                        dtpFinal.ValueChanged -= DtpFinal_ValueChanged;
                        try
                        {
                            if (vinc.FechaFin.HasValue && vinc.FechaFin.Value > new DateTime(2000, 1, 1))
                                dtpFinal.Value = vinc.FechaFin.Value.Date;
                            if (vinc.FechaInicio.HasValue && vinc.FechaInicio.Value > new DateTime(2000, 1, 1))
                                dtpInicio.Value = vinc.FechaInicio.Value.Date;
                        }
                        finally
                        {
                            dtpInicio.ValueChanged += DtpInicio_ValueChanged;
                            dtpFinal.ValueChanged += DtpFinal_ValueChanged;
                        }

                        // Restaurar Días y Horas en el control custom
                        if (!string.IsNullOrWhiteSpace(vinc.DiaHorasSemanales))
                        {
                            txtDiasHoras.SelectedDays = RebuildSelectedDays(vinc.DiaHorasSemanales);
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Control de días restaurado: '{vinc.DiaHorasSemanales}'");
                        }

                        // Restaurar Pernocta Último Día
                        chkPernoctaUltimoDia.Checked = vinc.IsPernoctaUltimoDia == 1;

                        // Restaurar Terminal Aéreo
                        chkTerminalAereo.CheckedChanged -= chkTerminalAereo_CheckedChanged;
                        try
                        {
                            if (vinc.IsTerminalAereo == 1)
                            {
                                chkTerminalAereo.Checked = true;

                                cmbTerminalesAereas.SelectedIndexChanged -= cmbTerminalesAereas_SelectedIndexChanged;
                                try
                                {
                                    if (cmbTerminalesAereas.DataSource == null && vinc.IdTerminales > 0)
                                    {
                                        InvokeRepository(termRepo =>
                                        {
                                            var termResult = termRepo.GetTerminales();
                                            if (termResult.IsSuccess && termResult.Value.Any())
                                            {
                                                cmbTerminalesAereas.DisplayMember = "TerminalName";
                                                cmbTerminalesAereas.ValueMember = "Id";
                                                cmbTerminalesAereas.DataSource = termResult.Value.ToList();
                                            }
                                        });
                                    }

                                    if (vinc.IdTerminales > 0 && cmbTerminalesAereas.DataSource != null)
                                    {
                                        cmbTerminalesAereas.SelectedValue = vinc.IdTerminales;
                                    }

                                    var terminal = cmbTerminalesAereas.SelectedItem as Terminal;
                                    // Fallback si SelectedValue falló: buscar en items por Id
                                    if (terminal == null && vinc.IdTerminales > 0 && cmbTerminalesAereas.Items.Count > 0)
                                    {
                                        foreach (var item in cmbTerminalesAereas.Items)
                                        {
                                            if (item is Terminal t && t.Id == vinc.IdTerminales)
                                            {
                                                cmbTerminalesAereas.SelectedItem = t;
                                                terminal = t;
                                                break;
                                            }
                                        }
                                    }

                                    double gasto = vinc.GastoTerminalAerea;
                                    if (terminal != null && gasto <= 0)
                                        gasto = terminal.MaxValor;

                                    if (terminal != null)
                                    {
                                        lblMaxValor.Text = $"(Máximo: {terminal.MaxValor.ToString("C0", CultureInfo.CurrentCulture)})";
                                        txtGastoTerminalAerea.Text = gasto.ToString(CultureInfo.CurrentCulture);
                                    }
                                    else
                                    {
                                        lblMaxValor.Text = string.Empty;
                                        txtGastoTerminalAerea.Text = string.Empty;
                                    }
                                }
                                finally
                                {
                                    cmbTerminalesAereas.SelectedIndexChanged += cmbTerminalesAereas_SelectedIndexChanged;
                                }
                            }
                            else
                            {
                                chkTerminalAereo.Checked = false;
                            }
                        }
                        finally
                        {
                            pnlTransporteAereo.Visible = chkTerminalAereo.Checked;
                            chkTerminalAereo.CheckedChanged += chkTerminalAereo_CheckedChanged;
                        }
                    }));

                    // Restaurar Sesiones en el grid
                    var sesResult = repo.GetSesionesByVinculacion(idVinculacion);
                    if (sesResult.IsSuccess)
                    {
                        var sesiones = sesResult.Value.ToList();
                        LoggerHelper.Info($"[DIAGNOSTICO DB] Se encontraron {sesiones.Count} sesiones históricas para Vinculación ID {idVinculacion}:");
                        for (int idx = 0; idx < sesiones.Count; idx++)
                        {
                            var s = sesiones[idx];
                            LoggerHelper.Info($"  - Sesión [{idx + 1}]: ID={s.Id}, IdTercero={s.IdTercero}, Fecha={s.FechaSesion:yyyy-MM-dd}, Horario={s.HorarioSesion}, Horas={s.Horas}, Viática={s.Viatica}");
                        }

                        LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Cargando {sesiones.Count} sesiones históricas en el grid.");

                        this.Invoke((Action)(() =>
                        {
                            _sincronizandoGrid = true;
                            try
                            {
                                dgvSesiones.Rows.Clear();
                                var terceros = cmbTerceros.DataSource as List<TerceroDisplayDto>;
                                foreach (var s in sesiones)
                                {
                                    int i = dgvSesiones.Rows.Add();
                                    var r = dgvSesiones.Rows[i];
                                    r.Cells["ColChecked"].Value = true;

                                    var docente = terceros?.FirstOrDefault(t => t.Id == s.IdTercero);
                                    r.Cells["ColTercero"].Value = docente?.Nombre ?? s.IdTercero.ToString();

                                    r.Cells["ColDias"].Value = s.FechaSesion.ToString("dddd", CulturaEs).ToUpper();
                                    r.Cells["ColHoras"].Value = s.Horas;
                                    r.Cells["ColHorario"].Value = s.HorarioSesion;
                                    r.Cells["ColFechaSesion"].Value = s.FechaSesion.ToString("dd/MM/yyyy");
                                    r.Cells["ColViatica"].Value = SafeCellBool(s.Viatica);
                                    // Guardar referencia al modelo original para poder reconstruir al guardar
                                    r.Tag = s;
                                }
                                UpdateSelectionSummary();

                                // Verificación de integridad visual post-carga:
                                // Compara las viáticas almacenadas en la DB (via Tag) contra las que
                                // el grid muestra realmente. Si difieren, indica un bug de binding residual.
                                int viaticasEnDb = sesiones.Count(s => s.Viatica);
                                int viaticasEnGrid = dgvSesiones.Rows
                                    .Cast<DataGridViewRow>()
                                    .Count(r => !r.IsNewRow && SafeCellBool(r.Cells["ColViatica"].Value));
                                LoggerHelper.Info($"[VIATICA-CHECK] Sesiones cargadas: {sesiones.Count}. " +
                                    $"Viáticas en DB: {viaticasEnDb}. Viáticas visibles en grid: {viaticasEnGrid}.");

                                if (viaticasEnDb != viaticasEnGrid)
                                {
                                    LoggerHelper.Warning($"[VIATICA-CHECK] ¡DISCREPANCIA DETECTADA! " +
                                        $"DB={viaticasEnDb} vs Grid={viaticasEnGrid}. Revisar binding de ColViatica.");
                                }
                            }
                            finally
                            {
                                _sincronizandoGrid = false;
                            }
                        }));
                    }
                });

                // Fijar el estado como "Generado y Sincronizado" para evitar el aviso si no han tocado nada
                _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en CargarEstadoPrevio", ex);
            }
        }



        /// <summary>
        /// Convierte la cadena "VIERNES|4, SABADO|2" almacenada en la DB
        /// de vuelta al diccionario que espera el control DayHoursDropdown.
        /// </summary>
        private Dictionary<string, DaySessionInfo> RebuildSelectedDays(string raw)
        {
            var result = new Dictionary<string, DaySessionInfo>();
            if (string.IsNullOrWhiteSpace(raw)) return result;

            foreach (var part in raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var segments = part.Trim().Split('|');
                // Fase 995: double.TryParse con InvariantCulture para admitir horas fraccionarias
                if (segments.Length >= 2
                    && double.TryParse(segments[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double horasCrudas))
                {
                    string dia = segments[0].Trim();
                    string schedule = segments.Length >= 3 ? segments[2].Trim() : string.Empty;
                    int horas = (int)Math.Round(horasCrudas);

                    // Fase 994: día con AM+PM; el control admite una franja por día
                    if (result.TryGetValue(dia, out DaySessionInfo previo))
                    {
                        LoggerHelper.Warning(
                            $"[FormSesionesAcademicas] '{dia}' tiene franjas '{previo.Horario}' ({previo.Horas} h) y '{schedule}' ({horas} h); se conserva la de mayor carga.");
                        if (horas <= previo.Horas) continue;
                    }

                    result[dia] = new DaySessionInfo(horas, schedule);
                }
            }
            return result;
        }



        // Eliminado TranslateDay legacy (reemplazado por DayPatternParser)

        /// <summary>
        /// Pinta el grid <c>dgvSesiones</c> a partir de <c>_sesionesGeneradas</c>.
        /// Es la única fuente de verdad para la presentación visual del grid:
        /// asigna el Tag (requerido por <see cref="GridToSesionMapper"/>), y fuerza
        /// un <c>Refresh()</c> final para que <see cref="DataGridViewCheckBoxColumn"/>
        /// renderice correctamente los valores asignados de forma programática.
        /// </summary>
        private void DisplaySesiones()
        {
            _sincronizandoGrid = true;
            try
            {
                dgvSesiones.Rows.Clear();
                foreach (var s in _sesionesGeneradas)
                {
                    int i = dgvSesiones.Rows.Add();
                    var r = dgvSesiones.Rows[i];
                    r.Cells["ColChecked"].Value = true;

                    if (cmbTerceros.SelectedItem is TerceroDisplayDto terceroDto)
                        r.Cells["ColTercero"].Value = terceroDto.Nombre;
                    else
                        r.Cells["ColTercero"].Value = cmbTerceros.Text;

                    r.Cells["ColDias"].Value = s.FechaSesion.ToString("dddd", CulturaEs).ToUpper();
                    r.Cells["ColHoras"].Value = s.Horas;
                    r.Cells["ColHorario"].Value = s.HorarioSesion;
                    r.Cells["ColFechaSesion"].Value = s.FechaSesion.ToString("dd/MM/yyyy");
                    r.Cells["ColViatica"].Value = SafeCellBool(s.Viatica);
                    // El Tag enlaza la fila con el modelo de dominio. Es requerido por
                    // GridToSesionMapper.Extract() para validar integridad al guardar.
                    r.Tag = s;
                }
                UpdateSelectionSummary();
                // DataGridViewCheckBoxColumn requiere Refresh() explícito para renderizar
                // el estado checked cuando los valores se asignan programáticamente
                // (sin interacción del usuario ni DataSource vinculado al grid).
                dgvSesiones.Refresh();
            }
            finally
            {
                _sincronizandoGrid = false;
            }
        }

        private void UpdateSelectionSummary()
        {
            int count = dgvSesiones.Rows.Cast<DataGridViewRow>().Count(r => SafeCellBool(r.Cells["ColChecked"].Value));
            lblSelectionSummary.Text = $"{count} sesiones seleccionadas...";
        }

        private string GetDepartamentoBase()
        {
            try
            {
                var excel = new ExcelRepository();
                var sheet = excel.GetWorksheet("Config");
                if (sheet != null)
                {
                    object val = excel.GetCellValue(sheet, "B14");
                    return val?.ToString()?.ToUpper()?.Trim();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[FormSesionesAcademicas] No se pudo obtener el departamento base de Excel: {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// Sincroniza bidireccionalmente y en tiempo real las ediciones manuales realizadas por el usuario
        /// en las celdas del DataGridView hacia el modelo interno <see cref="Sesion"/> hospedado en <see cref="DataGridViewRow.Tag"/>.
        /// Implementa protección anti-reentrada (<see cref="_sincronizandoGrid"/>) y reversión automática ante valores inválidos.
        /// (Fase 237 / F1.4, F1.5, F1.6).
        /// </summary>
        /// <param name="sender">Origen del evento (dgvSesiones).</param>
        /// <param name="e">Argumentos con índices de fila y columna modificada.</param>
        private void dgvSesiones_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || _sincronizandoGrid) return;

            var row = dgvSesiones.Rows[e.RowIndex];
            if (row.IsNewRow) return;

            var col = dgvSesiones.Columns[e.ColumnIndex];
            if (col == null) return;

            // FASE 237 / F1.6: Identificación de columna por nombre semántico
            if (col.Name == "ColChecked")
            {
                UpdateSelectionSummary();
                return;
            }

            if (!(row.Tag is Sesion tag)) return;

            // FASE 237 / F1.4 & F1.5: Sincronización bidireccional inmediata con anti-reentrada y reversión
            switch (col.Name)
            {
                case "ColViatica":
                    tag.Viatica = SafeCellBool(row.Cells["ColViatica"].Value);
                    break;

                case "ColHoras":
                    if (GridCellValueParser.TryParseHoras(row.Cells["ColHoras"].Value, out double h))
                    {
                        tag.Horas = h;
                    }
                    else
                    {
                        RevertirCelda(row, "ColHoras", tag.Horas);
                    }
                    break;

                case "ColFechaSesion":
                    if (GridCellValueParser.TryParseFecha(row.Cells["ColFechaSesion"].Value, out DateTime f))
                    {
                        tag.FechaSesion = f.Date;
                        _sincronizandoGrid = true;
                        try
                        {
                            row.Cells["ColDias"].Value = f.ToString("dddd", CulturaEs).ToUpper();
                        }
                        finally
                        {
                            _sincronizandoGrid = false;
                        }
                    }
                    else
                    {
                        RevertirCelda(row, "ColFechaSesion", tag.FechaSesion.ToString("dd/MM/yyyy"));
                    }
                    break;

                case "ColHorario":
                    var txtHorario = row.Cells["ColHorario"].Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(txtHorario))
                    {
                        tag.HorarioSesion = txtHorario;
                    }
                    else
                    {
                        RevertirCelda(row, "ColHorario", tag.HorarioSesion);
                    }
                    break;
            }
        }

        /// <summary>
        /// Restaura el valor de una celda del grid a su estado previo válido almacenado en el modelo,
        /// aplicando color de advertencia, protección anti-reentrada y registrando el evento en el log.
        /// (Fase 2 / F2.6).
        /// </summary>
        /// <param name="row">Fila del DataGridView afectada.</param>
        /// <param name="colName">Nombre semántico de la columna a revertir.</param>
        /// <param name="original">Valor previo proveniente del modelo <see cref="Sesion"/>.</param>
        private void RevertirCelda(DataGridViewRow row, string colName, object original)
        {
            bool previo = _sincronizandoGrid;
            _sincronizandoGrid = true;
            try
            {
                row.Cells[colName].Value = original;
                row.Cells[colName].Style.BackColor = Color.MistyRose;
            }
            finally
            {
                _sincronizandoGrid = previo;
            }
            LoggerHelper.Warning($"[GridEdit] Valor inválido en {colName} (Fila {row.Index + 1}). Revertido al valor del modelo: '{original}'.");
        }

        private void dgvSesiones_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            // F2.7: El commit inmediato es necesario únicamente para ColChecked a fin de
            // que CellValueChanged dispare el contador de selección sin esperar al foco siguiente.
            // Limitar a esta columna evita commits prematuros sobre ColViatica o celdas de texto.
            int idxChecked = dgvSesiones.Columns["ColChecked"]?.Index ?? -1;
            if (dgvSesiones.IsCurrentCellDirty && idxChecked >= 0 && dgvSesiones.CurrentCell?.ColumnIndex == idxChecked)
            {
                dgvSesiones.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void btnToggleViatica_Click(object sender, EventArgs e)
        {
            if (dgvSesiones.Rows.Count == 0) return;

            _sincronizandoGrid = true;
            try
            {
                // Determinar el nuevo estado: si todos están marcados, desmarcar. Si no, marcar.
                bool allChecked = true;
                foreach (DataGridViewRow row in dgvSesiones.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (!SafeCellBool(row.Cells["ColViatica"].Value))
                    {
                        allChecked = false;
                        break;
                    }
                }

                bool newState = !allChecked;

                foreach (DataGridViewRow row in dgvSesiones.Rows)
                {
                    if (row.IsNewRow) continue;
                    row.Cells["ColViatica"].Value = newState;
                    // FIX: Mantener el Tag sincronizado con el grid cuando el usuario alterna el estado,
                    // así la función PrepararSesionesDesdeGrid captura los cambios manuales intencionales.
                    if (row.Tag is Sesion tagSesion)
                    {
                        tagSesion.Viatica = newState;
                    }
                }

                if (dgvSesiones.IsCurrentCellDirty)
                {
                    dgvSesiones.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
            finally
            {
                _sincronizandoGrid = false;
            }
        }

        /// <summary>
        /// Verifica de forma asíncrona el perfil del usuario activo para habilitar controles protegidos (btnEliminar).
        /// Reutiliza la conexión compartida sin bloquear la UI y aplica reglas de arquitectura usando el repositorio.
        /// </summary>
        private async Task VerificarPerfilAdmin()
        {
            if (_currentVinculacionId <= 0) return;

            try
            {
                // Fase 6 y Fase 9: Reutilizar _sharedConn y evitar bloquear UI
                if (_sharedConn != null && _sharedConn.State == ConnectionState.Open)
                {
                    bool esAdmin = await Task.Run(() =>
                    {
                        var repo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.AcademicRepository(_sharedConn, suppressMarkDirty: true);
                        var result = repo.GetUsuarioByNombreUsuario(Environment.UserName);

                        return result.IsSuccess && result.Value != null &&
                               string.Equals(result.Value.Perfil, "Admin", StringComparison.OrdinalIgnoreCase);
                    });

                    if (esAdmin)
                    {
                        if (this.InvokeRequired)
                        {
                            this.Invoke((Action)(() => btnEliminar.Visible = true));
                        }
                        else
                        {
                            btnEliminar.Visible = true;
                        }
                    }
                }
            }
            catch (System.Data.SQLite.SQLiteException exSq)
            {
                LoggerHelper.Warning($"[FormSesionesAcademicas] Error SQLite verificando perfil: {exSq.Message}");
                if (this.InvokeRequired) { this.Invoke((Action)(() => btnEliminar.Visible = false)); } else { btnEliminar.Visible = false; }
            }
            catch (TimeoutException exTo)
            {
                LoggerHelper.Warning($"[FormSesionesAcademicas] Timeout verificando perfil: {exTo.Message}");
                if (this.InvokeRequired) { this.Invoke((Action)(() => btnEliminar.Visible = false)); } else { btnEliminar.Visible = false; }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error verificando perfil de usuario", ex);
                if (this.InvokeRequired) { this.Invoke((Action)(() => btnEliminar.Visible = false)); } else { btnEliminar.Visible = false; }
            }
        }

        private async void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentVinculacionId <= 0) return;

            if (_guardadoCts != null && !_guardadoCts.IsCancellationRequested)
            {
                MessageBox.Show("Hay un guardado en progreso. Espere a que termine antes de eliminar.",
                    "Operación en Curso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show(
                "¿Está seguro de que desea eliminar permanentemente esta vinculación y todas sus sesiones? Esta acción no se puede deshacer.",
                "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result != DialogResult.Yes) return;

            btnEliminar.Enabled = false;
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Fase 7: Validación de Conexión Nula al Eliminar
                if (_sharedConn == null || _sharedConn.State != ConnectionState.Open)
                {
                    MessageBox.Show("Conexión a la base de datos no disponible. Por favor cierre y reabra el formulario.", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Fase 17: Forzar persistencia al vault antes de cerrar/eliminar de manera asíncrona para no bloquear UI y evitar inconsistencia
                var scheduler = ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler;
                if (scheduler != null)
                {
                    using (var persistCtx = SqliteHelper.GetContextForPath())
                    {
                        if (persistCtx != null && !persistCtx.IsDisposed)
                        {
                            try
                            {
                                await scheduler.ForcePersistVaultAsync(persistCtx);
                                scheduler.MarkDirty();
                            }
                            catch (Exception exV)
                            {
                                LoggerHelper.Warning($"[FormSesionesAcademicas] No se pudo persistir vault: {exV.Message}");
                                MessageBox.Show("No se pudo confirmar la eliminación en la base de datos principal. Intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                btnEliminar.Enabled = true;
                                this.Cursor = Cursors.Default;
                                return;
                            }
                        }
                    }
                }

                Vinculacion vinculoEliminado = null;
                try
                {
                    vinculoEliminado = _sharedConn.QuerySingleOrDefault<Vinculacion>(
                        $"SELECT id as Id, idtercero as IdTercero, idprograma as IdPrograma, idasignatura as IdAsignatura, semestre as Semestre, subunidad as Subunidad FROM {_modo.TablaVinculaciones} WHERE id = @id",
                        new { id = _currentVinculacionId });
                }
                catch { }

                using (var transaction = _sharedConn.BeginTransaction())
                {
                    try
                    {
                        var repo = new AcademicRepository(_sharedConn, transaction, subunidad: null, suppressMarkDirty: false, modo: _modo);
                        var deleteSesionesResult = repo.DeleteSesionesByVinculacion(_currentVinculacionId);
                        if (deleteSesionesResult.IsFailure)
                            throw new Exception($"Error eliminando sesiones: {deleteSesionesResult.Error}");

                        var deleteVinculacionResult = repo.DeleteVinculacion(_currentVinculacionId);
                        if (deleteVinculacionResult.IsFailure)
                            throw new Exception(deleteVinculacionResult.Error);

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                if (vinculoEliminado != null)
                {
                    try
                    {
                        var poolRepo = new TerceroAsignaturaRepository(_connectionFactory);
                        var syncService = new TerceroAsignaturaSyncService(_connectionFactory, poolRepo);
                        syncService.SincronizarDocenteAsignaturaAsync(
                            vinculoEliminado.IdTercero,
                            vinculoEliminado.IdPrograma,
                            vinculoEliminado.IdAsignatura,
                            vinculoEliminado.Semestre,
                            vinculoEliminado.Subunidad,
                            CancellationToken.None).GetAwaiter().GetResult();
                    }
                    catch (Exception exSync)
                    {
                        LoggerHelper.Warning($"[FormSesionesAcademicas] Error no fatal al resincronizar pool tras eliminar: {exSync.Message}");
                    }
                }

                LoggerHelper.Info($"[FormSesionesAcademicas] Vinculación {_currentVinculacionId} y sus sesiones eliminadas exitosamente.");
                MessageBox.Show("La vinculación y sus sesiones se eliminaron correctamente.", "Eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DatosGuardados = true;
                this.Close();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error eliminando vinculación", ex);
                MessageBox.Show($"Ocurrió un error al intentar eliminar la vinculación:\n{ex.Message}", "Error de Eliminación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnEliminar.Enabled = true;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}
