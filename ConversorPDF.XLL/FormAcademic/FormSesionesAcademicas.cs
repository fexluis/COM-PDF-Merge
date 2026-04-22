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
using ConversorPDF.XLL.Core.Data;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Models.Academic;
using ConversorPDF.XLL.Core.Services.Academic;
using ConversorPDF.XLL.UI;
using Dapper;

namespace ConversorPDF.XLL.FormAcademic
{
    /// <summary>
    /// Formulario de gestión de sesiones académicas.
    /// </summary>
    public partial class FormSesionesAcademicas : Form
    {
        private readonly GeneradorSesionesRecurrentes _generador;
        private List<Sesion> _sesionesGeneradas;
        
        private bool _esquemaRequiereMigracion = false;
        private string _ddlPendiente = null;
        private int _currentVinculacionId = 1;
        private CancellationTokenSource _guardadoCts;

        private EncryptedSQLiteContext _sharedCtx;
        private System.Data.SQLite.SQLiteConnection _sharedConn;

        /// <summary>
        /// Guarda un "Hash" o texto codificado con el estado exacto de los controles vitales 
        /// (Tercero, Horas, Fechas, Distribución de días) al momento de dar clic en [Generar].
        /// Se usa para abortar el guardado si el usuario modificó estas variables sin sincronizar el grid.
        /// </summary>
        private string _ultimoEstadoGenerado = string.Empty;
        private readonly IAcademicDateSyncService _dateSyncService;

        public FormSesionesAcademicas(ProgressManager progress = null)
        {
            _generador = new GeneradorSesionesRecurrentes();
            _sesionesGeneradas = new List<Sesion>();
            _dateSyncService = new AcademicDateSyncService();

            progress?.AddProgress(10); // 10% - Inicio

            // 1. Asegurar motor (Nivel 1)
            SQLiteEngineBootstrap.EnsureInitialized();
            progress?.AddProgress(10); // 20% - Motor listo

            InitializeComponent();
            SetupGrid();
        }

        /// <summary>
        /// Wrapper efímero para interactuar con la base de datos. 
        /// Si el contexto compartido está abierto (OnShown), lo reutiliza para evitar I/O redundante.
        /// </summary>
        private void InvokeRepository(Action<IAcademicRepository> repoAction)
        {
            if (_sharedConn != null && _sharedConn.State == ConnectionState.Open)
            {
                repoAction(new AcademicRepository(_sharedConn));
                return;
            }

            LoggerHelper.Warning("[FormSesionesAcademicas] InvokeRepository: Contexto compartido no disponible. Usando fallback efímero (No recomendado en v1.2.50.0).");
            using (var ctx = SqliteHelper.GetContext())
            using (var conn = ctx.Open())
            {
                repoAction(new AcademicRepository(conn));
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            
            try
            {
                // B.2 - Limpieza proactiva de archivos huérfanos antes de abrir el nuevo
                EncryptedSQLiteContext.CleanupOrphanedTempFiles();

                // B.1 - Inicialización de Contexto Compartido (Una sola desencriptación)
                string vaultPath = SqliteHelper.GetVaultPath();
                LoggerHelper.Info($"[FormSesionesAcademicas] OnShown: Abriendo shared context en {vaultPath}...");
                _sharedCtx = SqliteHelper.GetContext();
                _sharedConn = _sharedCtx.Open();

                // 1. Inyectar datos iniciales si es necesario (Catálogos)
                ExecuteSeedData();

                // 2. Cargar datos iniciales en combos
                LoadInitialData(null);

                // 3. Restaurar estado guardado (Vinculación y Sesiones previas)
                CargarEstadoPrevio();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en OnShown al cargar datos iniciales", ex);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _guardadoCts?.Cancel();
            
            // Liberar contexto compartido
            if (_sharedCtx != null)
            {
                LoggerHelper.Info("[FormSesionesAcademicas] Cerrando shared context...");
                _sharedCtx.Dispose();
                _sharedCtx = null;
                _sharedConn = null;
            }

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

            dgvSesiones.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ColChecked", HeaderText = "", Width = 30 });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColTercero", HeaderText = "Tercero", Width = 150, DataPropertyName = "IdTercero" });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDias", HeaderText = "Día Semana", Width = 120 });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColHoras", HeaderText = "Horas", Width = 60, DataPropertyName = "Horas" });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDuracion", HeaderText = "Duración", Width = 80, DataPropertyName = "Duracion" });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColInicio", HeaderText = "Fecha Inicial", Width = 100, DataPropertyName = "fechasesionInicial" });
            dgvSesiones.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColFinal", HeaderText = "Fecha Final", Width = 100, DataPropertyName = "fechasesionFinal" });
            dgvSesiones.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ColViatica", HeaderText = "Viática", Width = 60, DataPropertyName = "Viatica" });
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

                if (CatalogCacheService.Sedes.Any())
                {
                    cmbSedes.DataSource = CatalogCacheService.Sedes;
                    cmbSedes.DisplayMember = "Ciudad";
                    cmbSedes.ValueMember = "Id";
                }
                else LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando sedes: Caché vacío.");
                progress?.AddProgress(10); // 50%

                // 2. Programas
                LoggerHelper.Info("[FormSesionesAcademicas] Aplicando Programas desde Caché...");
                if (CatalogCacheService.Programas.Any())
                {
                    cmbProgramas.DataSource = CatalogCacheService.Programas;
                    cmbProgramas.DisplayMember = "NombrePrograma";
                    cmbProgramas.ValueMember = "Id";
                }
                else LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando programas: Caché vacío.");
                progress?.AddProgress(10); // 60%

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

                    cmbTerceros.DataSource = tercerosList;
                    cmbTerceros.DisplayMember = "Nombre";
                    cmbTerceros.ValueMember = "Id";
                }
                else LoggerHelper.Warning($"[FormSesionesAcademicas] Error cargando terceros: Caché vacío.");
                progress?.AddProgress(20); // 80%

                // Las fechas se restaurarán desde la DB en CargarEstadoPrevio si hay estado previo.
                // Si no, usamos los valores por defecto.
                dtpInicio.Value = DateTime.Today;
                dtpFinal.Value = DateTime.Today.AddMonths(1);

                progress?.AddProgress(20); // 100% - Completo
                LoggerHelper.Info("[FormSesionesAcademicas] LoadInitialData: Carga completada con éxito.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en LoadInitialData", ex);
                MessageBox.Show($"Error cargando datos: {ex.Message}", "Error Diagnostics", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbProgramas_SelectedIndexChanged(object sender, EventArgs e) => UpdateAsignaturas();
        private void cmbSemestres_SelectedIndexChanged(object sender, EventArgs e) => UpdateAsignaturas();

        private void UpdateAsignaturas()
        {
            if (!(cmbProgramas.SelectedValue is int idProg) || !(cmbSemestres.SelectedItem is string sem))
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

                        cmbAsignaturas.DataSource = filtered;
                        cmbAsignaturas.DisplayMember = "NombreAsignatura";
                        cmbAsignaturas.ValueMember = "Id";
                        
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
                        var cols = _sharedConn.Query<dynamic>("PRAGMA table_info(vinculaciones);");
                        if (cols.Any(c => ((string)c.name).ToLower() == "idtercero"))
                        {
                            needsMigration = false;
                        }
                    }
                    catch { /* Fallback a True si falla el PRAGMA */ }

                    if (needsMigration)
                    {
                        _esquemaRequiereMigracion = true;
                        _ddlPendiente = "ALTER TABLE vinculaciones ADD COLUMN idtercero INTEGER REFERENCES terceros(id)";
                        LoggerHelper.Info("[Seed] Migración de esquema (B.4) detectada como pendiente.");
                    }

                    // 2. Poblado de Catálogos Unificado (B.5)
                    // Auditoría forense de ruta de lectura (Diagnostico C-02)
                    string vaultPath = SqliteHelper.DbPathEnc;
                    LoggerHelper.Info($"[AUDIT-VAULT] Iniciando formulario académico (ExecuteSeedData). Vault Lectura: {vaultPath}");

                    LoggerHelper.Info("[FormSesionesAcademicas] Ejecutando Seed Unificado desde Recurso Embebido...");
                    repo.SeedCatalogosTransaccional();
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en ExecuteSeedData (Phase B Optimization)", ex);
            }
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
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

                // 2. Confirmación si ya hay datos
                if (dgvSesiones.Rows.Count > 0)
                {
                    if (MessageBox.Show("¿Desea sobrescribir las sesiones actuales?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    {
                        LoggerHelper.Info("[FormSesionesAcademicas] btnGenerar: Usuario canceló la sobrescritura.");
                        return;
                    }
                }

                dgvSesiones.Rows.Clear();
                _sesionesGeneradas.Clear();

                // 3. Preparación de Variables
                DateTime fechaActual = dtpInicio.Value.Date;
                double horasAcumuladas = 0;
                var patronSemanas = ParsearSelectedDays(txtDiasHoras.SelectedDays);
                int safetyGuard = 0;

                string patronLog = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value}"));
                LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Configuraciones leídas. Objetivo: {totalHorasObjetivo}h. Patrón: {patronLog}");

                // 4. Bucle de Generación por "Horas Exactas"
                while (horasAcumuladas < totalHorasObjetivo && safetyGuard < 365)
                {
                    if (patronSemanas.ContainsKey(fechaActual.DayOfWeek))
                    {
                        double horasBase = patronSemanas[fechaActual.DayOfWeek];
                        double horasFaltantes = totalHorasObjetivo - horasAcumuladas;
                        double horasAsignar = Math.Min(horasBase, horasFaltantes);

                        LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Ciclo {safetyGuard}. Fecha: {fechaActual:dd/MM/yyyy}. " +
                                         $"Faltan: {horasFaltantes}h. Asignando en esta sesión: {horasAsignar}h.");

                        int rowIndex = dgvSesiones.Rows.Add();
                        var row = dgvSesiones.Rows[rowIndex];

                        row.Cells["ColChecked"].Value = true;
                        row.Cells["ColTercero"].Value = cmbTerceros.Text;
                        row.Cells["ColDias"].Value = fechaActual.ToString("dddd", new CultureInfo("es-ES")).ToUpper();
                        row.Cells["ColHoras"].Value = horasAsignar;
                        row.Cells["ColDuracion"].Value = horasAsignar;
                        row.Cells["ColInicio"].Value = fechaActual.ToString("dd/MM/yyyy");
                        row.Cells["ColFinal"].Value = fechaActual.ToString("dd/MM/yyyy");
                        row.Cells["ColViatica"].Value = false;

                        var sesion = new Sesion 
                        { 
                            IdVinculacion = _currentVinculacionId,
                            IdTercero = (int)cmbTerceros.SelectedValue,
                            fechasesionInicial = fechaActual,
                            fechasesionFinal = fechaActual,
                            Horas = horasAsignar,
                            Duracion = horasAsignar,
                            Viatica = false
                        };

                        _sesionesGeneradas.Add(sesion);
                        row.Tag = sesion;

                        horasAcumuladas += horasAsignar;
                    }
                    fechaActual = fechaActual.AddDays(1);
                    safetyGuard++;
                }

                // 5. Informe Final y Auto-Ajuste de Fechas
                UpdateSelectionSummary();
                int sesionesRealesGeneradas = dgvSesiones.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                LoggerHelper.Info($"[FormSesionesAcademicas] btnGenerar: Generación exitosa. Sesiones: {sesionesRealesGeneradas}. Total: {horasAcumuladas}h.");

                var adjustment = _dateSyncService.CalculateAdjustment(_sesionesGeneradas, dtpInicio.Value, dtpFinal.Value);
                
                string msgBase = $"Generación completada: {horasAcumuladas} horas repartidas en {sesionesRealesGeneradas} sesiones.";

                if (adjustment.NeedsAdjustment)
                {
                    // Actualizar controles
                    dtpInicio.Value = adjustment.CalculatedStart;
                    dtpFinal.Value = adjustment.CalculatedEnd;

                    string detalles = string.Join("\n", adjustment.Messages);
                    MessageBox.Show(msgBase + "\n\nSe detectaron ajustes automáticos basados en la lógica de generación:\n" + detalles, 
                                    "Generación Exitosa - Fechas Auto-Ajustadas", MessageBoxButtons.OK, adjustment.ExceededOriginalLimit ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(msgBase + $"\n\nPeríodo abarcado: {dtpInicio.Value:dd/MM/yyyy} al {dtpFinal.Value:dd/MM/yyyy}.", 
                                    "Generación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                // Guardar el estado validado temporal
                _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormSesionesAcademicas] Error en btnGenerar_Click", ex);
                MessageBox.Show($"Error durante la generación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnGuardar_Click(object sender, EventArgs e)
        {
            // C.1 - Corregido: La validación ahora tiene sentido lógico
            if (ParametrosEstanDesincronizados()) 
            {
                MessageBox.Show(
                    "Has modificado los parámetros maestros (Docente, Horas, Distribución de días o Fechas).\n\n" +
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

            _guardadoCts = new CancellationTokenSource();
            
            using (var dlg = new ConversorPDF.XLL.UI.ProgressDialog("Guardando Cambios..."))
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
                            
                            _ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos();
                            CargarEstadoPrevio(); // Refrescar vista
                            dlg.Close();
                        }));
                    }
                    catch (OperationCanceledException)
                    {
                        LoggerHelper.Info("[Guardado] Operación cancelada por el usuario.");
                        // El diálogo se cierra por el usuario o por el código si es necesario
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
                        _guardadoCts?.Dispose();
                        _guardadoCts = null;
                    }
                };

                dlg.ShowDialog(this);
            }
        }

        private async Task GuardarAsync(ProgressDialog dlg, CancellationToken ct)
        {
            var vinculoActual = PrepararVinculacionDesdeForm();
            var sesiones = PrepararSesionesDesdeGrid().ToList();
            
            byte[] plainDatabaseBytes = null;
            string tempDbPath = string.Empty;

            // FASE 1 y 2: Guardado Atómico usando Contexto Compartido
            await Task.Run(() => 
            {
                // Usamos la conexión ya abierta (Shared Context)
                if (_sharedConn == null || _sharedConn.State != ConnectionState.Open)
                    throw new InvalidOperationException("Contexto de base de datos no disponible para guardado.");

                tempDbPath = _sharedCtx.TempDbPath;
                
                // Transacción real a nivel de base de datos
                using (var transaction = _sharedConn.BeginTransaction())
                {
                    try
                    {
                        var repo = new AcademicRepository(_sharedConn, transaction);
                        
                        ct.ThrowIfCancellationRequested();
                        // dlg.Progress.Report("Sincronizando catálogos (v1.2.50.0)...");
                        // repo.SeedCatalogosTransaccional(); // FIX-1.2: Innecesario en cada guardado

                        ct.ThrowIfCancellationRequested();
                        dlg.Progress.Report($"Guardando vinculación y {sesiones.Count} sesiones...");

                        if (vinculoActual.Id > 0)
                        {
                            repo.UpdateVinculacion(vinculoActual);
                            repo.DeleteSesionesByVinculacion(vinculoActual.Id);
                        }
                        else
                        {
                            var resVinc = repo.InsertVinculacion(vinculoActual);
                            vinculoActual.Id = resVinc.Value;
                            _currentVinculacionId = vinculoActual.Id;
                            foreach(var s in sesiones) s.IdVinculacion = vinculoActual.Id;
                        }

                        repo.InsertSesiones(sesiones);

                        ct.ThrowIfCancellationRequested();
                        dlg.Progress.Report("Confirmando transacción...");
                        transaction.Commit();
                        
                        // FIX-0.2: Forzar checkpoint WAL antes de cerrar para asegurar que el .db est actualizado
                        try
                        {
                            using (var cmdWal = new System.Data.SQLite.SQLiteCommand("PRAGMA wal_checkpoint(TRUNCATE);", _sharedConn))
                            {
                                int walResult = cmdWal.ExecuteNonQuery();
                                LoggerHelper.Info($"[AUDIT-WAL] WAL checkpoint ejecutado exitosamente. Resultado={walResult}");
                            }
                        }
                        catch (Exception exWal)
                        {
                            LoggerHelper.Warning($"[AUDIT-WAL] Error no crítico en WAL checkpoint: {exWal.Message}");
                        }

                        LoggerHelper.Info($"[PERF] Transacción v1.2.50.3 completada. VincID={vinculoActual.Id}");
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                // === FIX-0.1: Liberar archivo ANTES de lectura ===
                // SQLite mantiene el file handle aunque la transaccin haya hecho commit.
                // Es obligatorio cerrar la conexin y limpiar pools nativos para que Windows
                // libere el lock sobre el archivo temporal.
                _sharedCtx.CloseConnectionResiliently();
                _sharedConn = null;
                System.Data.SQLite.SQLiteConnection.ClearAllPools();

                // La re-encriptación ocurre DESPUÉS del commit exitoso
                ct.ThrowIfCancellationRequested();
                dlg.Progress.Report("Protegiendo base de datos segura...");
                
                // Leemos del archivo temporal ya desbloqueado
                plainDatabaseBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(tempDbPath, dlg.Progress);
                
                if (plainDatabaseBytes == null)
                    throw new IOException("Fallo crítico: No se pudo leer la base de datos temporal tras liberar el lock.");

                // A.5 - Auditoría final de guardado
                LoggerHelper.Info($"[AUDITORIA] Guardado Atómico v1.2.50.2: Tercero={vinculoActual.IdTercero}, Sesiones={sesiones.Count}");
                
                SqliteHelper.UpdateVaultDirectly(plainDatabaseBytes, dlg.Progress);

                // === RESTAURAR CONEXIN PARA CONTINUIDAD DEL FORMULARIO ===
                dlg.Progress.Report("Reabriendo conexión segura...");
                _sharedConn = _sharedCtx.Open();
            });
        }

        private Dictionary<DayOfWeek, double> ParsearSelectedDays(Dictionary<string, int> input)
        {
            return DayPatternParser.Parse(input);
        }

        /// <summary>
        /// Escanea los controles visuales críticos que afectan matemáticamente el
        /// algoritmo generador de sesiones.
        /// </summary>
        private string ObtenerEstadoParametrosCriticos()
        {
            string diasHorasStr = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value}"));
            int idTercero = cmbTerceros.SelectedValue is int idT ? idT : 0;
            int idPrograma = cmbProgramas.SelectedValue is int idP ? idP : 0;
            int idAsignatura = cmbAsignaturas.SelectedValue is int idA ? idA : 0;
            string semestre = cmbSemestres.SelectedItem?.ToString() ?? "N/A";
            
            string totalHoras = txtTotalHoras.Text.Trim();
            string fi = dtpInicio.Value.Date.ToString("yyyyMMdd");
            string ff = dtpFinal.Value.Date.ToString("yyyyMMdd");

            return $"T={idTercero};P={idPrograma};S={semestre};A={idAsignatura};H={totalHoras};D={diasHorasStr};FI={fi};FF={ff}";
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
            int idTercero = cmbTerceros.SelectedValue is int idTerc ? idTerc : 0;
            
            // FASE 1: Validación Defensiva de Tercero (Cross-Validation)
            var vinculacion = new Vinculacion
            {
                Id = _currentVinculacionId, 
                IdSede = cmbSedes.SelectedValue is int idSede ? idSede : 0,
                IdTercero = idTercero,
                IdPrograma = cmbProgramas.SelectedValue is int idAs ? idAs : 0,
                Semestre = cmbSemestres.SelectedItem?.ToString() ?? "I",
                IdAsignatura = cmbAsignaturas.SelectedValue is int idAsig ? idAsig : 0,
                TotalHoras = int.TryParse(txtTotalHoras.Text, out int th) ? th : 0,
                DiaHorasSemanales = string.Join(", ", txtDiasHoras.SelectedDays.Select(k => $"{k.Key}|{k.Value}")),
                FechaInicio = dtpInicio.Value.Date,
                FechaFin = dtpFinal.Value.Date
            };

            // Validar que el objeto esté completo
            if (vinculacion.IdTercero == 0 || vinculacion.IdSede == 0 || vinculacion.IdAsignatura == 0)
            {
                throw new InvalidOperationException("Faltan parámetros maestros por seleccionar (Tercero, Sede o Asignatura).");
            }

            return vinculacion;
        }

        private IEnumerable<Sesion> PrepararSesionesDesdeGrid()
        {
            // FASE UI: Forzar commit global desenfocando la celda actual
            dgvSesiones.CurrentCell = null;
            if (dgvSesiones.IsCurrentCellDirty) dgvSesiones.CommitEdit(DataGridViewDataErrorContexts.Commit);
            dgvSesiones.EndEdit();
            
            // Construir DTOs abstractos desde WinForms
            var rows = new List<GridRowDto>();
            foreach (DataGridViewRow row in dgvSesiones.Rows)
            {
                double horas = double.TryParse(row.Cells["ColHoras"].Value?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double h) ? h : 0;
                double duracion = double.TryParse(row.Cells["ColDuracion"].Value?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double d) ? d : horas;

                DateTime fi = DateTime.TryParseExact(row.Cells["ColInicio"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var fiD) ? fiD : dtpInicio.Value.Date;
                DateTime ff = DateTime.TryParseExact(row.Cells["ColFinal"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var ffD) ? ffD : fi;

                rows.Add(new GridRowDto
                {
                    IsChecked = row.Cells["ColChecked"].Value is bool b && b,
                    IsNewRow = row.IsNewRow,
                    Tag = row.Tag,
                    Horas = horas,
                    Duracion = duracion,
                    FechaInicial = fi,
                    FechaFinal = ff,
                    Viatica = row.Cells["ColViatica"].Value is bool vb && vb
                });
            }

            int idTerceroMaster = cmbTerceros.SelectedValue is int id ? id : 0;
            
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
                        return;
                    }

                    var vinc = vincResult.Value;
                    LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Vinculación encontrada. " +
                                      $"IdTercero={vinc.IdTercero}, Horas={vinc.TotalHoras}, Días='{vinc.DiaHorasSemanales}', " +
                                      $"Inicio={vinc.FechaInicio:dd/MM/yyyy}, Fin={vinc.FechaFin:dd/MM/yyyy}");

                    // ── Restaurar combos de Vinculación ──
                    this.Invoke((Action)(() =>
                    {
                        if (vinc.IdTercero > 0)
                        {
                            cmbTerceros.SelectedValue = vinc.IdTercero;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Tercero restaurado Id={vinc.IdTercero}");
                        }

                        if (vinc.IdSede > 0)
                        {
                            cmbSedes.SelectedValue = vinc.IdSede;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Sede restaurada Id={vinc.IdSede}");
                        }

                        if (vinc.IdPrograma > 0)
                        {
                            // Asignar programa dispara UpdateAsignaturas, por eso se asigna primero
                            cmbProgramas.SelectedValue = vinc.IdPrograma;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Programa restaurado Id={vinc.IdPrograma}");
                        }

                        if (!string.IsNullOrWhiteSpace(vinc.Semestre))
                        {
                            cmbSemestres.SelectedItem = vinc.Semestre;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Semestre restaurado '{vinc.Semestre}'");
                        }

                        if (vinc.IdAsignatura > 0)
                        {
                            cmbAsignaturas.SelectedValue = vinc.IdAsignatura;
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Asignatura restaurada Id={vinc.IdAsignatura}");
                        }

                        // Restaurar Horas
                        txtTotalHoras.Text = vinc.TotalHoras.ToString();

                        // Restaurar Fechas
                        if (vinc.FechaInicio > new DateTime(2000, 1, 1))
                            dtpInicio.Value = vinc.FechaInicio.Date;
                        if (vinc.FechaFin > new DateTime(2000, 1, 1))
                            dtpFinal.Value = vinc.FechaFin.Date;

                        // Restaurar Días y Horas en el control custom
                        if (!string.IsNullOrWhiteSpace(vinc.DiaHorasSemanales))
                        {
                            txtDiasHoras.SelectedDays = RebuildSelectedDays(vinc.DiaHorasSemanales);
                            LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Control de días restaurado: '{vinc.DiaHorasSemanales}'");
                        }
                    }));

                    // Restaurar Sesiones en el grid
                    var sesResult = repo.GetSesionesByVinculacion(idVinculacion);
                    if (sesResult.IsSuccess)
                    {
                        var sesiones = sesResult.Value.ToList();
                        LoggerHelper.Info($"[FormSesionesAcademicas] CargarEstadoPrevio: Cargando {sesiones.Count} sesiones históricas en el grid.");
                        
                        this.Invoke((Action)(() =>
                        {
                            dgvSesiones.Rows.Clear();
                            var terceros = cmbTerceros.DataSource as List<Tercero>;
                            foreach (var s in sesiones)
                            {
                                int i = dgvSesiones.Rows.Add();
                                var r = dgvSesiones.Rows[i];
                                r.Cells["ColChecked"].Value = true;
                                
                                var docente = terceros?.FirstOrDefault(t => t.Id == s.IdTercero);
                                r.Cells["ColTercero"].Value = docente?.Nombre ?? s.IdTercero.ToString();

                                r.Cells["ColDias"].Value = s.fechasesionInicial.ToString("dddd", new CultureInfo("es-ES")).ToUpper();
                                r.Cells["ColHoras"].Value = s.Horas;
                                r.Cells["ColDuracion"].Value = s.Duracion;
                                r.Cells["ColInicio"].Value = s.fechasesionInicial.ToString("dd/MM/yyyy");
                                r.Cells["ColFinal"].Value = s.fechasesionFinal.ToString("dd/MM/yyyy");
                                r.Cells["ColViatica"].Value = s.Viatica;
                                    // Guardar referencia al modelo original para poder reconstruir al guardar
                                r.Tag = s;
                            }
                            UpdateSelectionSummary();
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
        private Dictionary<string, int> RebuildSelectedDays(string raw)
        {
            var result = new Dictionary<string, int>();
            if (string.IsNullOrWhiteSpace(raw)) return result;

            foreach (var part in raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var segments = part.Trim().Split('|');
                if (segments.Length == 2 && int.TryParse(segments[1].Trim(), out int hours))
                {
                    result[segments[0].Trim()] = hours;
                }
            }
            return result;
        }



        // Eliminado TranslateDay legacy (reemplazado por DayPatternParser)

        private void DisplaySesiones()
        {
            dgvSesiones.Rows.Clear();
            foreach (var s in _sesionesGeneradas)
            {
                int i = dgvSesiones.Rows.Add();
                var r = dgvSesiones.Rows[i];
                r.Cells["ColChecked"].Value = true;
                r.Cells["ColTercero"].Value = cmbTerceros.Text;
                r.Cells["ColDias"].Value = s.Duracion > 1 ? "VIERNES-SABADO" : s.fechasesionInicial.DayOfWeek.ToString().ToUpper();
                r.Cells["ColHoras"].Value = s.Horas;
                r.Cells[4].Value = s.Duracion;
                r.Cells["ColInicio"].Value = s.fechasesionInicial.ToString("dd/MM/yyyy");
                r.Cells["ColFinal"].Value = s.fechasesionFinal.ToString("dd/MM/yyyy");
                r.Cells["ColViatica"].Value = s.Viatica;
            }
            UpdateSelectionSummary();
        }

        private void UpdateSelectionSummary()
        {
            int count = dgvSesiones.Rows.Cast<DataGridViewRow>().Count(r => Convert.ToBoolean(r.Cells["ColChecked"].Value));
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

        private void dgvSesiones_CellValueChanged(object sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex == 0) UpdateSelectionSummary(); }
        private void dgvSesiones_CurrentCellDirtyStateChanged(object sender, EventArgs e) { if (dgvSesiones.IsCurrentCellDirty) dgvSesiones.CommitEdit(DataGridViewDataErrorContexts.Commit); }
    }
}
