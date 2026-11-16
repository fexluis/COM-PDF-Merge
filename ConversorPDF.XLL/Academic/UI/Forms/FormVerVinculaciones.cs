using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.VisTimeline.Application.Services;
using ConversorPDF.XLL.VisTimeline.UI;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Excel;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.UI.Enums;
using ConversorPDF.XLL.Academic.Domain.Enums;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormVerVinculaciones : Form
    {
        private readonly List<SesionListItem> _items;
        private readonly IDbConnection _dbConnection;
        private readonly dynamic _excelApp;
        private readonly AgModoDatos _modo;
        private Label lblModo;
        private Dictionary<string, Tercero> _cacheTerceros;
        private VistaVinculaciones _vistaActual = VistaVinculaciones.PorAsignaturas;
        private string _nombreLibroInicial;

        private readonly ISalvedadConflictoRepository _salvedadRepository;
        private IReadOnlyCollection<SalvedadConflicto> _salvedades;
        private Dictionary<int, EstadoConflicto> _estadoSesiones;
        private readonly Dictionary<int, List<ViolacionSesionInfo>> _violacionesPorSesion;
        private readonly List<ViolacionSesionInfo> _violacionesGlobales;
        private int _totalViolacionesHard;
        private int _sesionesConViolacionHard;
        private ContextMenuStrip _contextMenu;
        private TreeNode _nodoSeleccionadoContextual;
        private readonly ITrayectoEncoderService _trayectoEncoderService;

        private Font _boldFont;
        private Font _regularFont;
        private Font _monoFont;

        internal IReadOnlyDictionary<int, List<ViolacionSesionInfo>> ViolacionesPorSesion => _violacionesPorSesion;
        internal IReadOnlyList<ViolacionSesionInfo> ViolacionesGlobales => _violacionesGlobales;
        internal TreeView TreeVinculaciones => treeVinculaciones;

        public FormVerVinculaciones(List<SesionListItem> items, IDbConnection dbConnection, dynamic excelApp, AgModoDatos modo = null)
        {
            InitializeComponent();
            _items = items;
            _dbConnection = dbConnection;
            _excelApp = excelApp;
            _modo = modo ?? AgModoDatos.Determinista;
            _cacheTerceros = new Dictionary<string, Tercero>();

            _salvedadRepository = new SalvedadConflictoRepository(_dbConnection, modo: _modo);
            _salvedades = new List<SalvedadConflicto>().AsReadOnly();
            _estadoSesiones = new Dictionary<int, EstadoConflicto>();
            _violacionesPorSesion = new Dictionary<int, List<ViolacionSesionInfo>>();
            _violacionesGlobales = new List<ViolacionSesionInfo>();

            if (_modo.EsModoAg)
            {
                this.Text = $"{this.Text} [MODO AG - STAGING]";
            }

            try
            {
                var gastosRepo = new GastosTerrestresRepository(_dbConnection);
                var configRepo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.ConfigRepository(ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext());
                _trayectoEncoderService = new TrayectoEncoderService(gastosRepo, configRepo);
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[FormVerVinculaciones] No se pudo inicializar TrayectoEncoderService: {ex.Message}");
            }

            ConfigurarMenuContextual();

            // Configurar OwnerDraw para indentación personalizada
            treeVinculaciones.DrawMode = TreeViewDrawMode.OwnerDrawText;
            treeVinculaciones.DrawNode += TreeVinculaciones_DrawNode;
            treeVinculaciones.Indent = 25;   // Reducido de 35 a 25 para un look más compacto
            treeVinculaciones.ItemHeight = 18; // Reducido de 22 a 18 para un interlineado más justo

            _regularFont = new Font("Consolas", 9f);
            _boldFont = new Font(_regularFont, FontStyle.Bold);
            _monoFont = new Font("Consolas", 9f);
        }

        /// <summary>
        /// Carga asíncrona de terceros y construye el árbol inicial.
        /// <c>CargarTercerosAsync</c> garantiza no propagar excepciones (catch interno),
        /// por lo que <c>ConstruirArbol</c> siempre se ejecuta independientemente del resultado.
        /// </summary>
        protected override async void OnLoad(EventArgs e)
        {
            try
            {
                base.OnLoad(e);

                if (lblModo == null)
                {
                    lblModo = new Label
                    {
                        AutoSize = true,
                        Location = new Point(pnlBottom.Width - 220, 15),
                        Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                    };
                    pnlBottom.Controls.Add(lblModo);
                }
                lblModo.Text = _modo.EsModoAg ? "MODO AG (staging)" : "PRODUCCIÓN";
                lblModo.ForeColor = _modo.EsModoAg ? Color.OrangeRed : Color.DimGray;
                lblModo.Font = new Font(this.Font, FontStyle.Bold);
                this.Text = _modo.EsModoAg ? "Ver Vinculaciones — MODO AG" : "Ver Vinculaciones";

                // Fase 634 & 638: Capturar nombre del libro activo inicial liberando explícitamente el RCW
                dynamic wbInicial = null;
                try
                {
                    wbInicial = _excelApp != null ? _excelApp.ActiveWorkbook : null;
                    _nombreLibroInicial = wbInicial == null ? null : (string)wbInicial.Name;
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[FormVerVinculaciones] No se pudo capturar el nombre del libro inicial: {ex.Message}");
                }
                finally
                {
                    if (wbInicial != null) { try { Marshal.ReleaseComObject(wbInicial); } catch { } }
                }

                // PATRÓN ARQUITECTÓNICO: Ownership de recursos GDI
                // Se clona la fuente del TreeView para que este formulario controle el 
                // ciclo de vida completo de los objetos Font (ver método Dispose).
                // Nunca asignar _regularFont = treeVinculaciones.Font; directamente.
                try
                {
                    _regularFont = treeVinculaciones.Font != null ? (Font)treeVinculaciones.Font.Clone() : new Font("Consolas", 9f);
                    _boldFont = new Font(_regularFont, FontStyle.Bold);
                    _monoFont = new Font("Consolas", 9, FontStyle.Regular);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[FormVerVinculaciones] Error al inicializar fuentes GDI: {ex.Message}");
                    _regularFont = new Font("Consolas", 9f);
                    _boldFont = new Font("Consolas", 9f, FontStyle.Bold);
                    _monoFont = new Font("Consolas", 9f);
                }

                CargarTerceros(); // Llamada síncrona
                await CargarSalvedadesAsync();
                await CargarViolacionesAsync();
                ConstruirArbol();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Excepción no controlada en OnLoad", ex);
                MessageBox.Show("Ocurrió un error inesperado al cargar la vista.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task CargarSalvedadesAsync()
        {
            try
            {
                int idPeriodo = _items.FirstOrDefault()?.IdPeriodo ?? 0;
                if (idPeriodo > 0)
                {
                    _salvedades = await _salvedadRepository.GetByPeriodoAsync(idPeriodo);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error cargando salvedades", ex);
                _salvedades = new List<SalvedadConflicto>().AsReadOnly();
            }
        }

        private void ConfigurarMenuContextual()
        {
            _contextMenu = new ContextMenuStrip();
            var itemMarcar = new ToolStripMenuItem("Marcar conflicto como Viable (Salvedad)", null, async (s, e) => await MarcarConflictoViableAsync());
            var itemDesmarcar = new ToolStripMenuItem("Quitar salvedad de viabilidad", null, async (s, e) => await QuitarSalvedadAsync());
            var itemEditar = new ToolStripMenuItem("Editar Vinculación", null, (s, e) => EditarVinculacion());
            var itemImprimirViatico = new ToolStripMenuItem("Imprimir Viático", null, (s, e) => ImprimirViaticoIndividual());
            var itemPernocta = new ToolStripMenuItem("Pernocta Último Día", null, async (s, e) => await AlternarPernoctaUltimoDiaAsync());
            itemPernocta.CheckOnClick = false;

            var sepPool = new ToolStripSeparator();
            var itemHabilitarPool = new ToolStripMenuItem("Habilitar esta combinación en el Pool", null, async (s, e) => await HabilitarEnPoolAsync());
            var itemSyncSemestre = new ToolStripMenuItem("Sincronizar pool del semestre", null, async (s, e) => await SincronizarPoolSemestreAsync());
            var itemVerPoolDocente = new ToolStripMenuItem("Ver asignaturas habilitadas del docente...", null, (s, e) => VerAsignaturasHabilitadasDocente());

            _contextMenu.Items.Add(itemMarcar);
            _contextMenu.Items.Add(itemDesmarcar);
            _contextMenu.Items.Add(itemEditar);
            _contextMenu.Items.Add(itemImprimirViatico);
            _contextMenu.Items.Add(itemPernocta);
            _contextMenu.Items.Add(sepPool);
            _contextMenu.Items.Add(itemHabilitarPool);
            _contextMenu.Items.Add(itemSyncSemestre);
            _contextMenu.Items.Add(itemVerPoolDocente);

            treeVinculaciones.NodeMouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    treeVinculaciones.SelectedNode = e.Node;
                    _nodoSeleccionadoContextual = e.Node;

                    itemMarcar.Visible = false;
                    itemDesmarcar.Visible = false;
                    itemEditar.Visible = false;
                    itemImprimirViatico.Visible = false;
                    itemPernocta.Visible = false;
                    sepPool.Visible = false;
                    itemHabilitarPool.Visible = false;
                    itemSyncSemestre.Visible = false;
                    itemVerPoolDocente.Visible = false;

                    if (e.Node.Tag is SesionListItem sesion)
                    {
                        itemEditar.Visible = true;
                        sepPool.Visible = true;
                        itemHabilitarPool.Visible = true;
                        itemSyncSemestre.Visible = true;
                        itemVerPoolDocente.Visible = true;

                        // Solo mostrar las opciones de salvedades de viabilidad en los nodos hoja (sesiones individuales)
                        if (e.Node.Nodes.Count == 0 && _estadoSesiones.TryGetValue(sesion.IdSesion, out var estado))
                        {
                            if (estado == EstadoConflicto.RevisarViabilidad)
                            {
                                itemMarcar.Visible = true;
                            }
                            else if (estado == EstadoConflicto.Salvado)
                            {
                                itemDesmarcar.Visible = true;
                            }
                        }

                        _contextMenu.Show(treeVinculaciones, e.Location);
                    }
                    else if (e.Node.Tag is ViaticoBloque bloque)
                    {
                        itemImprimirViatico.Visible = true;

                        bool disponible = _vistaActual == VistaVinculaciones.Viaticos && !_modo.EsModoAg;
                        itemPernocta.Visible = disponible;
                        if (disponible)
                        {
                            itemPernocta.Checked = bloque.TienePernoctaUltimoDia;
                            itemPernocta.Text = bloque.TienePernoctaUltimoDia ? "🛏️ Quitar Pernocta Último Día" : "🛏️ Marcar Pernocta Último Día";
                        }

                        _contextMenu.Show(treeVinculaciones, e.Location);
                    }
                }
            };
        }

        private async Task AlternarPernoctaUltimoDiaAsync()
        {
            if (!(_nodoSeleccionadoContextual?.Tag is ViaticoBloque bloque) || bloque.Sesiones == null || !bloque.Sesiones.Any())
            {
                return;
            }

            var ids = bloque.IdsVinculacion;
            if (ids.Count == 0)
            {
                MessageBox.Show("No se encontraron vinculaciones válidas asociadas a este bloque.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int valorNuevo = bloque.TienePernoctaUltimoDia ? 0 : 1;
            string accionTexto = valorNuevo == 1 ? "marcar" : "quitar";
            double durActual = bloque.DuracionTotal;
            double durNueva = valorNuevo == 1 ? durActual + 1 : durActual - 1;
            string cambioDias = $"{durActual} ➔ {durNueva} días";

            // Mensaje informativo indicando impacto en las vinculaciones del bloque
            int totalSesionesAfectadas = _items.Count(i => ids.Contains(i.IdVinculacion));
            string mensaje = $"¿Desea {accionTexto} la pernocta en el último día para este bloque?\n\n" +
                             $"• Efecto: {cambioDias} (+1 día a la fecha de regreso).\n" +
                             $"• Vinculaciones afectadas: {ids.Count}.\n" +
                             $"• Sesiones recurrentes actualizadas: {totalSesionesAfectadas}.\n\n" +
                             $"Esta acción actualizará todas las sesiones asociadas a estas vinculaciones.";

            var confirm = MessageBox.Show(mensaje, "Confirmar Pernocta Último Día", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                string docenteSeleccionado = bloque.Sesiones.FirstOrDefault()?.DocenteNombre;
                DateTime? fechaMinima = bloque.FechaMinima;

                using (var tx = _dbConnection.BeginTransaction())
                {
                    var repo = new AcademicRepository(_dbConnection, tx, suppressMarkDirty: true, modo: _modo);
                    foreach (var id in ids)
                    {
                        var r = repo.SetPernoctaUltimoDia(id, valorNuevo);
                        if (r.IsFailure)
                        {
                            tx.Rollback();
                            MessageBox.Show($"Error al guardar pernocta en vinculación {id}: {r.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    tx.Commit();
                }

                ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                PersistirCambiosAlVault();

                // Actualizar en memoria la lista _items
                foreach (var item in _items.Where(i => ids.Contains(i.IdVinculacion)))
                {
                    item.IsPernoctaUltimoDia = valorNuevo;
                }

                // Reconstruir árbol
                ConstruirArbol();

                // Intentar restaurar selección en el árbol por coincidencia exacta
                if (!string.IsNullOrEmpty(docenteSeleccionado) && fechaMinima.HasValue)
                {
                    foreach (TreeNode nDoc in treeVinculaciones.Nodes)
                    {
                        foreach (TreeNode nBloque in nDoc.Nodes)
                        {
                            if (nBloque.Tag is ViaticoBloque b
                                && b.FechaMinima == fechaMinima
                                && b.Sesiones.Any(s => string.Equals(s.DocenteNombre, docenteSeleccionado, StringComparison.Ordinal)))
                            {
                                nDoc.Expand();
                                treeVinculaciones.SelectedNode = nBloque;
                                nBloque.EnsureVisible();
                                return;
                            }
                        }
                    }
                }

                LoggerHelper.Info($"[FormVerVinculaciones] Pernocta último día actualizada a {valorNuevo} para vinculaciones [{string.Join(", ", ids)}]");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error en AlternarPernoctaUltimoDiaAsync", ex);
                MessageBox.Show($"Ocurrió un error inesperado al actualizar la pernocta: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task HabilitarEnPoolAsync()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesion)
            {
                try
                {
                    var poolRepo = new TerceroAsignaturaRepository(() => _dbConnection);
                    long id = TerceroAsignaturaKey.Calcular(sesion.IdTercero, sesion.IdPrograma, sesion.IdAsignatura, sesion.Semestre);
                    var existenteResult = await poolRepo.GetByIdAsync(id, System.Threading.CancellationToken.None);

                    var item = existenteResult.IsSuccess ? existenteResult.Value : new TerceroAsignatura
                    {
                        Id = id,
                        IdTercero = sesion.IdTercero,
                        IdPrograma = sesion.IdPrograma,
                        IdAsignatura = sesion.IdAsignatura,
                        Semestre = sesion.Semestre,
                        Activo = 1,
                        Prioridad = 0,
                        Origen = "MANUAL",
                        Subunidad = sesion.Subunidad ?? string.Empty
                    };
                    item.Activo = 1;

                    var res = await poolRepo.UpsertAsync(item, System.Threading.CancellationToken.None);
                    if (res.IsSuccess)
                    {
                        MessageBox.Show($"Docente habilitado exitosamente en el pool para la asignatura '{sesion.AsignaturaNombre}'.", "Pool Docente", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"Error al habilitar en el pool: {res.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[FormVerVinculaciones] Error al habilitar en pool", ex);
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task SincronizarPoolSemestreAsync()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesion)
            {
                try
                {
                    var poolRepo = new TerceroAsignaturaRepository(() => _dbConnection);
                    var syncService = new TerceroAsignaturaSyncService(() => _dbConnection, poolRepo);
                    var res = await syncService.SincronizarPorSemestreAsync(sesion.IdPrograma, sesion.Semestre, sesion.Subunidad, System.Threading.CancellationToken.None);
                    if (res.IsSuccess)
                    {
                        MessageBox.Show($"Sincronización completada. Se actualizaron {res.Value} registros del pool para el semestre '{sesion.Semestre}'.", "Sincronización Pool", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"Error en sincronización: {res.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[FormVerVinculaciones] Error al sincronizar pool de semestre", ex);
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void VerAsignaturasHabilitadasDocente()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesion)
            {
                using (var form = new FormTerceros(() => _dbConnection))
                {
                    form.ShowDialog(this);
                }
            }
        }

        private void EditarVinculacion()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesion)
            {
                using (var form = new FormSesionesAcademicas(() => _dbConnection, null, sesion.IdVinculacion, _modo))
                {
                    form.ShowDialog(this);
                    if (form.DatosGuardados)
                    {
                        // Actualizar la lista en memoria
                        var nuevasSesiones = _dbConnection.Query<SesionListItem>(
                            $"SELECT * FROM {_modo.VistaSesionesDetalle} WHERE IdVinculacion = @IdVinculacion",
                            new { IdVinculacion = sesion.IdVinculacion }).ToList();

                        _items.RemoveAll(x => x.IdVinculacion == sesion.IdVinculacion);
                        if (nuevasSesiones.Any())
                        {
                            _items.AddRange(nuevasSesiones);
                        }
                        ConstruirArbol();
                    }
                }
            }
        }

        private void ImprimirViaticoIndividual()
        {
            if (_nodoSeleccionadoContextual?.Tag is ViaticoBloque bloque)
            {
                using (var formCdp = new FormCdpToPDF())
                {
                    var result = formCdp.ShowDialog(this);
                    if (result == DialogResult.OK)
                    {
                        ExportarBloquesViaticos(new List<ViaticoBloque> { bloque });
                    }
                }
            }
        }

        private async Task MarcarConflictoViableAsync()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesionA)
            {
                var sesionesDelDia = _items.Where(s => s.IdTercero == sesionA.IdTercero && s.FechaSesion?.Date == sesionA.FechaSesion?.Date).ToList();
                var candidatas = sesionesDelDia
                    .Where(s => s.IdSesion != sesionA.IdSesion && !string.Equals(s.SedeCiudad ?? "", sesionA.SedeCiudad ?? "", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (candidatas.Count == 0) return;

                bool cambiosRealizados = false;

                foreach (var sesionB in candidatas)
                {
                    string msg = $"¿Confirmar que el traslado entre:\n\n📍 {sesionA.SedeCiudad} ({sesionA.HorarioSesion})\n📍 {sesionB.SedeCiudad} ({sesionB.HorarioSesion})\n\n...es viable para este docente?";
                    var resultDialog = MessageBox.Show(msg, "Confirmar Viabilidad", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (resultDialog == DialogResult.Cancel)
                        break;

                    if (resultDialog == DialogResult.Yes)
                    {
                        var salvedad = new SalvedadConflicto
                        {
                            IdSesionA = Math.Min(sesionA.IdSesion, sesionB.IdSesion),
                            IdSesionB = Math.Max(sesionA.IdSesion, sesionB.IdSesion),
                            IdTercero = sesionA.IdTercero,
                            IdPeriodo = sesionA.IdPeriodo,
                            Justificacion = "Confirmado manualmente por el usuario."
                        };

                        var result = await _salvedadRepository.SaveAsync(salvedad);
                        if (result.IsSuccess)
                        {
                            cambiosRealizados = true;
                        }
                        else
                        {
                            MessageBox.Show($"Error al guardar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                if (cambiosRealizados)
                {
                    PersistirCambiosAlVault();
                    await CargarSalvedadesAsync();
                    await CargarViolacionesAsync();
                    ConstruirArbol();
                }
            }
        }

        private async Task QuitarSalvedadAsync()
        {
            if (_nodoSeleccionadoContextual?.Tag is SesionListItem sesionA)
            {
                var sesionesDelDia = _items.Where(s => s.IdTercero == sesionA.IdTercero && s.FechaSesion?.Date == sesionA.FechaSesion?.Date).ToList();
                var candidatas = sesionesDelDia
                    .Where(s => s.IdSesion != sesionA.IdSesion && !string.Equals(s.SedeCiudad ?? "", sesionA.SedeCiudad ?? "", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (candidatas.Count == 0) return;

                bool cambiosRealizados = false;

                foreach (var sesionB in candidatas)
                {
                    string msg = $"¿Desea quitar la confirmación de viabilidad entre:\n\n📍 {sesionA.SedeCiudad} ({sesionA.HorarioSesion})\n📍 {sesionB.SedeCiudad} ({sesionB.HorarioSesion})?";
                    var resultDialog = MessageBox.Show(msg, "Quitar Salvedad", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (resultDialog == DialogResult.Cancel)
                        break;

                    if (resultDialog == DialogResult.Yes)
                    {
                        var result = await _salvedadRepository.DeleteAsync(sesionA.IdSesion, sesionB.IdSesion);
                        if (result.IsSuccess)
                        {
                            cambiosRealizados = true;
                        }
                        else
                        {
                            MessageBox.Show($"Error al eliminar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                if (cambiosRealizados)
                {
                    PersistirCambiosAlVault();
                    await CargarSalvedadesAsync();
                    await CargarViolacionesAsync();
                    ConstruirArbol();
                }
            }
        }

        private static System.Data.SQLite.SQLiteConnection UnwrapSqlite(IDbConnection conn)
        {
            while (conn is ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context.NonDisposableDbConnection ndc)
                conn = ndc.Unwrap();
            return conn as System.Data.SQLite.SQLiteConnection;
        }

        /// <summary>
        /// Parsea de forma robusta y tipada los intervalos de horario (ej: "14:00 A 18:00", "08:00-12:00").
        /// </summary>
        private static bool TryParseIntervalo(string texto, out TimeSpan ini, out TimeSpan fin)
        {
            ini = TimeSpan.Zero;
            fin = TimeSpan.Zero;
            if (string.IsNullOrWhiteSpace(texto)) return false;
            string[] seps = new string[] { " A ", " a ", " - ", "-" };
            foreach (var sep in seps)
            {
                int idx = texto.IndexOf(sep, StringComparison.OrdinalIgnoreCase);
                if (idx > 0)
                {
                    string s1 = texto.Substring(0, idx).Trim();
                    string s2 = texto.Substring(idx + sep.Length).Trim();
                    if (TimeSpan.TryParse(s1, out ini) && TimeSpan.TryParse(s2, out fin))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void PersistirCambiosAlVault()
        {
            try
            {
                System.Data.SQLite.SQLiteConnection sqliteConn = UnwrapSqlite(_dbConnection);

                if (sqliteConn != null)
                {
                    string tempDbPath = sqliteConn.FileName;

                    // 1. Flush de memoria/WAL hacia el archivo físico .db
                    using (var cmd = new System.Data.SQLite.SQLiteCommand("PRAGMA wal_checkpoint(TRUNCATE);", sqliteConn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 2. La persistencia al Vault ahora la gestiona el SyncScheduler
                    LoggerHelper.Info("[FormVerVinculaciones] Cambios de salvedades guardados. Persistencia asíncrona delegada al Scheduler.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error al persistir salvedades al Vault.", ex);
            }
        }

        private void CargarTerceros()
        {
            try
            {
                // M6: Uso de repositorio para desacoplar UI de Dapper
                var repo = new AcademicRepository(_dbConnection);
                var result = repo.GetTerceros(soloActivos: false);
                if (result.IsSuccess && result.Value != null)
                {
                    _cacheTerceros = result.Value.ToDictionary(t => t.Cedula, t => t);
                }
            }
            catch (Exception ex)
            {
                // La caché queda vacía; el árbol renderiza con DocenteNombre como fallback
                LoggerHelper.Error("[FormVerVinculaciones] Error cargando terceros para caché", ex);
            }
        }

        internal void ConstruirArbol()
        {
            if (_vistaActual == VistaVinculaciones.PorAsignaturas)
                ConstruirArbolPorAsignaturas();
            else if (_vistaActual == VistaVinculaciones.PorSemestres)
                ConstruirArbolPorSemestres();
            else if (_vistaActual == VistaVinculaciones.PorFechas)
                ConstruirArbolPorFechas();
            else if (_vistaActual == VistaVinculaciones.Viaticos)
                ConstruirArbolViaticos();

            AgregarNodoViolacionesGlobales();

            // M3: Actualizar estado visual de botones de vista
            ActualizarEstadoBotones();
        }

        private void TreeVinculaciones_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            try
            {
                // Ocultar el dibujado por defecto del texto
                e.DrawDefault = false;

                // 1. Dibujar el fondo (soporta selección y colores de conflicto)
                bool isSelected = (e.State & TreeNodeStates.Selected) != 0;
                if (isSelected)
                {
                    e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds);
                }
                else if (e.Node.BackColor != Color.Empty)
                {
                    using (SolidBrush bgBrush = new SolidBrush(e.Node.BackColor))
                        e.Graphics.FillRectangle(bgBrush, e.Bounds);
                }
                else
                {
                    e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
                }

                // 2. Determinar fuente y color
                // Si es nivel de detalle (Level > 0), usamos fuente monoespaciada para alineación columnar
                Font font = e.Node.Level > 0 ? _monoFont : (e.Node.NodeFont ?? e.Node.TreeView.Font);

                Color textColor = isSelected ? SystemColors.HighlightText
                                : (e.Node.ForeColor != Color.Empty ? e.Node.ForeColor : e.Node.TreeView.ForeColor);

                // 3. Aplicar suavizado de texto premium
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // 4. Calcular indentación extra más justa (Nivel 1 y superiores)
                int extraIndent = (e.Node.Level > 0) ? 10 : 0; // Reducido de 30 a 10 para mayor compacidad

                int textX = e.Bounds.X + extraIndent;
                int textY = e.Bounds.Y + (e.Bounds.Height - font.Height) / 2;

                // 5. Dibujar el texto respetando la indentación
                TextRenderer.DrawText(e.Graphics, e.Node.Text, font, new Point(textX, textY), textColor, TextFormatFlags.Left);
            }
            catch (ArgumentException)
            {
                // Ignorar error si el objeto GDI fue liberado durante el cierre (Add-in Excel COM)
                return;
            }
            catch
            {
                // Ignorar otros errores de renderizado para evitar crashes en Excel
                return;
            }
        }

        private void ConstruirArbolPorAsignaturas()
        {
            treeVinculaciones.Nodes.Clear();
            treeVinculaciones.BeginUpdate();

            var docentes = _items.GroupBy(s => s.DocenteNombre).OrderBy(g => g.Key);

            foreach (var grupoDocente in docentes)
            {
                double horasTotalesDocente = Math.Round(grupoDocente.Sum(s => s.Horas), 1);
                var nodoDocente = new TreeNode($"👤 {grupoDocente.Key} | {horasTotalesDocente}h")
                {
                    ForeColor = Color.DarkGreen,
                    NodeFont = _boldFont
                };

                var asignaturas = grupoDocente
                    .GroupBy(s => new { s.SedeCiudad, s.AsignaturaNombre, s.Semestre })
                    .OrderBy(g => g.Key.SedeCiudad)
                    .ThenBy(g => g.Key.Semestre)
                    .ThenBy(g => g.Key.AsignaturaNombre);

                foreach (var grupoAsignatura in asignaturas)
                {
                    double horasSubtotal = Math.Round(grupoAsignatura.Sum(s => s.Horas), 1);
                    string sede = grupoAsignatura.Key.SedeCiudad ?? "SIN SEDE";

                    // Calcular rango de fechas (mínima y máxima) basándose en las sesiones del Nivel 3
                    var fechasValidas = grupoAsignatura.Where(s => s.FechaSesion.HasValue).Select(s => s.FechaSesion.Value).ToList();
                    string rangoFechas = "S/F | S/F";
                    if (fechasValidas.Any())
                    {
                        string minDate = fechasValidas.Min().ToString("dd/MM/yyyy");
                        string maxDate = fechasValidas.Max().ToString("dd/MM/yyyy");
                        rangoFechas = $"{minDate} | {maxDate}";
                    }

                    var nodoAsignatura = new TreeNode(
                        $"📍 {rangoFechas} | {sede} | {horasSubtotal}h | {grupoAsignatura.Key.Semestre} | {grupoAsignatura.Key.AsignaturaNombre}")
                    {
                        ForeColor = Color.DarkBlue,
                        NodeFont = _boldFont,
                        Tag = grupoAsignatura.FirstOrDefault() // Permite editar vinculación desde el nivel 2
                    };

                    foreach (var sesion in grupoAsignatura.OrderBy(s => s.FechaSesion))
                    {
                        nodoAsignatura.Nodes.Add(CrearNodoSesion(sesion));
                    }

                    nodoDocente.Nodes.Add(nodoAsignatura);
                }

                treeVinculaciones.Nodes.Add(nodoDocente);
            }

            treeVinculaciones.EndUpdate();
            lblTotal.Text = $"{_items.Count} sesiones · Vista: ASIGNATURAS";
            lblTotal.ForeColor = Color.Black;
            ExpandirHastaNivel(1);
            AgregarNodoViolacionesGlobales();
        }

        private void ConstruirArbolPorSemestres()
        {
            treeVinculaciones.Nodes.Clear();
            treeVinculaciones.BeginUpdate();

            var gruposNivel1 = _items
                .GroupBy(s => new { s.ProgramaDetalle, s.Semestre, s.AsignaturaNombre })
                .OrderBy(g => g.Key.ProgramaDetalle)
                .ThenBy(g => SemestreParser.Parsear(g.Key.Semestre))
                .ThenBy(g => g.Key.Semestre)
                .ThenBy(g => g.Key.AsignaturaNombre);

            foreach (var g1 in gruposNivel1)
            {
                double horasNivel1 = Math.Round(g1.Sum(s => s.Horas), 1);
                string tituloNivel1 = $"{g1.Key.ProgramaDetalle} | Semestre {g1.Key.Semestre} | {horasNivel1}h | {g1.Key.AsignaturaNombre}";

                var nodoNivel1 = new TreeNode(tituloNivel1)
                {
                    ForeColor = Color.DarkMagenta,
                    NodeFont = _boldFont
                };

                var gruposNivel2 = g1
                    .GroupBy(s => new { s.SedeCiudad, s.DocenteNombre })
                    .OrderBy(g => g.Key.SedeCiudad)
                    .ThenBy(g => g.Key.DocenteNombre);

                foreach (var g2 in gruposNivel2)
                {
                    var fechasValidas = g2.Where(s => s.FechaSesion.HasValue).Select(s => s.FechaSesion.Value).ToList();
                    string minDate = fechasValidas.Any() ? fechasValidas.Min().ToString("dd/MM/yyyy") : "S/F";
                    string maxDate = fechasValidas.Any() ? fechasValidas.Max().ToString("dd/MM/yyyy") : "S/F";

                    string sedeText = g2.Key.SedeCiudad ?? "SIN SEDE";
                    string docenteText = g2.Key.DocenteNombre ?? "SIN DOCENTE";

                    var nodoNivel2 = new TreeNode($"📍 {sedeText} | 👤 {docenteText} | {minDate} | {maxDate}")
                    {
                        ForeColor = Color.DarkBlue,
                        NodeFont = _boldFont,
                        Tag = g2.FirstOrDefault() // Permite editar vinculación desde el nivel 2
                    };

                    foreach (var sesion in g2.OrderBy(s => s.FechaSesion))
                    {
                        nodoNivel2.Nodes.Add(CrearNodoSesion(sesion));
                    }

                    nodoNivel1.Nodes.Add(nodoNivel2);
                }

                treeVinculaciones.Nodes.Add(nodoNivel1);
            }

            treeVinculaciones.EndUpdate();
            lblTotal.Text = $"{_items.Count} sesiones · Vista: SEMESTRES";
            lblTotal.ForeColor = Color.Black;
            ExpandirHastaNivel(1);
            AgregarNodoViolacionesGlobales();
        }

        /// <summary>
        /// Carga el detalle persistido en ag_violacion para la corrida mostrada y lo indexa por sesión
        /// o como violaciones generales si no tienen id_gen o sesión asociada (Fase 1812).
        /// Solo aplica en Modo AG (staging) y soporta múltiples corridas (IN @IdCorridas).
        /// </summary>
        internal async Task CargarViolacionesAsync()
        {
            _violacionesPorSesion.Clear();
            _violacionesGlobales.Clear();
            _totalViolacionesHard = 0;
            _sesionesConViolacionHard = 0;

            // En modo Producción no existe ag_violacion ni staging de AG
            if (!_modo.EsModoAg) return;

            try
            {
                var idsCorrida = _items.Select(i => i.IdCorrida).Where(id => id > 0).Distinct().ToList();
                if (idsCorrida.Count == 0) return;

                const string sql = @"
SELECT v.id AS Id, v.id_individuo AS IdIndividuo, v.id_gen AS IdGen,
       CAST(v.regla AS TEXT) AS Regla, v.es_dura AS EsDura,
       v.mensaje AS Mensaje, v.detalle_json AS DetalleJson,
       c.id AS IdCorrida
FROM ag_corrida c
JOIN ag_individuo i ON i.id = c.id_individuo_optimo AND i.id_corrida = c.id
JOIN ag_violacion v ON v.id_individuo = i.id
WHERE c.id IN @IdCorridas AND (@Subunidad = '' OR COALESCE(c.subunidad, '') IN ('', @Subunidad))
ORDER BY v.es_dura DESC, v.regla, v.id;";

                var conn = UnwrapSqlite(_dbConnection) ?? _dbConnection;
                if (conn == null) return;

                string subunidad = _items.Select(i => i.Subunidad).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? string.Empty;

                var filas = (await conn.QueryAsync(new CommandDefinition(sql, new { IdCorridas = idsCorrida, Subunidad = subunidad }))).ToList();
                if (filas.Count == 0)
                {
                    string corridasTexto = string.Join(", ", idsCorrida);
                    LoggerHelper.Info($"[FormVerVinculaciones] Corrida(s) #{corridasTexto} sin detalle en ag_violacion; use «Recalcular violaciones» en el Módulo R.");
                    return;
                }

                foreach (var fila in filas)
                {
                    object idGenRaw = fila.IdGen;
                    int? idGen = (idGenRaw != null && !(idGenRaw is DBNull)) ? Convert.ToInt32(idGenRaw) : (int?)null;

                    object esDuraRaw = fila.EsDura;
                    bool esDura = esDuraRaw != null && !(esDuraRaw is DBNull) && Convert.ToInt32(esDuraRaw) == 1;

                    object reglaRaw = fila.Regla;
                    string regla = reglaRaw != null && !(reglaRaw is DBNull) ? reglaRaw.ToString() : "?";

                    object mensajeRaw = fila.Mensaje;
                    string mensaje = mensajeRaw != null && !(mensajeRaw is DBNull) ? mensajeRaw.ToString() : string.Empty;

                    object detalleRaw = fila.DetalleJson;
                    string detalleJson = detalleRaw != null && !(detalleRaw is DBNull) ? detalleRaw.ToString() : null;

                    object idRaw = fila.Id;
                    long id = idRaw != null && !(idRaw is DBNull) ? Convert.ToInt64(idRaw) : 0;

                    object idIndividuoRaw = fila.IdIndividuo;
                    long idIndividuo = idIndividuoRaw != null && !(idIndividuoRaw is DBNull) ? Convert.ToInt64(idIndividuoRaw) : 0;

                    object idCorridaRaw = fila.IdCorrida;
                    long idCorridaFila = idCorridaRaw != null && !(idCorridaRaw is DBNull) ? Convert.ToInt64(idCorridaRaw) : 0;

                    var info = new ViolacionSesionInfo
                    {
                        Id = id,
                        IdIndividuo = idIndividuo,
                        IdGen = idGen,
                        Regla = regla,
                        EsDura = esDura,
                        Mensaje = mensaje,
                        DetalleJson = detalleJson,
                        IdCorrida = idCorridaFila
                    };

                    if (esDura)
                    {
                        _totalViolacionesHard++;
                    }

                    var idsGenes = new HashSet<int>();
                    if (idGen.HasValue && idGen.Value > 0)
                    {
                        idsGenes.Add((int)idGen.Value);
                    }

                    DateTime? fechaFiltro = null;
                    var intervalosFiltro = new List<(TimeSpan Inicio, TimeSpan Fin)>();

                    if (!string.IsNullOrWhiteSpace(detalleJson))
                    {
                        try
                        {
                            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            var det = serializer.DeserializeObject(detalleJson) as Dictionary<string, object>;
                            if (det != null)
                            {
                                if (det.TryGetValue("idGenes", out var valIdGenes) && valIdGenes != null)
                                {
                                    var arr = valIdGenes as object[] ?? (valIdGenes as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                                    if (arr != null)
                                    {
                                        foreach (var item in arr)
                                        {
                                            if (int.TryParse(item?.ToString(), out int gId) && gId > 0)
                                            {
                                                idsGenes.Add(gId);
                                            }
                                        }
                                    }
                                }

                                if (det.TryGetValue("fecha", out var valFecha) && valFecha != null)
                                {
                                    if (DateTime.TryParse(valFecha.ToString(), out DateTime fParsed))
                                    {
                                        fechaFiltro = fParsed.Date;
                                    }
                                }

                                if (det.TryGetValue("intervalos", out var valIntervalos) && valIntervalos != null)
                                {
                                    var arrInt = valIntervalos as object[] ?? (valIntervalos as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                                    if (arrInt != null)
                                    {
                                        foreach (var elem in arrInt)
                                        {
                                            if (elem is Dictionary<string, object> dictInt)
                                            {
                                                if (dictInt.TryGetValue("inicio", out var ini) && ini != null &&
                                                    dictInt.TryGetValue("fin", out var fin) && fin != null)
                                                {
                                                    if (TimeSpan.TryParse(ini.ToString(), out TimeSpan tIni) &&
                                                        TimeSpan.TryParse(fin.ToString(), out TimeSpan tFin))
                                                    {
                                                        intervalosFiltro.Add((tIni, tFin));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // DetalleJson no parseable o sin campos esperados
                        }
                    }

                    List<SesionListItem> sesionesAsociadas = null;
                    if (idsGenes.Count > 0)
                    {
                        var candidatas = _items.Where(s => idsGenes.Contains(s.IdGen)).ToList();

                        if (candidatas.Count == 0)
                        {
                            try
                            {
                                var idsSesionDb = (await conn.QueryAsync<int>(
                                    "SELECT id FROM ag_sesiones WHERE idgen IN @IdsGenes;",
                                    new { IdsGenes = idsGenes.ToList() })).ToList();

                                if (idsSesionDb.Count > 0)
                                {
                                    candidatas = _items.Where(s => idsSesionDb.Contains(s.IdSesion)).ToList();
                                }
                            }
                            catch
                            {
                                // Si ag_sesiones no existe o no tiene registros
                            }
                        }

                        // Fase 1815: Asociar ambos genes únicamente en la fecha e intervalo tipados afectados.
                        // La ausencia de coincidencia exacta se clasifica como violación global, evitando falsos positivos en el resto del semestre.
                        if (fechaFiltro.HasValue)
                        {
                            var porFecha = candidatas.Where(s => s.FechaSesion.HasValue && s.FechaSesion.Value.Date == fechaFiltro.Value).ToList();
                            if (porFecha.Count > 0)
                            {
                                if (intervalosFiltro.Count > 0)
                                {
                                    var porIntervalo = porFecha.Where(s =>
                                    {
                                        if (TryParseIntervalo(s.HorarioSesion, out var iniSes, out var finSes))
                                        {
                                            return intervalosFiltro.Any(inter => iniSes < inter.Fin && finSes > inter.Inicio);
                                        }
                                        return false;
                                    }).ToList();

                                    sesionesAsociadas = porIntervalo.Count > 0 ? porIntervalo : null;
                                }
                                else
                                {
                                    sesionesAsociadas = porFecha;
                                }
                            }
                            else
                            {
                                sesionesAsociadas = null;
                            }
                        }
                        else if (intervalosFiltro.Count > 0)
                        {
                            var porIntervalo = candidatas.Where(s =>
                            {
                                if (TryParseIntervalo(s.HorarioSesion, out var iniSes, out var finSes))
                                {
                                    return intervalosFiltro.Any(inter => iniSes < inter.Fin && finSes > inter.Inicio);
                                }
                                return false;
                            }).ToList();

                            sesionesAsociadas = porIntervalo.Count > 0 ? porIntervalo : null;
                        }
                        else
                        {
                            sesionesAsociadas = candidatas.Count > 0 ? candidatas : null;
                        }
                    }

                    if (sesionesAsociadas != null && sesionesAsociadas.Count > 0)
                    {
                        foreach (var sesion in sesionesAsociadas)
                        {
                            if (!_violacionesPorSesion.TryGetValue(sesion.IdSesion, out var lista))
                            {
                                lista = new List<ViolacionSesionInfo>();
                                _violacionesPorSesion[sesion.IdSesion] = lista;
                            }
                            lista.Add(info);
                        }
                    }
                    else
                    {
                        _violacionesGlobales.Add(info);
                    }
                }

                _sesionesConViolacionHard = _violacionesPorSesion.Count(kv => kv.Value.Any(v => v.EsDura));
                string corridasReporte = string.Join(", ", idsCorrida);
                LoggerHelper.Info($"[FormVerVinculaciones] Corrida(s) #{corridasReporte}: {filas.Count} violaciones ({_totalViolacionesHard} HARD, {_violacionesGlobales.Count} globales) en {_sesionesConViolacionHard} sesiones.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[FormVerVinculaciones] No se pudo cargar el detalle de violaciones: {ex.Message}");
            }
        }

        /// <summary>Detalle de una violación asociada a una sesión o corrida (proyección de ag_violacion).</summary>
        internal sealed class ViolacionSesionInfo
        {
            public long Id { get; set; }
            public long IdIndividuo { get; set; }
            public int? IdGen { get; set; }
            public string Regla { get; set; }
            public bool EsDura { get; set; }
            public string Mensaje { get; set; }
            public string DetalleJson { get; set; }
            public long IdCorrida { get; set; }
        }

        internal void ConstruirArbolPorFechas()
        {
            treeVinculaciones.Nodes.Clear();
            treeVinculaciones.BeginUpdate();

            _estadoSesiones.Clear();

            int docentesConConflictoReal = 0;
            int docentesParaRevisar = 0;
            int totalSesionesConflicto = 0;
            int totalSesionesRevisar = 0;

            var docentes = _items.GroupBy(s => s.DocenteNombre).OrderBy(g => g.Key);

            foreach (var grupoDocente in docentes)
            {
                double horasTotalesDocente = Math.Round(grupoDocente.Sum(s => s.Horas), 1);
                var sesiones = grupoDocente.ToList();

                var analisis = SesionConflictAnalyzer.Analizar(sesiones, _salvedades);
                foreach (var kvp in analisis) _estadoSesiones[kvp.Key] = kvp.Value;

                bool tieneConflictoReal = analisis.Values.Any(v => v == EstadoConflicto.ConflictoReal);
                bool tieneParaRevisar = analisis.Values.Any(v => v == EstadoConflicto.RevisarViabilidad);

                if (tieneConflictoReal) docentesConConflictoReal++;
                else if (tieneParaRevisar) docentesParaRevisar++;

                int countConflicto = analisis.Values.Count(v => v == EstadoConflicto.ConflictoReal);
                int countRevisar = analisis.Values.Count(v => v == EstadoConflicto.RevisarViabilidad);

                totalSesionesConflicto += countConflicto;
                totalSesionesRevisar += countRevisar;

                string sufijoDocente = "";
                Color colorDocente = Color.DarkGreen;

                if (tieneConflictoReal)
                {
                    sufijoDocente = $" — ⚠️ {countConflicto} conflictos de horario";
                    colorDocente = Color.DarkRed;
                }
                else if (tieneParaRevisar)
                {
                    sufijoDocente = $" — ⚡ {countRevisar} requieren revisión de viabilidad";
                    colorDocente = Color.DarkGoldenrod;
                }

                string textoDocente = $"👤 {grupoDocente.Key} | {horasTotalesDocente}h{sufijoDocente}";

                var nodoDocente = new TreeNode(textoDocente)
                {
                    ForeColor = colorDocente,
                    NodeFont = _boldFont
                };

                foreach (var sesion in sesiones.OrderBy(s => s.FechaSesion))
                {
                    if (!sesion.FechaSesion.HasValue) continue;

                    var estado = analisis.ContainsKey(sesion.IdSesion) ? analisis[sesion.IdSesion] : EstadoConflicto.Ninguno;

                    Color backColor = SystemColors.Window;
                    Color foreColor = SystemColors.ControlText;
                    string prefijo = "";
                    bool esPlataforma = ModalidadTutoriaCatalog.EsPlataforma(sesion.Modalidad);

                    if (esPlataforma)
                    {
                        backColor = Color.LightCyan;
                        foreColor = Color.DarkBlue;
                    }

                    switch (estado)
                    {
                        case EstadoConflicto.ConflictoReal:
                            backColor = Color.MistyRose;
                            foreColor = Color.DarkRed;
                            prefijo = "⚠️ ";
                            break;
                        case EstadoConflicto.RevisarViabilidad:
                            backColor = Color.LightYellow;
                            foreColor = Color.DarkGoldenrod;
                            prefijo = "⚡ ";
                            break;
                        case EstadoConflicto.Salvado:
                            backColor = Color.Honeydew;
                            foreColor = Color.DarkGreen;
                            prefijo = "✓ ";
                            break;
                    }

                    var violacionesSesion = _violacionesPorSesion.TryGetValue(sesion.IdSesion, out var vlSes) ? vlSes : null;
                    bool tieneViolacionHard = violacionesSesion != null && violacionesSesion.Any(v => v.EsDura);
                    if (tieneViolacionHard)
                    {
                        backColor = Color.FromArgb(255, 205, 210);
                        foreColor = Color.Firebrick;
                        prefijo = "⛔ ";
                    }

                    string sufijoViolacion = tieneViolacionHard
                        ? " | ⛔ " + string.Join(", ", violacionesSesion.Where(v => v.EsDura).Select(v => v.Regla).Distinct())
                        : "";

                    string fecha = sesion.FechaSesion.Value.ToString("dd/MM/yyyy");
                    string diaSemana = (sesion.DiaSemana?.ToUpper() ?? "S/D").PadRight(9);
                    string sede = (sesion.SedeCiudad?.ToUpper() ?? "SIN SEDE").PadRight(20);
                    string semestre = (sesion.Semestre ?? "S/S").PadRight(5);
                    string sufijoPlataforma = esPlataforma ? " | 🌐 Plataforma" : "";

                    var nodoSesion = new TreeNode(
                        $"{prefijo}• {fecha} | {sesion.Horas}h | {diaSemana} | {sede} | {semestre} | {sesion.AsignaturaNombre} | {sesion.HorarioSesion}{sufijoPlataforma}{sufijoViolacion}")
                    {
                        ForeColor = foreColor,
                        BackColor = backColor,
                        NodeFont = (estado == EstadoConflicto.ConflictoReal || estado == EstadoConflicto.RevisarViabilidad || tieneViolacionHard) ? _boldFont : _regularFont,
                        Tag = sesion
                    };

                    if (estado == EstadoConflicto.ConflictoReal || estado == EstadoConflicto.RevisarViabilidad)
                    {
                        nodoSesion.ToolTipText = ObtenerDetalleCruceSesion(sesion, sesiones, estado);
                    }
                    else if (estado == EstadoConflicto.Salvado)
                    {
                        nodoSesion.ToolTipText = "✅ Viabilidad de traslado confirmada manualmente.";
                    }

                    if (tieneViolacionHard)
                    {
                        string detalleViolacion = string.Join("\n", violacionesSesion.Select(v =>
                            $"{(v.EsDura ? "HARD" : "SOFT")} {v.Regla}: {v.Mensaje}"));
                        nodoSesion.ToolTipText = string.IsNullOrWhiteSpace(nodoSesion.ToolTipText)
                            ? detalleViolacion
                            : detalleViolacion + "\n\n" + nodoSesion.ToolTipText;
                    }

                    nodoDocente.Nodes.Add(nodoSesion);
                }

                treeVinculaciones.Nodes.Add(nodoDocente);
            }

            treeVinculaciones.EndUpdate();

            // Se acortan los textos de estado para evitar superposiciones con los botones
            // en resoluciones pequeñas o cuando el nombre de las etiquetas es extenso.
            if (docentesConConflictoReal > 0)
            {
                lblTotal.Text = $"⚠️ Conflictos: {docentesConConflictoReal} Docentes / {totalSesionesConflicto} sesiones";
                lblTotal.ForeColor = Color.DarkRed;
            }
            else if (docentesParaRevisar > 0)
            {
                lblTotal.Text = $"⚡ Revisar: {docentesParaRevisar} Docentes / {totalSesionesRevisar} sesiones";
                lblTotal.ForeColor = Color.DarkGoldenrod;
            }
            else
            {
                lblTotal.Text = $"✓ Sin conflictos";
                lblTotal.ForeColor = Color.DarkGreen;
            }
            if (_totalViolacionesHard > 0)
            {
                lblTotal.Text = $"⛔ Violaciones HARD: {_totalViolacionesHard} en {_sesionesConViolacionHard} sesiones | " + lblTotal.Text;
                lblTotal.ForeColor = Color.DarkRed;
            }
            ExpandirHastaNivel(0);
            AgregarNodoViolacionesGlobales();
        }

        private void ConstruirArbolViaticos()
        {
            treeVinculaciones.Nodes.Clear();
            treeVinculaciones.BeginUpdate();

            var docentes = _items.GroupBy(s => s.DocenteNombre).OrderBy(g => g.Key);
            int totalViajes = 0;

            foreach (var grupoDocente in docentes)
            {
                var sesionesGlobalesDocente = grupoDocente.Where(s => s.Viatica).OrderBy(s => s.FechaSesion).ToList();
                if (sesionesGlobalesDocente.Count == 0) continue;

                var bloquesGlobales = ViaticoBloque.AgruparConsecutivos(sesionesGlobalesDocente);
                if (bloquesGlobales.Count == 0) continue;

                string cedula = sesionesGlobalesDocente.FirstOrDefault()?.CedulaTercero ?? string.Empty;
                _cacheTerceros.TryGetValue(cedula, out Tercero tercero);

                var nodoDocente = new TreeNode($"👤 {grupoDocente.Key}")
                {
                    ForeColor = Color.DarkGreen,
                    NodeFont = _boldFont
                };

                foreach (var bloque in bloquesGlobales.OrderBy(b => b.FechaMinima))
                {
                    var (origen, escala, destino, rutaFormateada, verificacion) = ConstruirRutaViaje(bloque, tercero?.Residencia);

                    var primeraSesion = bloque.Sesiones.FirstOrDefault();
                    var ultimaSesionBloque = bloque.Sesiones.OrderBy(s => s.FechaSesion).ThenBy(s => s.SedeCiudad).LastOrDefault();

                    var sesionAerea = bloque.Sesiones.FirstOrDefault(s => s.IsTerminalAereo == 1);
                    bool esAereo = sesionAerea != null;
                    string tipoTransporte = esAereo ? "Aéreo" : "Terrestre";

                    string territorial = primeraSesion?.SedeDepartamento;
                    decimal? transporteBase = ultimaSesionBloque?.TransporteSede;
                    double transporteValor = esAereo
                        ? (sesionAerea?.GastoTerminalAerea ?? 0)
                        : (transporteBase.HasValue ? (double)transporteBase.Value : 0);

                    // El código de trayecto se pre-genera (escritura en BD) aunque ya no se muestre
                    // en la fila de información del Nivel 2 (ver Fase 716).
                    if (!string.IsNullOrEmpty(verificacion) && _trayectoEncoderService != null)
                    {
                        _trayectoEncoderService.GetOrCreateCodigoTrayecto(territorial, verificacion, transporteBase);
                    }

                    string fechaIniStr = bloque.FechaMinima.HasValue ? bloque.FechaMinima.Value.ToString("dd/MM/yyyy") : "S/F";
                    DateTime? fFinEfectiva = bloque.FechaFinComision ?? bloque.FechaMaxima;
                    string fechaFinStr = fFinEfectiva.HasValue ? fFinEfectiva.Value.ToString("dd/MM/yyyy") : "S/F";

                    // Formato: 📅 FechaInicio - FechaFin | duración (d) | Aereo/Terrestre | Ruta | valor | DIA HORA | Pernocta
                    string textoBloque = ViaticoBloqueFormatter.FormatearNodoNivel2(
                        bloque, fechaIniStr, fechaFinStr, tipoTransporte, rutaFormateada, (decimal)transporteValor);

                    var nodoBloque = new TreeNode(textoBloque)
                    {
                        ForeColor = bloque.TienePernoctaUltimoDia ? Color.DarkOrange : Color.DarkBlue,
                        NodeFont = _boldFont,
                        Tag = bloque
                    };

                    foreach (var sesion in bloque.Sesiones.OrderBy(s => s.FechaSesion))
                    {
                        // Se envía 'true' para que se incluya la SedeCiudad en el detalle (nivel 3) de la vista de Viáticos.
                        nodoBloque.Nodes.Add(CrearNodoSesion(sesion, mostrarSede: true));
                    }

                    nodoDocente.Nodes.Add(nodoBloque);
                    totalViajes++;
                }

                treeVinculaciones.Nodes.Add(nodoDocente);
            }

            treeVinculaciones.EndUpdate();
            lblTotal.Text = $"{totalViajes} bloques de viáticos.";
            lblTotal.ForeColor = Color.Black;
            ExpandirHastaNivel(1);
            AgregarNodoViolacionesGlobales();
        }

        private void AgregarNodoViolacionesGlobales()
        {
            if (_violacionesGlobales == null || _violacionesGlobales.Count == 0) return;

            for (int i = 0; i < treeVinculaciones.Nodes.Count; i++)
            {
                if (treeVinculaciones.Nodes[i].Text.StartsWith("⚠️ Violaciones Generales de la Corrida"))
                    return;
            }

            TreeNode nodoGlobales = new TreeNode("⚠️ Violaciones Generales de la Corrida")
            {
                ForeColor = Color.DarkRed,
                NodeFont = _boldFont
            };

            foreach (var v in _violacionesGlobales)
            {
                string prefijo = v.EsDura ? "⛔ [HARD] " : "⚠️ [SOFT] ";
                var hijo = new TreeNode($"{prefijo}{v.Regla}: {v.Mensaje}")
                {
                    ForeColor = v.EsDura ? Color.Firebrick : Color.DarkOrange,
                    NodeFont = v.EsDura ? _boldFont : _regularFont,
                    Tag = v,
                    ToolTipText = !string.IsNullOrWhiteSpace(v.DetalleJson) ? $"{v.Mensaje}\n\nDetalle: {v.DetalleJson}" : v.Mensaje
                };
                nodoGlobales.Nodes.Add(hijo);
            }

            nodoGlobales.Expand();
            treeVinculaciones.Nodes.Insert(0, nodoGlobales);
        }

        private (string Origen, string Escala, string Destino, string RutaFormateada, string Verificacion) ConstruirRutaViaje(ViaticoBloque bloque, string residenciaDocente)
        {
            var sedesOrdenadas = bloque.Sesiones
                .Where(s => !string.IsNullOrWhiteSpace(s.SedeCiudad) || !string.IsNullOrWhiteSpace(s.Subsede))
                .OrderBy(s => s.FechaSesion)
                .ThenBy(s => string.IsNullOrWhiteSpace(s.SedeCiudad) ? s.Subsede : s.SedeCiudad)
                .Select(s => (string.IsNullOrWhiteSpace(s.SedeCiudad) ? s.Subsede : s.SedeCiudad).ToUpper())
                .ToList();

            var rutaCompleta = new List<string>();
            if (!string.IsNullOrWhiteSpace(residenciaDocente))
            {
                rutaCompleta.Add(residenciaDocente.ToUpper());
            }
            rutaCompleta.AddRange(sedesOrdenadas);

            var rutaLimpia = new List<string>();
            for (int i = 0; i < rutaCompleta.Count; i++)
            {
                if (i == 0 || rutaCompleta[i] != rutaCompleta[i - 1])
                    rutaLimpia.Add(rutaCompleta[i]);
            }

            string origen = !string.IsNullOrWhiteSpace(residenciaDocente) ? residenciaDocente.ToUpper() : null;
            string destino = rutaLimpia.LastOrDefault();

            string escala = null;
            if (rutaLimpia.Count > 2)
            {
                escala = string.Join("-", rutaLimpia.Skip(1).Take(rutaLimpia.Count - 2));
            }

            var segmentos = new List<string>();
            if (!string.IsNullOrWhiteSpace(residenciaDocente)) segmentos.Add(residenciaDocente.ToUpper());

            if (!string.IsNullOrWhiteSpace(escala))
            {
                segmentos.AddRange(escala.Split('-').Select(s => s.Trim().ToUpper()).Where(s => !string.IsNullOrWhiteSpace(s)));
            }

            if (!string.IsNullOrWhiteSpace(destino)) segmentos.Add(destino);
            if (!string.IsNullOrWhiteSpace(residenciaDocente) && !string.Equals(destino, residenciaDocente, StringComparison.OrdinalIgnoreCase))
                segmentos.Add(residenciaDocente.ToUpper());

            string verificacion = segmentos.Count >= 2 ? string.Join(" - ", segmentos) : null;
            string rutaFormateada = verificacion ?? "Ruta Inválida";

            return (origen, escala, destino, rutaFormateada, verificacion);
        }

        /// <summary>
        /// Crea un nodo de TreeView para representar una sesión académica (Nivel 3).
        /// </summary>
        /// <param name="sesion">Los datos de la sesión académica.</param>
        /// <param name="mostrarSede">Indica si se debe concatenar la SedeCiudad al final del texto del nodo.</param>
        private TreeNode CrearNodoSesion(SesionListItem sesion, bool mostrarSede = false)
        {
            string fecha = sesion.FechaSesion.HasValue ? sesion.FechaSesion.Value.ToString("dd/MM/yyyy") : "S/F";
            string diaSemana = (sesion.DiaSemana?.ToUpper() ?? "S/D").PadRight(9);
            bool esPlataforma = ModalidadTutoriaCatalog.EsPlataforma(sesion.Modalidad);

            // Construir el texto base de la sesión
            string textoSesion = $"   • {fecha} | {sesion.Horas}h | {diaSemana} | {sesion.HorarioSesion}";

            // Si se solicita, se anexa la sede extraída directamente de SedeCiudad (sin fallback) y en mayúsculas,
            // de forma análoga a como se muestra en la vista Por Fechas.
            if (mostrarSede)
            {
                string sede = sesion.SedeCiudad?.ToUpper() ?? "SIN SEDE";
                textoSesion += $" | {sede}";
            }

            string sufijoPlataforma = esPlataforma ? " | 🌐 Plataforma" : "";

            return new TreeNode($"{textoSesion}{sufijoPlataforma}")
            {
                Tag = sesion,
                ForeColor = esPlataforma ? Color.DarkBlue : SystemColors.ControlText,
                BackColor = esPlataforma ? Color.LightCyan : SystemColors.Window,
                NodeFont = _regularFont
            };
        }

        /// <summary>
        /// Genera el texto del tooltip describiendo con qué sesiones colisiona <paramref name="sesionA"/>.
        /// </summary>
        private string ObtenerDetalleCruceSesion(SesionListItem sesionA, List<SesionListItem> todasLasSesiones, EstadoConflicto estado)
        {
            var sb = new StringBuilder();
            if (estado == EstadoConflicto.ConflictoReal)
            {
                sb.AppendLine("⚠️ CONFLICTO REAL DE HORARIO");
                sb.AppendLine($"Las franjas horarias se solapan con otra sede:");
            }
            else
            {
                sb.AppendLine("⚡ REVISAR VIABILIDAD");
                sb.AppendLine($"Ocurre el mismo día en sedes distintas. Confirme viabilidad del traslado (Click derecho).");
            }

            bool huboCruce = false;
            foreach (var sesionB in todasLasSesiones)
            {
                if (sesionB == sesionA) continue;
                if (!sesionA.FechaSesion.HasValue || !sesionB.FechaSesion.HasValue) continue;
                if (string.Equals(sesionA.SedeCiudad ?? "", sesionB.SedeCiudad ?? "", StringComparison.OrdinalIgnoreCase)) continue;
                if (ModalidadTutoriaCatalog.EsPlataforma(sesionB.Modalidad)) continue;

                if (sesionA.FechaSesion.Value.Date == sesionB.FechaSesion.Value.Date)
                {
                    huboCruce = true;
                    sb.AppendLine();
                    sb.AppendLine($"📍 {sesionB.SedeCiudad?.ToUpper() ?? "SIN SEDE"}");
                    sb.AppendLine($"   {sesionB.Semestre} | {sesionB.AsignaturaNombre}");
                    sb.AppendLine($"   {sesionB.HorarioSesion}");
                }
            }

            return huboCruce ? sb.ToString() : string.Empty;
        }

        private void ExpandirHastaNivel(int nivelDeseado)
        {
            treeVinculaciones.BeginUpdate();
            treeVinculaciones.CollapseAll();
            foreach (TreeNode node in treeVinculaciones.Nodes)
            {
                ExpandirNodoRecursivo(node, nivelDeseado);
            }
            if (treeVinculaciones.Nodes.Count > 0)
                treeVinculaciones.Nodes[0].EnsureVisible();
            treeVinculaciones.EndUpdate();
        }

        private void ExpandirNodoRecursivo(TreeNode nodo, int nivelDeseado)
        {
            if (nodo.Level < nivelDeseado)
            {
                nodo.Expand();
                foreach (TreeNode hijo in nodo.Nodes)
                {
                    ExpandirNodoRecursivo(hijo, nivelDeseado);
                }
            }
        }

        private void btnPorAsignaturas_Click(object sender, EventArgs e)
        {
            _vistaActual = VistaVinculaciones.PorAsignaturas;
            ConstruirArbol();
        }

        private void btnPorSemestres_Click(object sender, EventArgs e)
        {
            _vistaActual = VistaVinculaciones.PorSemestres;
            ConstruirArbol();
        }

        private void btnPorFechas_Click(object sender, EventArgs e)
        {
            _vistaActual = VistaVinculaciones.PorFechas;
            ConstruirArbol();
        }

        private void btnViaticos_Click(object sender, EventArgs e)
        {
            _vistaActual = VistaVinculaciones.Viaticos;
            ConstruirArbol();
        }

        private void btnNivel1_Click(object sender, EventArgs e) => ExpandirHastaNivel(0);
        private void btnNivel2_Click(object sender, EventArgs e) => ExpandirHastaNivel(1);

        /// <summary>Expande todos los nodos con BeginUpdate/EndUpdate para evitar parpadeo (M2).</summary>
        private void btnNivel3_Click(object sender, EventArgs e)
        {
            treeVinculaciones.BeginUpdate();
            treeVinculaciones.ExpandAll();
            treeVinculaciones.EndUpdate();
        }

        private void btnVisualizar_Click(object sender, EventArgs e)
        {
            try
            {
                var (items, groups, cssRules) = ViaticoTimelineMapper.Mapear(_items, _cacheTerceros, incluirDatosBancarios: true);
                TimelineViewer.Mostrar(items, groups, cssRules);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error al abrir VisTimeline", ex);
                MessageBox.Show($"Error al generar la visualización:\n{ex.Message}", "VisTimeline", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (_vistaActual == VistaVinculaciones.PorFechas)
            {
                MessageBox.Show("No se puede exportar a Excel en la vista 'Por Fechas'. Seleccione 'Asignaturas', 'Semestres' o 'Viaticos'.", "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_vistaActual == VistaVinculaciones.Viaticos)
            {
                using (var formCdp = new FormCdpToPDF())
                {
                    var result = formCdp.ShowDialog(this);

                    if (result == DialogResult.OK)
                    {
                        var bloques = new List<ViaticoBloque>();
                        foreach (TreeNode nodoDocente in treeVinculaciones.Nodes)
                        {
                            foreach (TreeNode nodoBloque in nodoDocente.Nodes)
                            {
                                if (nodoBloque.Tag is ViaticoBloque b)
                                {
                                    bloques.Add(b);
                                }
                            }
                        }

                        if (bloques.Count == 0)
                        {
                            MessageBox.Show("No hay bloques de viáticos para exportar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        ExportarBloquesViaticos(bloques);
                    }
                    else if (result == DialogResult.Yes)
                    {
                        MessageBox.Show("Exportación masiva iniciada (Simulación).", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                return;
            }

            // --- Exportación Asignaturas/Semestres ---
            try
            {
                if (_vistaActual == VistaVinculaciones.PorSemestres)
                {
                    // Fase 635 & 637: Restaurar libro inicial o crear nuevo libro para evitar contaminar la plantilla FV-FO-033
                    bool restaurado = RestaurarFocoLibroInicial();
                    if (!restaurado)
                    {
                        dynamic activeWb = null;
                        try
                        {
                            activeWb = _excelApp.ActiveWorkbook;
                            string wbName = (string)activeWb?.Name;
                            // Si el libro activo es la plantilla presupuestal, crear un libro independiente para no contaminarla
                            if (wbName != null && wbName.StartsWith("FV-FO-033", StringComparison.OrdinalIgnoreCase))
                            {
                                dynamic nuevoLibro = _excelApp.Workbooks.Add();
                                if (nuevoLibro != null)
                                {
                                    try { Marshal.ReleaseComObject(nuevoLibro); } catch { }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[FormVerVinculaciones] Error verificando libro activo para Semestres: {ex.Message}");
                        }
                        finally
                        {
                            if (activeWb != null) { try { Marshal.ReleaseComObject(activeWb); } catch { } }
                        }
                    }

                    string periodoNombre = ObtenerNombrePeriodo();
                    var festivos = ObtenerRangosPorTipo("FESTIVO");
                    var presenciales = ObtenerRangosPorTipo("PRESENCIAL");

                    // Confirmado: SemestresExcelExporter existe en Academic\Infrastructure\Excel\SemestresExcelExporter.cs
                    var exporter = new SemestresExcelExporter(_excelApp);
                    exporter.Exportar(_items, periodoNombre, festivos, presenciales, _cacheTerceros);
                    MessageBox.Show("Exportación finalizada.", "Exportar Semestres", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AsignaturasExcelExporter exporter;
                    using (var progress = new ConversorPDF.XLL.UI.ProgressManager("Exportando Asignaturas a Excel... (max: 100%)", 100))
                    {
                        var repository = new GastosTerrestresRepository(_dbConnection);
                        var configRepo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.ConfigRepository(ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext());
                        var encoderService = new TrayectoEncoderService(repository, configRepo);
                        exporter = new AsignaturasExcelExporter(_excelApp, encoderService);
                        exporter.Exportar(_items, _cacheTerceros, progress);
                    }

                    var avisoCritico = exporter.UltimosAvisos.FirstOrDefault(a =>
                            a.IndexOf("continuación", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            a.IndexOf("saturación", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            a.IndexOf("Capacidad", StringComparison.OrdinalIgnoreCase) >= 0)
                        ?? exporter.UltimosAvisos.FirstOrDefault();

                    string detalleAvisos = avisoCritico != null ? $"\n\nPrimer aviso: {avisoCritico}" : string.Empty;
                    string mensajeFinal = $"Exportación finalizada. {exporter.UltimosBloquesReasignados} bloques reasignados / {exporter.UltimasFilasContinuacion} filas de continuación / {exporter.UltimosAvisos.Count} advertencias (ver log).{detalleAvisos}";
                    MessageBox.Show(mensajeFinal, "Exportar Asignaturas", MessageBoxButtons.OK, exporter.UltimosAvisos.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error al exportar a Excel", ex);
                MessageBox.Show($"Error al exportar:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Restaura el foco en Excel al libro de trabajo que estaba activo al abrir el formulario,
        /// liberando todas las referencias COM intermedias para evitar fugas de memoria o procesos huérfanos.
        /// </summary>
        /// <returns>True si el libro inicial fue localizado y activado; False en caso contrario.</returns>
        private bool RestaurarFocoLibroInicial()
        {
            if (string.IsNullOrWhiteSpace(_nombreLibroInicial))
            {
                return false;
            }

            dynamic workbooks = null;
            try
            {
                workbooks = _excelApp.Workbooks;
                if (workbooks == null) return false;

                int count = (int)workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = null;
                    try
                    {
                        wb = workbooks[i];
                        if (string.Equals((string)wb.Name, _nombreLibroInicial, StringComparison.OrdinalIgnoreCase))
                        {
                            wb.Activate();
                            return true;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (wb != null) { try { Marshal.ReleaseComObject(wb); } catch { } }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[FormVerVinculaciones] Error al restaurar foco a '{_nombreLibroInicial}': {ex.Message}");
            }
            finally
            {
                if (workbooks != null) { try { Marshal.ReleaseComObject(workbooks); } catch { } }
            }
            return false;
        }

        /// <summary>
        /// Exporta la lista proporcionada de bloques de viáticos a PDF utilizando la configuración actual del usuario.
        /// </summary>
        /// <param name="bloques">Lista de bloques de viáticos que se exportarán.</param>
        private void ExportarBloquesViaticos(List<ViaticoBloque> bloques)
        {
            var configRepo = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories.ConfigRepository(ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContext());
            var cfgResult = configRepo.GetConfig();
            if (cfgResult.IsFailure)
            {
                MessageBox.Show("Error al cargar configuración: " + cfgResult.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var viaticosRepo = new ConversorPDF.XLL.Core.Repositories.ViaticosRepository(new ConversorPDF.XLL.Core.Repositories.ExcelRepository());
            ConversorPDF.XLL.Core.Models.ViaticosConfig fullCfg = null;
            try
            {
                fullCfg = viaticosRepo.GetConfig(null, null, cfgResult.Value.Cdp);
            }
            catch (ConversorPDF.XLL.Core.Models.ViaticosFormatException ex)
            {
                MessageBox.Show(ex.Message, "Formato Incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inesperado al leer configuración: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (fullCfg == null)
            {
                MessageBox.Show("Error al cargar la configuración de viáticos desde la base de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var exporter = new ConversorPDF.XLL.Academic.Infrastructure.Excel.ViaticosIndividualExporter();

            try
            {
                Cursor = Cursors.WaitCursor;
                using (var progress = new ConversorPDF.XLL.UI.ProgressManager("Exportando viáticos a PDF...", bloques.Count))
                {
                    int procesados = exporter.Exportar(_excelApp, bloques, fullCfg, _cacheTerceros, new Action<int, int>((actual, total) =>
                    {
                        progress.ReportProgress(actual);
                    }));
                    MessageBox.Show($"Exportación finalizada.\n\nSe procesaron {procesados} viáticos.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    try
                    {
                        System.Diagnostics.Process.Start("explorer.exe", fullCfg.CarpetaDestino);
                    }
                    catch (Exception ex)
                    {
                        ConversorPDF.XLL.Core.Helpers.LoggerHelper.Warning("No se pudo abrir la carpeta destino: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Error("[ExportarBloquesViaticos] Error en exportador de viáticos", ex);
                MessageBox.Show("Ocurrió un error en la exportación: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private string ObtenerNombrePeriodo()
        {
            var primerItem = _items.FirstOrDefault();
            if (primerItem == null) return "";

            try
            {
                using (var cmd = _dbConnection.CreateCommand())
                {
                    cmd.CommandText = "SELECT periodo FROM periodos WHERE idperiodo = @id";
                    var param = cmd.CreateParameter();
                    param.ParameterName = "@id";
                    param.Value = primerItem.IdPeriodo;
                    cmd.Parameters.Add(param);

                    var result = cmd.ExecuteScalar();
                    return result != null && result != DBNull.Value ? result.ToString() : "";
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormVerVinculaciones] Error al obtener nombre del periodo", ex);
                return "";
            }
        }

        /// <summary>
        /// Obtiene de la base de datos una lista de rangos de fechas (Inicio, Fin) correspondientes a los festivos
        /// registrados en el periodo académico de las sesiones actuales, filtrados por un tipo de festivo específico.
        /// </summary>
        /// <param name="tipoFestivo">Tipo de fecha a consultar, por ejemplo 'FESTIVO' o 'PRESENCIAL'.</param>
        /// <returns>Lista de tuplas con las fechas de inicio y fin de cada rango encontrado.</returns>
        private List<(DateTime Inicio, DateTime Fin)> ObtenerRangosPorTipo(string tipoFestivo)
        {
            var rangos = new List<(DateTime Inicio, DateTime Fin)>();
            var primerItem = _items.FirstOrDefault();
            if (primerItem == null) return rangos;

            try
            {
                using (var cmd = _dbConnection.CreateCommand())
                {
                    cmd.CommandText = "SELECT fechafestivoinicio, fechafestivofinal FROM festivos WHERE idperiodo = @id AND tipofestivo = @tipo";

                    var pf = cmd.CreateParameter();
                    pf.ParameterName = "@id";
                    pf.Value = primerItem.IdPeriodo;
                    cmd.Parameters.Add(pf);

                    var pt = cmd.CreateParameter();
                    pt.ParameterName = "@tipo";
                    pt.Value = tipoFestivo;
                    cmd.Parameters.Add(pt);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var valIni = reader[0];
                            var valFin = reader[1];

                            if (valIni == DBNull.Value || valFin == DBNull.Value) continue;

                            if (DateTime.TryParse(valIni.ToString(), out DateTime ini) &&
                                DateTime.TryParse(valFin.ToString(), out DateTime fin))
                            {
                                rangos.Add((ini.Date, fin.Date));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[FormVerVinculaciones] Error al cargar rangos de tipo {tipoFestivo}", ex);
            }
            return rangos;
        }

        /// <summary>
        /// M3: Actualiza el estado visual de los botones de vista para indicar cuál está activo.
        /// Usa colores del sistema para compatibilidad con todos los temas de Windows.
        /// </summary>
        private void ActualizarEstadoBotones()
        {
            bool esPorAsignaturas = _vistaActual == VistaVinculaciones.PorAsignaturas;
            bool esPorSemestres = _vistaActual == VistaVinculaciones.PorSemestres;
            bool esPorFechas = _vistaActual == VistaVinculaciones.PorFechas;
            bool esViaticos = _vistaActual == VistaVinculaciones.Viaticos;

            btnPorAsignaturas.BackColor = esPorAsignaturas ? SystemColors.Highlight : SystemColors.Control;
            btnPorAsignaturas.ForeColor = esPorAsignaturas ? SystemColors.HighlightText : SystemColors.ControlText;
            btnPorAsignaturas.UseVisualStyleBackColor = false;

            btnPorSemestres.BackColor = esPorSemestres ? SystemColors.Highlight : SystemColors.Control;
            btnPorSemestres.ForeColor = esPorSemestres ? SystemColors.HighlightText : SystemColors.ControlText;
            btnPorSemestres.UseVisualStyleBackColor = false;

            btnPorFechas.BackColor = esPorFechas ? SystemColors.Highlight : SystemColors.Control;
            btnPorFechas.ForeColor = esPorFechas ? SystemColors.HighlightText : SystemColors.ControlText;
            btnPorFechas.UseVisualStyleBackColor = false;

            btnViaticos.BackColor = esViaticos ? SystemColors.Highlight : SystemColors.Control;
            btnViaticos.ForeColor = esViaticos ? SystemColors.HighlightText : SystemColors.ControlText;
            btnViaticos.UseVisualStyleBackColor = false;
        }

        private void btnCerrar_Click(object sender, EventArgs e) => Close();
    }
}
