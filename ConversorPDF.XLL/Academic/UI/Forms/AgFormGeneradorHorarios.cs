using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using CSharpFunctionalExtensions;
using Dapper;
using ExcelDna.Integration;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Formulario integral de optimización y programación automática de horarios mediante Algoritmo Genético.
    /// Incorpora Módulo S (Simulación y ejecución evolutiva) y Módulo R (Auditoría, revisión de staging y promoción).
    /// </summary>
    public partial class AgFormGeneradorHorarios : Form
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly IAgResolutorVentanaSemestre _resolutorVentana;
        private readonly IAgValidadorVentana _validadorVentana;
        private readonly IAgGeneradorHorarios _generadorHorarios;
        private readonly IAgScheduleRepository _scheduleRepository;
        private readonly IAgPlanAprobadoRepository _planAprobadoRepository;
        private readonly AgConfirmadorPlanAprobado _confirmadorPlanAprobado;
        private readonly HashSet<int> _aulasConfirmadasEnIntento = new HashSet<int>();
        private readonly IAgSalvedadRepository _salvedadRepository;
        private readonly IAgMaterializadorSesiones _materializador;
        private readonly IAgPromotorVinculaciones _promotor;
        private readonly AgSincronizadorCorrecciones _sincronizador;
        private readonly IAgArbolRepository _arbolRepository;
        private readonly IAgArbolService _arbolService;
        private readonly AgAulaDimensionesSync _aulaDimSync;
        private readonly IAgFiltroSeleccionRepository _filtroSeleccionRepo;
        private readonly IPeriodoCrudService _periodoCrudServiceOverride;
        private IPeriodoCrudService ObtenerPeriodoCrudService() =>
            _periodoCrudServiceOverride ?? new PeriodoCrudService(new AcademicRepository(_connectionFactory()));
        private bool _ventanaEnActualizacion;
        private HashSet<int> _tercerosSeleccionadosIds = new HashSet<int>();
        private readonly HashSet<int> _programasOcultos = new HashSet<int>();
        private bool _suprimirAutoFiltro;
        private bool _catalogosCargados;
        private bool _filtrosRestaurados;
        private bool _filtroArbolDesactualizado;
        private const string TituloGrpSeleccion = "1. Contexto Académico (Jerarquía 6 Niveles)";

        private System.Windows.Forms.Timer _timerSyncDimensiones;
        private (string AulaRef, string Jornada, int HorasDia, string DiasVentana, string Subunidad)? _dimensionesPendientes; // Fase 1731: aislar dimensión pendiente por subunidad activa
        private long _revisionDimensiones; // Fase 1728: contador de revisión para evitar que un await descarte una edición concurrente
        private readonly SemaphoreSlim _semaforoSyncDimensiones = new SemaphoreSlim(1, 1); // Fase 1716: exclusión mutua entre debounce de sincronización y escalado
        private bool _diagnosticoEnCurso; // Fase 1727: exclusión de reentrada concurrente en diagnóstico pre-vuelo

        private CancellationTokenSource _ctsEjecucion;
        private Task _tareaOptimizacion; // Fase 1900: Tarea de optimización en segundo plano para cierre seguro
        private bool _operacionEnCurso;  // Fase 1905: Ciclo de vida completo de la optimización y post-proceso
        private Guid? _intentoOptimizacion;
        private bool _aceptaProgresoOptimizacion;
        private enum EstadoVisualOptimizacion
        {
            Preparando, Validando, RequiereDecision, Optimizando, Materializando,
            Completado, Bloqueado, Cancelado, Error
        }
        private EstadoVisualOptimizacion _estadoVisualOptimizacion;
        private bool _terminalVisualOptimizacion;
        private string _faseFinOptimizacion;
        private string _motivoFinOptimizacion;
        private AgContexto _contextoActual;
        private AgResultadoCorrida _ultimoResultado;
        private long _idCorridaCargada;
        private bool _actualizandoComboCorridas;
        private bool _refrescandoComboCorridas;
        private string _claveFiltroEnVuelo; // Fase 1637: filtro en vuelo para no marcar pendiente si los filtros no cambiaron
        private bool _solicitudPendienteRefrescoCombo;
        private bool _cargaInicialPendiente; // Fase 1624: acumular autorización de carga inicial ante refrescos en curso
        private bool _cargandoCorridaDesdeCombo; // Fase 1616: exclusión mutua para evitar cargas concurrentes de corrida
        private bool _recargaCorridaPendiente; // Fase 1625: coalescencia de selección de corrida para evitar desalineación combo/grilla
        private bool _ultimaCargaDesdeComboOk; // Fase 1629: estado de última carga para inhibir cascada de modales
        private DateTime _ultimoRefrescoComboUtc = DateTime.MinValue; // Fase 1632: deduplicación temporal de refrescos
        private int _reabriendoPopupCombo; // Fase 1633, 1636: contador de reaperturas programáticas para no disparar refresco recursivo
        private bool _cerrandoPopupInterno; // Fase 1634: cierre programático para no disparar refresco en DropDownClosed
        private VentanaSemestre _ventanaResuelta;
        private bool _isUpdatingTreeChecks = false;
        private bool _restaurandoPerfil = false;
        private ImageList _imageListEstados;

        private List<Periodo> _periodos = new List<Periodo>();
        private List<Programa> _programas = new List<Programa>();
        private List<Sede> _sedes = new List<Sede>();
        private List<AgArbol> _perfilesArbol = new List<AgArbol>();
        private List<AgVacanteGridDto> _todasLasVacantes = new List<AgVacanteGridDto>();

        /// <summary>Compone servicios con la fábrica académica existente y un lector canónico opcional para pruebas.</summary>
        public AgFormGeneradorHorarios(Func<IDbConnection> connectionFactory, IAgPlanAprobadoRepository planAprobadoRepository = null,
            AcademicDbSession sesionRatificacion = null, EncryptedSQLiteContext vaultRatificacion = null, IPeriodoCrudService periodoCrudService = null)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _periodoCrudServiceOverride = periodoCrudService;

            // Inyección y cableado de servicios del AG
            _resolutorVentana = new AgResolutorVentanaSemestre(_connectionFactory);
            _validadorVentana = new AgValidadorVentana();
            var validadorReglas = new AgValidadorReglas();
            var evaluadorAptitud = new AgEvaluadorAptitud(validadorReglas);
            _generadorHorarios = new AgGeneradorHorarios(_validadorVentana, evaluadorAptitud);

            _scheduleRepository = new AgScheduleRepository(_connectionFactory);
            _planAprobadoRepository = planAprobadoRepository ?? new AgPlanAprobadoRepository(_connectionFactory);
            // Sin sesión y contexto de publicación explícitos, toda confirmación aprobada falla cerrada.
            _confirmadorPlanAprobado = sesionRatificacion == null || vaultRatificacion == null ? null
                : new AgConfirmadorPlanAprobado(sesionRatificacion, vaultRatificacion);
            _salvedadRepository = new AgSalvedadRepository(_connectionFactory);
            _arbolRepository = new AgArbolRepository(_connectionFactory);
            _arbolService = new AgArbolService(_arbolRepository);
            _aulaDimSync = new AgAulaDimensionesSync(_connectionFactory);
            _filtroSeleccionRepo = new AgFiltroSeleccionRepository(_connectionFactory);

            var expansor = new AgExpansorSesiones();
            _materializador = new AgMaterializadorSesiones(_scheduleRepository, expansor);

            var poolRepo = new TerceroAsignaturaRepository(_connectionFactory);
            var syncService = new TerceroAsignaturaSyncService(_connectionFactory, poolRepo);
            _promotor = new AgPromotorVinculaciones(_connectionFactory, _scheduleRepository, _salvedadRepository, syncService);
            _sincronizador = new AgSincronizadorCorrecciones(_scheduleRepository, syncService);

            InitializeComponent();
            ConfigurarEventos();

            // Fase 1193, 1215, 1219 & 1716: Temporizador con debounce (300 ms) para sincronización UI ↔ BD de dimensiones en ag_aula con semáforo
            _timerSyncDimensiones = components != null ? new System.Windows.Forms.Timer(components) : new System.Windows.Forms.Timer();
            _timerSyncDimensiones.Interval = 300;
            _timerSyncDimensiones.Tick += async (s, e) =>
            {
                _timerSyncDimensiones.Stop();
                if (_dimensionesPendientes == null || IsDisposed || Disposing)
                {
                    return;
                }

                await _semaforoSyncDimensiones.WaitAsync();
                try
                {
                    if (_dimensionesPendientes == null || IsDisposed || Disposing)
                    {
                        return;
                    }

                    long revision = _revisionDimensiones;
                    var pend = _dimensionesPendientes.Value;
                    // Fase 1214, Fase 1726 & Fase 1731: Sincronizar confirmadamente preservando la subunidad activa del pendiente
                    var resultado = await _aulaDimSync.SincronizarConfirmadoAsync(pend.AulaRef, pend.Jornada, pend.HorasDia, pend.DiasVentana, pend.Subunidad, CancellationToken.None);
                    if (IsDisposed || Disposing || !IsHandleCreated)
                    {
                        return;
                    }

                    if (resultado.IsFailure)
                    {
                        Log($"[DIM ERROR] Error al sincronizar {pend.AulaRef}: {resultado.Error}");
                        return;
                    }

                    // Fase 1728: Limpiar pendiente únicamente si no hubo una nueva asignación concurrente durante el await
                    if (resultado.IsSuccess && revision == _revisionDimensiones)
                    {
                        _dimensionesPendientes = null;
                    }

                    int n = resultado.Value;
                    Log($"[DIM] ag_aula {pend.AulaRef} → {pend.Jornada}/{pend.DiasVentana}/{pend.HorasDia}h (filas={n})");
                    if (n == 0)
                    {
                        Log("[DIM] aula_ref sin cambios pendientes en ag_aula");
                    }
                }
                finally
                {
                    _semaforoSyncDimensiones.Release();
                }
            };

            // Fase 299 & 305: Indicador visual canónico de Modo AG (Staging)
            if (lblAgModo != null)
            {
                lblAgModo.AutoSize = true;
                lblAgModo.ForeColor = Color.OrangeRed;
                lblAgModo.Text = $"Modo AG · staging: {AgModoDatos.Ag.TablaVinculaciones} / {AgModoDatos.Ag.TablaSesiones}";
            }

            // Fase 594/596: Detalle opt-in deshabilitado por defecto
            if (chkModoDebug != null)
            {
                chkModoDebug.Checked = false;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (splitContenedorSim == null) return;

            const int altoPanel2 = 239; // pnlBotonesSim 58 + txtLogAG 181
            int alto = splitContenedorSim.Height;
            int anchoSeparador = splitContenedorSim.SplitterWidth;

            splitContenedorSim.Panel1MinSize = 300; // 82 (config) + 218 mínimo útil del grupo
            splitContenedorSim.Panel2MinSize = 120;

            int destino = alto - altoPanel2 - anchoSeparador;
            int maximo = alto - splitContenedorSim.Panel2MinSize - anchoSeparador;
            if (destino >= splitContenedorSim.Panel1MinSize && destino <= maximo)
            {
                splitContenedorSim.SplitterDistance = destino;
            }
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            await CargarCatalogosAsync();
            await RefrescarComboArbolesAsync();

            // Fase 269 & Fase 1221: Restaurar automáticamente el último perfil guardado sin bloquear con modales
            if (_perfilesArbol != null && _perfilesArbol.Count > 0)
            {
                cmbArboles.SelectedIndex = 0;
                await CargarPerfilArbolSeleccionadoAsync(mostrarAvisos: false);
            }

            // Fase 1617: única invocación con autorización de carga inicial automática
            RefrescarComboCorridasSeguro(permitirCargaInicial: true);
        }

        /// <summary>
        /// Fuerza la persistencia inmediata al Vault cifrado (.db.enc) y el flush de OneDrive antes de cerrar el formulario (Fase 270).
        /// Emite traza diagnóstica estructurada P11 (Fase 279).
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Fase 1900 / 1905 / 1906: Si la optimización o el post-proceso de la corrida están en curso, cancelar y diferir el cierre
            if (_operacionEnCurso || (_tareaOptimizacion != null && !_tareaOptimizacion.IsCompleted))
            {
                _ctsEjecucion?.Cancel();
                e.Cancel = true;
                if (_tareaOptimizacion != null)
                {
                    _tareaOptimizacion.ContinueWith(_ =>
                    {
                        if (!IsDisposed && !Disposing && IsHandleCreated)
                        {
                            try { BeginInvoke(new Action(Close)); } catch (InvalidOperationException) { }
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
                return;
            }

            // Fase 1723: Aplazar cierre si existe una operación protegida (escalado/recarga/sincronización) en curso
            if (_semaforoSyncDimensiones.CurrentCount == 0)
            {
                e.Cancel = true;
                return;
            }

            var traza = CrearTraza("CierreFormulario");
            traza.Inicio();
            traza.Paso("P11_CierreFormulario", "Persistencia forzada de Vault y sincronización de estado");

            try
            {
                // Fase 1215, Fase 1218 & Fase 1730: detener debounce y volcar confirmadamente el último cambio de dimensiones antes de persistir el Vault
                _timerSyncDimensiones.Stop();
                if (_dimensionesPendientes != null)
                {
                    var pend = _dimensionesPendientes.Value;
                    try
                    {
                        var resultado = Task.Run(() => _aulaDimSync.SincronizarConfirmadoAsync(
                                pend.AulaRef, pend.Jornada, pend.HorasDia, pend.DiasVentana,
                                pend.Subunidad, CancellationToken.None))
                            .GetAwaiter().GetResult();

                        if (resultado.IsFailure)
                        {
                            e.Cancel = true;
                            traza.Fallo("P11_FlushDimensiones", new Exception(resultado.Error), AgCausaFallo.ExcepcionNoControlada);
                            MessageBox.Show(this, resultado.Error, "Cierre de formulario", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        _dimensionesPendientes = null;
                        int nCierre = resultado.Value;
                        traza.Dato("DimensionesAlCierre", $"{pend.AulaRef}={pend.Jornada}/{pend.DiasVentana}/{pend.HorasDia}h filas={nCierre}");
                    }
                    catch (Exception ex)
                    {
                        e.Cancel = true;
                        traza.Fallo("P11_FlushDimensiones", ex, AgCausaFallo.ExcepcionNoControlada);
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] Volcado de dimensiones al cierre falló: {ex.Message}");
                        MessageBox.Show(this, ex.Message, "Cierre de formulario", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                var scheduler = SyncContext.Scheduler;
                if (scheduler != null)
                {
                    var sharedCtx = SqliteHelper.GetContext();
                    if (sharedCtx != null)
                    {
                        try
                        {
                            scheduler.ForcePersistVaultAsync(sharedCtx).GetAwaiter().GetResult();
                            traza.Dato("VaultPersistido", true);
                        }
                        catch (Exception ex)
                        {
                            traza.Fallo("P11_ForcePersistVaultAsync", ex, AgCausaFallo.ExcepcionNoControlada);
                            LoggerHelper.Error("[AgFormGeneradorHorarios] Error al persistir el Vault en el cierre", ex);
                        }
                    }

                    // Forzar el PUSH a OneDrive de los cambios recientes (fire-and-forget)
                    scheduler.Flush();
                    traza.Dato("SchedulerFlushed", true);
                }
                traza.Fin(true, "Cierre seguro de formulario completado");
            }
            catch (Exception ex)
            {
                traza.Fallo("P11_OnFormClosing", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error en secuencia de cierre del formulario", ex);
            }

            try
            {
                GuardarFiltrosSincrono();
            }
            catch { }

            base.OnFormClosing(e);
        }

        private void ConfigurarEventos()
        {
            btnFiltrarArbol.Click += async (s, e) => await EjecutarFiltroArbolAsync();
            btnSeleccionarTerceros.Click += (s, e) => AbrirSeleccionTerceros();
            // Fase 1431, 1440, 1443 & 1449: Al cambiar tipo de programa, filtrar programas dinámicamente, marcar árbol desactualizado y actualizar resumen
            cmbTipoPrograma.SelectedIndexChanged += (s, e) =>
            {
                if (_suprimirAutoFiltro) return;
                string tipo = cmbTipoPrograma.SelectedItem?.ToString() ?? AgTipoProgramaCatalog.Todos;
                CargarProgramasSegunTipo(tipo);
                _filtroArbolDesactualizado = true;
                ActualizarResumenFiltros();
            };

            // Fase 1454: Resumen y banner en vivo al cambiar selecciones en los combos
            EventHandler onSeleccion = (s, e) =>
            {
                if (_suprimirAutoFiltro) return;
                _filtroArbolDesactualizado = true;
                ActualizarResumenFiltros();

                // Fase 1519: el selector de corridas depende del período/programa seleccionados
                RefrescarComboCorridasSeguro();
            };
            cmbPeriodos.SeleccionCambiada += onSeleccion;
            cmbSedes.SeleccionCambiada += onSeleccion;
            cmbProgramas.SeleccionCambiada += onSeleccion;

            cmbPeriodos.SelectedIndexChanged += OnVentanaChanged;
            chkManualVentana.CheckedChanged += (s, e) =>
            {
                dtpVentanaInicio.Enabled = chkManualVentana.Checked;
                dtpVentanaFin.Enabled = chkManualVentana.Checked;
                OnVentanaChanged(s, e);
            };
            // FASE 1863 & 1868: La extensión automática de período solo se autoriza ante cambios explícitos de fecha por el usuario
            dtpVentanaInicio.ValueChanged += DtpVentana_ValueChanged;
            dtpVentanaFin.ValueChanged += DtpVentana_ValueChanged;

            // Eventos del árbol jerárquico de 6 niveles (Fases 157, 158, 160)
            // Fase 542: radio en nivel 6 (y checks en 1-5)
            _imageListEstados = CrearImageListEstados();
            if (components != null)
            {
                components.Add(_imageListEstados);
            }
            treeContextoAcademico.CheckBoxes = false;
            treeContextoAcademico.ShowNodeToolTips = true;
            treeContextoAcademico.StateImageList = _imageListEstados;
            treeContextoAcademico.BeforeExpand += (s, e) => CargarHijosLazy(e.Node);
            treeContextoAcademico.AfterCheck += TreeContextoAcademico_AfterCheck;
            treeContextoAcademico.DrawNode += TreeContextoAcademico_DrawNode;
            treeContextoAcademico.NodeMouseClick += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return; // Fase 1025: clic derecho sólo abre el menú contextual sin mutar selección
                if (e.Node?.Tag is AgNodoArbolKey key)
                {
                    if (key.Nivel == 6)
                    {
                        int dy = e.Y - e.Node.Bounds.Y;
                        if (dy < 0 || dy > e.Node.Bounds.Height)
                        {
                            return;
                        }

                        int dx = e.X - e.Node.Bounds.X;
                        var c = CalcularColumnasNivel6();
                        if (dx >= c.DxManana)
                        {
                            ProcesarClickDimensionNivel6(e.Node, key, dx);
                            // Fase 663: Invalidar todas las modalidades hermanas para refrescar activación de radios
                            if (e.Node.Parent != null)
                            {
                                foreach (TreeNode hno in e.Node.Parent.Nodes)
                                {
                                    var rh = hno.Bounds;
                                    treeContextoAcademico.Invalidate(new Rectangle(0, rh.Y, treeContextoAcademico.ClientSize.Width, rh.Height));
                                }
                            }
                            else
                            {
                                var r = e.Node.Bounds;
                                treeContextoAcademico.Invalidate(new Rectangle(0, r.Y, treeContextoAcademico.ClientSize.Width, r.Height));
                            }
                            return;
                        }
                    }

                    var hit = treeContextoAcademico.HitTest(e.Location);
                    if (hit.Location == TreeViewHitTestLocations.StateImage || hit.Location == TreeViewHitTestLocations.Label)
                    {
                        e.Node.Checked = !e.Node.Checked;
                        TreeContextoAcademico_AfterCheck(this, new TreeViewEventArgs(e.Node));
                    }
                }
            };
            treeContextoAcademico.AfterExpand += (s, e) => ActualizarCabecerasNivel6();
            treeContextoAcademico.AfterCollapse += (s, e) => ActualizarCabecerasNivel6();
            treeContextoAcademico.ClientSizeChanged += (s, e) => ActualizarCabecerasNivel6();
            splitContenedorSim.SplitterMoved += (s, e) => ActualizarCabecerasNivel6();
            treeContextoAcademico.MouseMove += (s, e) =>
            {
                var node = treeContextoAcademico.GetNodeAt(e.Location);
                if (node?.Tag is AgNodoArbolKey key && key.Nivel == 6)
                {
                    int dx = e.X - node.Bounds.X;
                    var c = CalcularColumnasNivel6();
                    if ((dx >= c.DxManana && dx < c.DxSep1) || (dx >= c.DxLunJue && dx < c.DxSep2) || (dx >= c.DxHoras4 && dx < c.DxFin))
                    {
                        treeContextoAcademico.Cursor = Cursors.Hand;
                        return;
                    }
                }
                treeContextoAcademico.Cursor = Cursors.Default;
            };

            ActualizarCabecerasNivel6();
            treeContextoAcademico.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Space && treeContextoAcademico.SelectedNode?.Tag is AgNodoArbolKey)
                {
                    treeContextoAcademico.SelectedNode.Checked = !treeContextoAcademico.SelectedNode.Checked;
                    TreeContextoAcademico_AfterCheck(this, new TreeViewEventArgs(treeContextoAcademico.SelectedNode));
                    e.Handled = true;
                }
            };
            btnExpandirArbol.Click += (s, e) => ExpandirTodoElArbol();
            btnContraerArbol.Click += (s, e) => ContraerTodoElArbol();

            // Fase 1024: Menú contextual del árbol de selección académica (traslado de Marcar/Desmarcar/Invertir)
            var mnuArbol = new ContextMenuStrip(components ?? new System.ComponentModel.Container());
            mnuArbol.Items.Add("Expandir todo", null, (s, e) => ExpandirTodoElArbol());
            mnuArbol.Items.Add("Contraer todo", null, (s, e) => ContraerTodoElArbol());
            mnuArbol.Items.Add(new ToolStripSeparator());
            mnuArbol.Items.Add("Marcar todo", null, (s, e) => MarcarTodoElArbol(true));
            mnuArbol.Items.Add("Desmarcar todo", null, (s, e) => MarcarTodoElArbol(false));
            mnuArbol.Items.Add("Invertir selección", null, (s, e) => InvertirSeleccionArbol());
            treeContextoAcademico.ContextMenuStrip = mnuArbol;

            btnGenerarAulas.Click += async (s, e) => await GenerarEncabezadosAulasAsync();

            // Eventos de perfiles de árbol (Fase 240 & 241)
            btnNuevoArbol.Click += async (s, e) => await GuardarNuevoPerfilArbolAsync();
            btnSobrescribirArbol.Click += async (s, e) => await SobrescribirPerfilArbolAsync();
            btnCargarArbol.Click += async (s, e) => await CargarPerfilArbolSeleccionadoAsync();
            btnEliminarArbol.Click += async (s, e) => await EliminarPerfilArbolSeleccionadoAsync();

            btnValidarFactibilidad.Click += async (s, e) => await EjecutarDiagnosticoPreVueloAsync();
            btnEjecutarAG.Click += async (s, e) => await IniciarOptimizacionAsync();
            btnCancelar.Click += (s, e) => CancelarOptimizacion();
            btnPromoverProduccion.Click += async (s, e) => await PromoverAProduccionAsync();

            // Fase 1519: el botón "Cargar última optimización" se sustituye por un selector de corridas
            cmbCargarCorridas.SelectedIndexChanged += async (s, e) => await CargarCorridaDesdeComboAsync();
            // Fases 1622, 1623, 1628, 1631, 1632, 1633, 1634: refresco anticipado y deduplicado con inhibición de eventos programáticos
            cmbCargarCorridas.Enter += (s, e) => RefrescarComboCorridasSeguro(omitirSiReciente: true);
            cmbCargarCorridas.DropDown += (s, e) =>
            {
                if (_reabriendoPopupCombo > 0) return; // Fase 1633, 1636: reapertura programática, no refrescar
                RefrescarComboCorridasSeguro(omitirSiReciente: true);
            };
            cmbCargarCorridas.DropDownClosed += (s, e) =>
            {
                if (_cerrandoPopupInterno) return; // Fase 1634: cierre programático por mutación interna de Items
                if (_solicitudPendienteRefrescoCombo)
                {
                    _solicitudPendienteRefrescoCombo = false;
                    RefrescarComboCorridasSeguro();
                }
            };
            btnAprobarCorrida.Click += async (s, e) => await AprobarCorridaDesdeModuloRAsync();
            btnRecalcularViolaciones.Click += async (s, e) => await RecalcularViolacionesDesdeModuloRAsync();

            // Eventos Modo AG (Staging) desde Módulo R (Fases 297, 298, 299)
            btnAgVinculaciones.Click += async (s, e) => await AbrirVinculacionesModoAgAsync();
            btnAgVerVinculaciones.Click += async (s, e) => await AbrirDetalleSesionesModoAgAsync();
            btnAgSesiones.Click += async (s, e) => await AbrirSesionModoAgAsync();

            // Fases 322 y 331: Inicializar combo de motivos dinámicamente y eventos del Panel de Vacantes (Módulo R)
            cmbMotivoVacante.Items.Clear();
            cmbMotivoVacante.Items.Add("(Todos)");
            foreach (var cod in Enum.GetValues(typeof(AgMotivoVacante)).Cast<AgMotivoVacante>()
                    .Where(m => m != AgMotivoVacante.Ninguno).Select(m => m.Codigo()).Distinct())
            {
                cmbMotivoVacante.Items.Add(cod);
            }
            cmbMotivoVacante.SelectedIndex = 0;

            cmbMotivoVacante.SelectedIndexChanged += (s, e) => FiltrarYMostrarVacantes();
            dgvVacantes.SelectionChanged += (s, e) => ActualizarBotonContextualVacante();
            btnAccionVacante.Click += async (s, e) => await EjecutarAccionVacanteAsync();
        }

        /// <summary>
        /// Manejador centralizado y seguro con try/catch para eventos de cambio en la ventana semestral (Fase 254).
        /// Ignora invocaciones si se está restaurando programáticamente un perfil para evitar carreras (Fase 253).
        /// </summary>
        private async void OnVentanaChanged(object sender, EventArgs e)
        {
            await EjecutarVentanaChangedAsync(permitirExtension: false);
        }

        /// <summary>
        /// FASE 1868: Manejador tipado no anónimo con captura de excepciones para cambios de fecha manual de ventana,
        /// evitando async void sin protección en eventos de WinForms.
        /// </summary>
        private async void DtpVentana_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkManualVentana.Checked)
                {
                    await EjecutarVentanaChangedAsync(permitirExtension: true);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error en cambio de fecha de ventana", ex);
            }
        }

        /// <summary>
        /// FASE 1863: Desacopla la invocación de cambio de ventana distinguiendo si procede de edición de fechas
        /// (permitirExtension=true) o de cambio de período/check (permitirExtension=false) para evitar mutaciones espurias.
        /// </summary>
        private async Task EjecutarVentanaChangedAsync(bool permitirExtension)
        {
            // Fase 1439: Ignorar si se restaura perfil o si se están cargando/restaurando filtros de selección
            if (_restaurandoPerfil || _suprimirAutoFiltro) return;
            try
            {
                await ActualizarVentanaSemestralAsync(permitirExtension: permitirExtension);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al recalcular ventana semestral en evento de UI", ex);
            }
        }

        private async Task CargarCatalogosAsync()
        {
            var traza = CrearTraza("CargarCatalogos");
            traza.Inicio();
            traza.Paso("P01_CargarCatalogos", "Verificación de esquemas relacionales y carga de catálogos base");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    // Fase 278, 285 & 611: Verificación de esquema relacional consolidada
                    var chequeos = new (string Tabla, string[] Columnas)[]
                    {
                        ("periodos", new[] { "idperiodo", "periodo", "activo" }),
                        ("programas", new[] { "id", "nombreprograma", "activo" }),
                        ("sedes", new[] { "id", "ciudad", "subsede", "activo" }),
                        ("ag_aula", new[] { "id", "idsede", "idperiodo", "idprograma", "aula_ref", "estado" }),
                        ("ag_aula_asignatura", new[] { "id_aula", "id_asignatura", "horas_objetivo" }),
                        ("asignaturas", new[] { "id", "idprograma", "semestre", "nombre" }),
                        ("terceros_asignaturas", new[] { "id", "idtercero", "idprograma", "idasignatura", "activo" }),
                        ("terceros", new[] { "id", "nombre", "activo" }),
                        ("ag_corrida", new[] { "id", "idperiodo", "idprograma", "idaula", "estado" }),
                        ("ag_bloque_horario", new[] { "id", "dia_semana", "hora_inicio", "hora_fin" })
                    };
                    int esquemaOk = 0, esquemaFallo = 0;
                    foreach (var c in chequeos)
                    {
                        var faltantes = AgEsquemaDiagnostico.Verificar(db, traza, c.Tabla, c.Columnas);
                        if (faltantes.Count == 0) esquemaOk++; else esquemaFallo++;
                    }
                    LoggerHelper.Resumen("EsquemaBD", ("Tablas", chequeos.Length), ("OK", esquemaOk), ("ConFaltantes", esquemaFallo));

                    bool tieneVentana = false;
                    try
                    {
                        var cols = (await db.QueryAsync<string>("SELECT name FROM pragma_table_info('periodos');"))
                            .Select(c => c.ToLowerInvariant())
                            .ToList();
                        tieneVentana = cols.Contains("periodo_inicio") && cols.Contains("periodo_fin");
                    }
                    catch { }

                    string colInicio = tieneVentana ? "periodo_inicio" : "NULL";
                    string colFin = tieneVentana ? "periodo_fin" : "NULL";
                    string sqlPeriodos = $"SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, {colInicio} as PeriodoInicio, {colFin} as PeriodoFin, activo as Activo FROM periodos ORDER BY activo DESC, idperiodo DESC";
                    _periodos = (await db.QueryAsync<Periodo>(sqlPeriodos)).GroupBy(p => p.Id).Select(g => g.First())
                        .OrderByDescending(p => p.Activo).ThenByDescending(p => p.Id).ToList();

                    // Fase 216 & Fase 1466: deduplicar por Id. Catálogo completo (sin filtro de activo)
                    // para alinear cmbProgramas con FormVinculaciones (ProgramaLookupRepository.GetProgramasLookupAsync);
                    // el nivel académico se resuelve en memoria por modalidad (AgTipoProgramaCatalog).
                    _programas = (await db.QueryAsync<Programa>(AgProgramaCatalogoSql.SeleccionProgramas))
                        .GroupBy(p => p.Id)
                        .Select(g => g.First())
                        .ToList();
                    _sedes = (await db.QueryAsync<Sede>("SELECT id as Id, iddivipola as IdDivipola, departamento as Departamento, ciudad as Ciudad, subsede as Subsede, categoria as Categoria, transporte as Transporte, activo as Activo, subunidad as Subunidad FROM sedes WHERE activo = 1 ORDER BY ciudad, subsede")).ToList();

                    traza.Conteo("Periodos", _periodos.Count);
                    traza.Conteo("Programas", _programas.Count);
                    traza.Conteo("Sedes", _sedes.Count);

                    // Fase 1466: desglose por modalidad para diagnóstico (opt-in vía traza.Habilitado)
                    int nPregrado = _programas.Count(p => AgTipoProgramaCatalog.Coincide(p, AgTipoProgramaCatalog.Pregrado));
                    int nPosgrado = _programas.Count(p => AgTipoProgramaCatalog.Coincide(p, AgTipoProgramaCatalog.Posgrado));
                    traza.Dato("ProgramasPorModalidad", $"PREGRADO={nPregrado}; POSGRADO={nPosgrado}");
                    if (_programas.Count == 0)
                    {
                        LoggerHelper.Warning("[AgFormGeneradorHorarios] Catálogo de programas vacío: revise la sincronización académica antes de filtrar por tipo.");
                    }

                    cmbPeriodos.Cargar(_periodos, p => p.Id, p => p.Nombre);

                    // Fase 1447, 1448, 1457 & 1465: cmbProgramas con pseudo-ítem Todos, abreviaturas compactas y filtro inicial
                    cmbProgramas.IncluirItemTodos = true;
                    cmbProgramas.TextoItemTodos = "-- Todos los Programas --";
                    cmbProgramas.TextoVacio = "-- Todos los Programas --";
                    _programasOcultos.Clear(); // Fase 1465
                    CargarProgramasSegunTipo(AgTipoProgramaCatalog.Todos, null);

                    cmbSedes.Cargar(_sedes, s => s.Id, s => s.Nombre);

                    // Fase 1423: Suprimir autofiltro durante carga de catálogos para no sobreescribir la persistencia con listas vacías
                    _suprimirAutoFiltro = true;
                    bool filtrosOk;
                    try
                    {
                        cmbTipoPrograma.Items.Clear();
                        cmbTipoPrograma.Items.Add(AgTipoProgramaCatalog.Todos);
                        cmbTipoPrograma.Items.Add(AgTipoProgramaCatalog.Pregrado);
                        cmbTipoPrograma.Items.Add(AgTipoProgramaCatalog.Posgrado);
                        cmbTipoPrograma.SelectedIndex = 0;

                        filtrosOk = await RestaurarFiltrosAsync();
                    }
                    finally
                    {
                        _suprimirAutoFiltro = false;
                    }

                    // Fase 1438, 1442 & 1446: separar "combos poblados" de "lectura de filtros confiable"
                    _catalogosCargados = true;
                    _filtrosRestaurados = filtrosOk;
                    if (!filtrosOk)
                    {
                        LoggerHelper.Warning("[AgFormGeneradorHorarios] Persistencia de filtros inhibida hasta que el usuario aplique filtros manualmente.");
                    }
                }

                PoblarArbolContextoAcademico();
                await ActualizarVentanaSemestralAsync();
                traza.Fin(true, $"Catálogos cargados ({_periodos.Count} períodos, {_programas.Count} programas, {_sedes.Count} sedes)");
            }
            catch (Exception ex)
            {
                traza.Fallo("P01_CargarCatalogos", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error cargando catálogos", ex);
                MessageBox.Show($"Error al cargar catálogos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task ActualizarVentanaSemestralAsync(Periodo periodoObjetivo = null, bool permitirExtension = false)
        {
            if (_ventanaEnActualizacion) return;
            _ventanaEnActualizacion = true;
            try
            {
                var periodoSel = periodoObjetivo ?? cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo);
                if (periodoSel == null) return;

                var traza = CrearTraza("VentanaSemestral");
                traza.Inicio();
                traza.Paso("P02_ActualizarVentana", $"Periodo='{periodoSel.Nombre}', Manual={chkManualVentana.Checked}, PermitirExt={permitirExtension}");

                // Fase 164: Advertencia inmediata si el periodo seleccionado carece de fechas configuradas
                bool sinFechas = string.IsNullOrWhiteSpace(periodoSel.PeriodoInicio) || string.IsNullOrWhiteSpace(periodoSel.PeriodoFin);

                bool periodoExtendido = false;
                DateTime? fIniManual = chkManualVentana.Checked ? (DateTime?)dtpVentanaInicio.Value.Date : null;
                DateTime? fFinManual = chkManualVentana.Checked ? (DateTime?)dtpVentanaFin.Value.Date : null;

                // FASE 1855 & 1863: Solo extender período si fue autorizado explícitamente por edición de fecha manual (no al conmutar período)
                if (permitirExtension && chkManualVentana.Checked && fIniManual.HasValue && fFinManual.HasValue && periodoSel.TryObtenerVentana(out var perIni, out var perFin))
                {
                    var manualIni = fIniManual.Value.Date;
                    var manualFin = fFinManual.Value.Date;

                    if (manualFin > perFin || manualIni < perIni)
                    {
                        DateTime? nuevoIni = manualIni < perIni ? (DateTime?)manualIni : null;
                        DateTime nuevoFin = manualFin > perFin ? manualFin : perFin;

                        // FASE 1866 & 1869: Resolución lazy y ejecución fuera del hilo UI (STA) para evitar contención con CR-SQLite
                        var rExt = await Task.Run(() => ObtenerPeriodoCrudService().ExtenderVentanaFin(
                            periodoSel.IdPeriodo,
                            nuevoFin,
                            Environment.UserName,
                            "AG_VENTANA_MANUAL",
                            $"Extensión automática desde ventana manual {manualIni:yyyy-MM-dd}..{manualFin:yyyy-MM-dd}",
                            nuevoIni));

                        if (rExt.IsFailure)
                        {
                            lblOrigenVentana.Text = $"· No se pudo extender el período: {rExt.Error}";
                            lblOrigenVentana.ForeColor = Color.DarkRed;
                            lblOrigenVentana.BringToFront();
                            _ventanaResuelta = null;
                            _contextoActual = null;
                            MarcarErrorPreparacion("PERIODO_EXTENSION_" + rExt.Error);
                            traza.Fallo("P02_ExtenderPeriodo", new InvalidOperationException(rExt.Error), AgCausaFallo.SinVentana);
                            traza.Fin(false, rExt.Error);
                            return;
                        }

                        // Actualizar periodo en memoria y lista
                        periodoSel.PeriodoInicio = rExt.Value.PeriodoInicio;
                        periodoSel.PeriodoFin = rExt.Value.PeriodoFin;
                        var pLista = _periodos.FirstOrDefault(p => p.Id == periodoSel.Id);
                        if (pLista != null)
                        {
                            pLista.PeriodoInicio = rExt.Value.PeriodoInicio;
                            pLista.PeriodoFin = rExt.Value.PeriodoFin;
                        }

                        periodoExtendido = true;
                        _contextoActual = null;
                        traza.Dato("PeriodoExtendido", $"{perFin:yyyy-MM-dd} -> {periodoSel.PeriodoFin}");
                        LoggerHelper.Info($"[PERIODO] Ventana extendida idperiodo={periodoSel.IdPeriodo} {perFin:yyyy-MM-dd}->{periodoSel.PeriodoFin} actor={Environment.UserName}");

                        // Refrescar combo bajo _suprimirAutoFiltro para evitar re-entradas espurias (FASE 1860: Marcar + RefrescarTexto)
                        bool prevSuprimir = _suprimirAutoFiltro;
                        _suprimirAutoFiltro = true;
                        try
                        {
                            int selId = periodoSel.Id;
                            cmbPeriodos.Cargar(_periodos, p => p.Id, p => p.Nombre);
                            cmbPeriodos.Marcar(selId);
                            cmbPeriodos.RefrescarTexto();
                        }
                        finally
                        {
                            _suprimirAutoFiltro = prevSuprimir;
                        }
                    }
                }

                var res = await _resolutorVentana.ResolverVentanaAsync(periodoSel.Id, fIniManual, fFinManual, CancellationToken.None);
                if (res.IsSuccess)
                {
                    _ventanaResuelta = res.Value;
                    if (!chkManualVentana.Checked)
                    {
                        dtpVentanaInicio.Value = _ventanaResuelta.FechaInicio;
                        dtpVentanaFin.Value = _ventanaResuelta.FechaFin;
                    }

                    traza.Dato("Origen", _ventanaResuelta.Origen);
                    traza.Dato("Rango", $"{_ventanaResuelta.FechaInicio:yyyy-MM-dd} -> {_ventanaResuelta.FechaFin:yyyy-MM-dd}");
                    traza.Dato("DuracionDias", _ventanaResuelta.DuracionDias);

                    if (_ventanaResuelta.Origen == AgOrigenVentana.Periodos)
                    {
                        lblOrigenVentana.Text = $"· Origen: PERIODOS [{periodoSel.Nombre}] ({_ventanaResuelta.DuracionDias} días)";
                        lblOrigenVentana.ForeColor = Color.DarkGreen;
                    }
                    else if (_ventanaResuelta.Origen == AgOrigenVentana.Presencial)
                    {
                        lblOrigenVentana.Text = $"· Origen: FESTIVOS (PRESENCIAL) ({_ventanaResuelta.DuracionDias} días)";
                        lblOrigenVentana.ForeColor = Color.DarkBlue;
                    }
                    else if (_ventanaResuelta.Origen == AgOrigenVentana.Manual)
                    {
                        lblOrigenVentana.Text = periodoExtendido
                            ? $"· Origen: MANUAL (usuario) — período extendido a {periodoSel.PeriodoFin}"
                            : $"· Origen: MANUAL (usuario) ({_ventanaResuelta.DuracionDias} días)";
                        lblOrigenVentana.ForeColor = Color.DarkOrange;
                    }
                    else
                    {
                        lblOrigenVentana.Text = $"· Origen: {_ventanaResuelta.Origen} ({_ventanaResuelta.DuracionDias} días)";
                        lblOrigenVentana.ForeColor = Color.DarkGoldenrod;
                    }

                    lblOrigenVentana.BringToFront();
                    traza.Fin(true, $"Ventana resuelta vía {_ventanaResuelta.Origen} ({_ventanaResuelta.DuracionDias} días)");
                }
                else
                {
                    // FASE 1855: Supresión del fallback silencioso a la ventana vieja
                    _ventanaResuelta = null;
                    _contextoActual = null;
                    MarcarErrorPreparacion("SIN_VENTANA_" + res.Error);

                    traza.Fallo("P02_ActualizarVentana", new InvalidOperationException(res.Error), AgCausaFallo.SinVentana);
                    traza.Fin(false, res.Error);

                    if (sinFechas)
                    {
                        lblOrigenVentana.Text = $"· [AVISO] Período '{periodoSel.Nombre}' sin fechas en DB. Ingrese fechas manuales.";
                        lblOrigenVentana.ForeColor = Color.DarkOrange;
                    }
                    else
                    {
                        lblOrigenVentana.Text = $"· Sin ventana: {res.Error}";
                        lblOrigenVentana.ForeColor = Color.DarkRed;
                    }

                    lblOrigenVentana.BringToFront();
                }
            }
            finally
            {
                _ventanaEnActualizacion = false;
            }
        }

        /// <summary>
        /// Construye el contexto de ejecución del algoritmo genético cargando aulas, bloques, asignaturas
        /// y pool docente. Incorpora fallback defensivo (Fase 274) cuando la tabla asignaturas almacena el
        /// semestre con formato alterno (ej. números arábigos vs números romanos).
        /// Permite forzar período, programa y subunidad para recálculo de corridas ampliadas (Fase 1546).
        /// </summary>
        /// <param name="ct">Token de cancelación.</param>
        /// <param name="idPeriodoForzado">Período específico de la corrida a evaluar (opcional).</param>
        /// <param name="idProgramaForzado">Programa específico de la corrida a evaluar (opcional).</param>
        /// <param name="subunidadForzada">Subunidad territorial específica de la corrida (opcional).</param>
        /// <returns>Instancia de <see cref="AgContexto"/> con índices de ejecución construidos.</returns>
        private async Task<Result<AgContexto>> ConstruirContextoAsync(
            CancellationToken ct, int? idPeriodoForzado = null, int? idProgramaForzado = null, string subunidadForzada = null)
        {
            var periodoSel = cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo);
            var programaSel = cmbProgramas.Primario<Programa>() ?? (cmbProgramas.SelectedItem as Programa);
            var sedeSel = cmbSedes.Primario<Sede>() ?? (cmbSedes.SelectedItem as Sede);

            var combinaciones = ObtenerCombinacionesSeleccionadas();
            if (combinaciones.Any())
            {
                // Fase 545: Validar que no haya mezcla de múltiples períodos
                var periodosDistintos = combinaciones.Select(c => c.IdPeriodo).Where(p => p > 0).Distinct().ToList();
                if (periodosDistintos.Count > 1)
                {
                    return Result.Failure<AgContexto>($"Las combinaciones seleccionadas pertenecen a {periodosDistintos.Count} períodos académicos distintos (IDs: {string.Join(", ", periodosDistintos)}). No es posible generar u optimizar simultáneamente a través de múltiples períodos. Seleccione ramas correspondientes a un único período.");
                }

                var primera = combinaciones.First();
                var pTree = _periodos.FirstOrDefault(p => p.Id == primera.IdPeriodo);
                var progTree = _programas.FirstOrDefault(p => p.Id == primera.IdPrograma);
                var sTree = _sedes.FirstOrDefault(s => s.Id == primera.IdSede);
                if (pTree != null) periodoSel = pTree;
                if (progTree != null)
                {
                    programaSel = progTree;
                    var homonimos = _programas.Where(p => string.Equals(p.NombrePrograma?.Trim(), programaSel.NombrePrograma?.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                    if (homonimos.Count > 1)
                    {
                        LoggerHelper.Info($"[AgFormGeneradorHorarios] Programa homónimo detectado para '{programaSel.NombrePrograma}': IDs [{string.Join(", ", homonimos.Select(h => h.Id))}]");
                    }
                }
                if (sTree != null) sedeSel = sTree;
            }

            // Fase 1546: forzar período y programa de la corrida a recalcular cuando se indiquen
            if (idPeriodoForzado.HasValue && idPeriodoForzado.Value > 0)
            {
                periodoSel = _periodos.FirstOrDefault(p => p.Id == idPeriodoForzado.Value) ?? periodoSel;
            }
            if (idProgramaForzado.HasValue && idProgramaForzado.Value > 0)
            {
                programaSel = _programas.FirstOrDefault(p => p.Id == idProgramaForzado.Value) ?? programaSel;
            }

            if (periodoSel == null || programaSel == null || sedeSel == null)
            {
                return Result.Failure<AgContexto>("Debe seleccionar Periodo, Programa y Sede (en el árbol o en los selectores).");
            }

            // Fase 445: Asegurar sincronización de la ventana con el período objetivo seleccionado
            if (_ventanaResuelta == null || _ventanaResuelta.IdPeriodo != periodoSel.Id)
            {
                await ActualizarVentanaSemestralAsync(periodoSel);
            }

            // Fase 445: El activo no filtra; solo informa
            var activos = _periodos.Where(p => p.Activo).ToList();
            if (activos.Count > 1)
            {
                Log($"[AVISO] {activos.Count} períodos activos; el objetivo es el seleccionado: {periodoSel.Nombre ?? periodoSel.Id.ToString()}.");
            }

            var traza = CrearTraza("ConstruirContexto");
            traza.Inicio($"Periodo={periodoSel.Id}, Programa={programaSel.Id}, Sede={sedeSel.Id}");
            // Fase 453: Traza de período objetivo
            traza.Dato("PeriodoObjetivo", $"{periodoSel.Nombre ?? periodoSel.Id.ToString()} ({periodoSel.PeriodoInicio ?? ""}→{periodoSel.PeriodoFin ?? ""})");
            traza.Paso("P05_ConstruirContexto", "Carga relacional de aulas, bloques, asignaturas, pool docente y ocupaciones");

            using (var db = _connectionFactory())
            {
                if (db.State != ConnectionState.Open) db.Open();
                ct.ThrowIfCancellationRequested();

                // Fase 1546: respetar subunidad forzada de la corrida si se provee
                string sub = string.IsNullOrWhiteSpace(subunidadForzada)
                    ? (SyncContext.ActiveSubunidad ?? string.Empty)
                    : subunidadForzada;

                // Fases 454 & 477: Extraer semestres distintos de todas las combinaciones seleccionadas
                var semestresCombinaciones = combinaciones
                    .Select(c => AgFormatoAsignatura.ParsearSemestreGrupo(c.Semestre).Semestre)
                    .Where(s => !string.IsNullOrWhiteSpace(s) && s != "SIN ASIGNAR")
                    .Distinct().ToList();

                var semestresFiltro = semestresCombinaciones
                    .SelectMany(s => new[] { s, AgFormatoAsignatura.SemestreArabigo(s) })
                    .Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToArray();

                var (semestreRomano, grupoParsed) = AgFormatoAsignatura.ParsearSemestreGrupo(
                    combinaciones.Any() ? combinaciones.First().Semestre : string.Empty);
                string semestreRama = semestresCombinaciones.Count == 1 ? semestresCombinaciones[0] : (semestresCombinaciones.Count > 1 ? string.Join(",", semestresCombinaciones) : semestreRomano);
                string grupoRama = !string.IsNullOrWhiteSpace(grupoParsed)
                    ? grupoParsed
                    : (combinaciones.Any() ? combinaciones.First().Grupo : string.Empty);

                // Fase 1575: si el programa viene forzado (recálculo/reparación de una corrida), la corrida manda.
                // La resolución de homónimos (Fase 446/477) sustituía el programa solicitado por otro con el mismo
                // nombre (APT-53 id=1 vs APT-52 id=2), dejando contexto=3/1 frente a corrida=3/2.
                bool programaForzado = idProgramaForzado.HasValue && idProgramaForzado.Value > 0;
                // Fase 1575: la lista de homónimos se conserva a nivel de método porque también filtra aulas más abajo.
                var homonimosProg = new List<Programa>();
                if (!programaForzado)
                {
                    homonimosProg = _programas.Where(p => string.Equals(p.NombrePrograma?.Trim(), programaSel.NombrePrograma?.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                    if (homonimosProg.Count > 1)
                    {
                        var ids = homonimosProg.Select(h => h.Id).ToArray();
                        var semArrayHomonimo = semestresFiltro.Length == 0 ? new[] { "" } : semestresFiltro;
                        var conAsig = (await db.QueryAsync<int>(new CommandDefinition(
                            @"SELECT idprograma FROM asignaturas
                              WHERE idprograma IN @ids AND (@SinFiltro = 1 OR semestre IN @Semestres)
                              GROUP BY idprograma ORDER BY COUNT(*) DESC",
                            new { ids, SinFiltro = semestresFiltro.Length == 0 ? 1 : 0, Semestres = semArrayHomonimo }, cancellationToken: ct))).ToList();
                        if (conAsig.Count > 0 && conAsig[0] != programaSel.Id)
                        {
                            traza.Dato("ProgramaHomonimoResuelto", $"solicitado={programaSel.Id} → usado={conAsig[0]} ({semestreRama})");
                            LoggerHelper.Info($"[AgFormGeneradorHorarios] Programa homónimo resuelto: solicitado={programaSel.Id} → usado={conAsig[0]} ({semestreRama})");
                            programaSel = _programas.First(p => p.Id == conAsig[0]);
                        }
                    }
                }
                else
                {
                    traza.Dato("ProgramaHomonimoOmitido", $"forzado={programaSel.Id} ({programaSel.NombrePrograma})");
                    LoggerHelper.Info($"[AgFormGeneradorHorarios] Programa forzado #{programaSel.Id} '{programaSel.NombrePrograma}': se omite la resolución de homónimos (Fase 1575).");
                }

                // FASE 1855: Si no existe ventana resuelta, fallar cerrado y no construir contexto con ventana nula ni stale
                if (_ventanaResuelta == null)
                {
                    traza.Fallo("ConstruirContexto", new InvalidOperationException("No se ha resuelto una ventana semestral válida."), AgCausaFallo.SinVentana);
                    traza.Fin(false, "No se ha resuelto una ventana semestral válida.");
                    MarcarErrorPreparacion("SIN_VENTANA");
                    Log("[ERROR] No se ha resuelto una ventana semestral válida. Ingrese fechas válidas o configure el período.");
                    return Result.Failure<AgContexto>("No se ha resuelto una ventana semestral válida.");
                }

                var ctx = new AgContexto
                {
                    IdPeriodo = periodoSel.Id,
                    IdPrograma = programaSel.Id,
                    IdSede = sedeSel.Id,
                    SemestreActual = semestreRama,
                    GrupoActual = grupoRama,
                    Ventana = _ventanaResuelta,
                    Subunidad = sub,
                    Traza = traza
                };

                // Fase 525: alcance multi-sede / multi-programa derivado de las combinaciones marcadas
                var sedesSel = combinaciones.Select(c => c.IdSede).Where(i => i > 0).Distinct().ToList();
                var progsSel = combinaciones.Select(c => c.IdPrograma).Where(i => i > 0).Distinct().ToList();
                // Fase 1575: con programa forzado el alcance debe ser el de la corrida; de lo contrario la vista de
                // detalle staging filtra por el programa del árbol (homónimo) y devuelve 0 sesiones.
                if (programaForzado) progsSel = new List<int> { programaSel.Id };
                ctx.SedesAlcance = sedesSel;
                ctx.ProgramasAlcance = progsSel;

                // 1. Cargar Aulas (Fase 525: cargar aulas del período SIN filtro de sede ni programa)
                traza.Sql("ag_aula", "GetAulasPorPeriodoAsync", new { ctx.IdPeriodo, IdSede = (int?)null, IdPrograma = (int?)null, Semestre = (string)null, Grupo = (string)null, Sub = ctx.Subunidad });
                var aulasRes = await _scheduleRepository.GetAulasPorPeriodoAsync(
                    ctx.IdPeriodo,
                    null,
                    null,
                    null,
                    null,
                    ctx.Subunidad,
                    ct);

                ctx.Aulas = aulasRes.IsSuccess ? aulasRes.Value.ToList() : new List<AgAula>();
                traza.Conteo("Aulas(CargaGlobal)", ctx.Aulas?.Count ?? 0);

                int aulasCargadas = ctx.Aulas?.Count ?? 0;
                var idsCargadas = ctx.Aulas?.Select(a => a.Id).ToList() ?? new List<int>();

                // Fase 484/491/494, Fase 526, Fase 640 & Fase 645: Filtro en dos pasadas (coincidencia exacta AulaRef con idprograma, y tolerancia a homónimos SOLO para cohortes pendientes)
                var refsSel = new HashSet<string>(
                    combinaciones.Select(c => AgAulaFactory.ExtraerCohorte(c.AulaRef)).Where(r => r.Length > 0),
                    StringComparer.OrdinalIgnoreCase);

                if (refsSel.Count > 0)
                {
                    var refsExactas = new HashSet<string>(
                        combinaciones.Select(c => c.AulaRef).Where(r => !string.IsNullOrWhiteSpace(r)),
                        StringComparer.OrdinalIgnoreCase);

                    // Pasada 1: coincidencia exacta AulaRef (incluye idprograma canónico)
                    var filtradas = ctx.Aulas
                        .Where(a => refsExactas.Contains(a.AulaRef))
                        .ToList();

                    // Pasada 2: tolerancia a homónimos (Fase 491/640/645) SOLO para cohortes que no tuvieron match exacto en Pasada 1
                    var cohortesCubiertas = new HashSet<string>(
                        filtradas.Select(a => AgAulaFactory.ExtraerCohorte(a.AulaRef)), StringComparer.OrdinalIgnoreCase);
                    var pendientes = new HashSet<string>(
                        refsSel.Where(c => !cohortesCubiertas.Contains(c)), StringComparer.OrdinalIgnoreCase);
                    if (pendientes.Count > 0)
                    {
                        var idsHomonimos = new HashSet<int>(homonimosProg.Select(h => h.Id));
                        filtradas.AddRange(ctx.Aulas
                            .Where(a => !refsExactas.Contains(a.AulaRef))
                            .Where(a => pendientes.Contains(AgAulaFactory.ExtraerCohorte(a.AulaRef)))
                            .Where(a => progsSel.Contains(a.IdPrograma) || idsHomonimos.Contains(a.IdPrograma)));
                    }

                    traza.Conteo("Aulas(Seleccion)", filtradas.Count);
                    if (filtradas.Count > 0) ctx.Aulas = filtradas;
                    else traza.Advertencia("AulasSinMatchSeleccion",
                        $"0 aulas para {refsSel.Count} cohortes del perfil; se conserva alcance completo.");

                    // Fase 511/516/640: Logging fidedigno de combinaciones y cohortes
                    LoggerHelper.Info($"[Cohortes] solicitadas={combinaciones.Count} | refs=[{string.Join(" ; ", combinaciones.Select(c => c.AulaRef))}]");
                    LoggerHelper.Info($"[Cohortes] aulasCargadas={aulasCargadas} ids=[{string.Join(",", idsCargadas)}]");
                    LoggerHelper.Info($"[Cohortes] aulasFiltradas={filtradas.Count} refs=[{string.Join(" ; ", filtradas.Select(a => a.AulaRef))}]");
                }
                else
                {
                    traza.Advertencia("AulasSinCohortes", "Sin cohortes reconocibles: se conserva el alcance completo del período.");
                }

                // Fase 452: Validar aulas del período objetivo
                if (ctx.Aulas.Count == 0)
                {
                    traza.Fallo("P05_ValidarAulas", null, AgCausaFallo.SinAulas);
                    string sedesTxt = sedesSel.Count > 0 ? string.Join(",", sedesSel) : sedeSel.Id.ToString();
                    string progsTxt = progsSel.Count > 0 ? string.Join(",", progsSel) : programaSel.Id.ToString();
                    return Result.Failure<AgContexto>(
                        $"Período {periodoSel.Nombre}: no hay aulas (ag_aula) para sedes=[{sedesTxt}] programas=[{progsTxt}]. Ejecute 'Generar Aulas'.");
                }

                // Cargar todas las aulas del alcance (período y sede) para oferta global pre-vuelo (Fase 379 / Fase 440)
                ctx.AulasAlcance = ctx.Aulas != null ? ctx.Aulas.ToList() : new List<AgAula>();
                traza.Conteo("Aulas(Alcance)", ctx.AulasAlcance.Count);

                // Fases 231, 235 & 454: IdAulaCorrida anclada solo si hay exactamente UNA combinación y un aula coincidente
                ctx.IdAulaCorrida = 0;
                if (combinaciones.Count == 1)
                {
                    var comb = combinaciones.First();
                    string aulaRefEsperada = comb.AulaRef;

                    var coincidencias = ctx.Aulas.Where(a => string.Equals(a.AulaRef, aulaRefEsperada, StringComparison.OrdinalIgnoreCase)).ToList();

                    // Fase 647: fallback por cohorte cuando no hay coincidencia exacta (programa homónimo)
                    if (coincidencias.Count == 0)
                    {
                        string cohorteEsperada = AgAulaFactory.ExtraerCohorte(aulaRefEsperada);
                        if (cohorteEsperada.Length > 0)
                        {
                            coincidencias = ctx.Aulas
                                .Where(a => string.Equals(AgAulaFactory.ExtraerCohorte(a.AulaRef), cohorteEsperada, StringComparison.OrdinalIgnoreCase))
                                .ToList();
                            if (coincidencias.Count == 1)
                                LoggerHelper.Info($"[AgFormGeneradorHorarios] Anclaje por cohorte homónima: {aulaRefEsperada} → aula {coincidencias[0].Id} ({coincidencias[0].AulaRef})");
                        }
                    }

                    if (coincidencias.Count == 1)
                    {
                        ctx.IdAulaCorrida = coincidencias[0].Id;
                        traza.Dato("IdAulaCorrida(Anclada)", ctx.IdAulaCorrida);
                    }
                    else if (coincidencias.Count > 1)
                    {
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] Combinación ambigua {aulaRefEsperada}: {coincidencias.Count} aulas encontradas.");
                        traza.Dato("IdAulaCorrida(Ambigua)", $"{coincidencias.Count} coincidencias");
                        ctx.IdAulaCorrida = 0;
                    }
                    else
                    {
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] No existe Aula (ag_aula) para la combinación exacta {aulaRefEsperada}.");
                        traza.Dato("IdAulaCorrida(NoExiste)", aulaRefEsperada);
                        ctx.IdAulaCorrida = 0;
                    }
                }

                // 2. Cargar Bloques
                traza.Sql("ag_bloque_horario", "GetBloquesHorariosAsync", new { Sub = ctx.Subunidad });
                var bloquesRes = await _scheduleRepository.GetBloquesHorariosAsync(ctx.Subunidad, ct);
                ctx.Bloques = bloquesRes.IsSuccess ? bloquesRes.Value.ToList() : new List<AgBloqueHorario>();
                ctx.InvalidarIndiceBloques(); // Fase 1128: invalidar índice de bloques ante recarga
                traza.Conteo("Bloques", ctx.Bloques?.Count ?? 0);

                // 3. Cargar Asignaturas del Programa con Demanda Acotada Multi-Semestre (Fase 379 / Fase 400 / Fase 477 / Fase 527)
                const string sqlAsigMulti = @"
SELECT id AS Id, idprograma AS IdPrograma, semestre AS Semestre, num_nivel AS NumNivel,
       cod_unidad AS CodUnidad, cod_materia AS CodMateria, uni_teorica AS UniTeorica,
       metodologia AS Metodologia, abreviatura AS Abreviatura, nucleo AS Nucleo,
       nombre AS Nombre, creditos AS Creditos, int_horaria AS IntHoraria,
       tipoasignatura AS TipoAsignatura, tipocredito AS TipoCredito, prerequisito AS Prerequisito
FROM asignaturas
WHERE idprograma IN @IdProgramas
  AND (@SinFiltro = 1 OR semestre IN @Semestres)
  AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)
ORDER BY num_nivel, nombre";

                var progsParam = progsSel.Count > 0 ? progsSel : new List<int> { programaSel.Id };
                var semArray = semestresFiltro.Length == 0 ? new[] { "" } : semestresFiltro;
                int sinFiltroVal = semestresFiltro.Length == 0 ? 1 : 0;
                traza.Sql("asignaturas", sqlAsigMulti, new { IdProgramas = progsParam, SinFiltro = sinFiltroVal, Semestres = semArray, Sub = sub });
                ctx.Asignaturas = (await db.QueryAsync<Asignatura>(new CommandDefinition(
                    sqlAsigMulti, new { IdProgramas = progsParam, SinFiltro = sinFiltroVal, Semestres = semArray, Sub = sub }, cancellationToken: ct))).ToList();
                traza.Conteo("Asignaturas(MultiSemestre)", ctx.Asignaturas?.Count ?? 0);

                // Fase 448 & 449: Validar asignaturas con diagnóstico completo y propagación de fallo
                if (ctx.Asignaturas.Count == 0)
                {
                    traza.Fallo("P05_DemandaAsignaturas",
                        new Exception($"Sin asignaturas: idprogramas=[{string.Join(",", progsParam)}], semestres='{semestreRama}', subunidad='{sub}'"),
                        AgCausaFallo.SinAsignaturas);

                    string progsDetalle = string.Join(", ", progsParam);
                    string msgError = $"El programa/período seleccionado no tiene asignaturas registradas.\n" +
                                     $"Período: {periodoSel.Nombre} ({periodoSel.PeriodoInicio}→{periodoSel.PeriodoFin})\n" +
                                     $"Programas: [{progsDetalle}] · Semestres: {semestreRama} · Subunidad: {sub}\n" +
                                     "Acción: cargue el pensum del programa o seleccione el programa con catálogo.";
                    return Result.Failure<AgContexto>(msgError);
                }

                // 4. Cargar Pool Docente por Unión Multi-Semestre (Fase 384 / Fase 407 / Fase 477 / Fase 527)
                const string sqlPoolMulti = @"
SELECT id as Id, idtercero as IdTercero, idprograma as IdPrograma, idasignatura as IdAsignatura,
       semestre as Semestre, activo as Activo, prioridad as Prioridad, veces_asignada as VecesAsignada,
       fecha_ultima_asignacion as FechaUltimaAsignacion, origen as Origen, observaciones as Observaciones, subunidad as Subunidad
FROM terceros_asignaturas
WHERE activo = 1 AND idprograma IN @IdProgramas
  AND (@SinFiltro = 1 OR semestre IN @Semestres OR semestre = '' OR semestre IS NULL)
  AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)";
                traza.Sql("terceros_asignaturas", sqlPoolMulti, new { IdProgramas = progsParam, SinFiltro = sinFiltroVal, Semestres = semArray, Sub = sub });
                ctx.PoolDocentes = (await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(
                    sqlPoolMulti, new { IdProgramas = progsParam, SinFiltro = sinFiltroVal, Semestres = semArray, Sub = sub }, cancellationToken: ct))).ToList();
                traza.Conteo("PoolDocentes(Union)", ctx.PoolDocentes?.Count ?? 0);
                traza.Dato("PoolDocentes(Desglose)",
                    $"filtrado={ctx.PoolDocentes.Count(p => semestresFiltro.Contains(p.Semestre))} general={ctx.PoolDocentes.Count(p => string.IsNullOrWhiteSpace(p.Semestre))}");

                // 5. Cargar Docentes (Fase 135: solo activos, correo as Correo, subunidad)
                string sqlDocentes = @"
SELECT id AS Id, nombre AS Nombre, cedula AS Cedula, residencia AS Residencia,
       correo AS Correo, telefono AS Telefono, activo AS Activo, subunidad AS Subunidad
FROM terceros
WHERE activo = 1
  AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)
ORDER BY nombre";
                traza.Sql("terceros", sqlDocentes, new { Sub = sub });
                ctx.Docentes = (await db.QueryAsync<Tercero>(new CommandDefinition(sqlDocentes, new { Sub = sub }, cancellationToken: ct))).ToList();
                traza.Conteo("Docentes", ctx.Docentes?.Count ?? 0);

                // Fase 462: Resolver tercero centinela VACANTE (cédula 1111111111)
                ctx.IdTerceroVacante = (await db.QueryFirstOrDefaultAsync<int?>(new CommandDefinition(
                    @"SELECT id FROM terceros
                      WHERE (cedula = '1111111111' OR UPPER(TRIM(nombre)) = 'VACANTE')
                        AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)
                      ORDER BY CASE WHEN subunidad = @Sub THEN 0 ELSE 1 END, id LIMIT 1",
                    new { Sub = sub }, cancellationToken: ct))) ?? 0;
                traza.Dato("IdTerceroVacante", ctx.IdTerceroVacante);

                ctx.Programas = _programas;
                ctx.Sedes = _sedes;

                // 6. Cargar Festivos y Presenciales
                string sqlFest = "SELECT fechafestivoinicio, fechafestivofinal, tipofestivo FROM festivos WHERE idperiodo = @idPeriodo";
                var filasFest = (await db.QueryAsync(new CommandDefinition(sqlFest, new { idPeriodo = periodoSel.Id }, cancellationToken: ct))).ToList();

                foreach (var f in filasFest)
                {
                    DateTime dIni = DateTime.MinValue;
                    DateTime dFin = DateTime.MinValue;
                    if (DateTime.TryParse(f.fechafestivoinicio?.ToString(), out dIni) &&
                        DateTime.TryParse(f.fechafestivofinal?.ToString(), out dFin))
                    {
                        string tipo = (f.tipofestivo?.ToString() ?? "").Trim().ToUpperInvariant();
                        if (tipo == "PRESENCIAL") ctx.Presenciales.Add((dIni.Date, dFin.Date));
                        else ctx.Festivos.Add((dIni.Date, dFin.Date));
                    }
                }
                traza.Conteo("Festivos", ctx.Festivos.Count, 0);
                traza.Conteo("Presenciales", ctx.Presenciales.Count, 0);

                // Fase 450: Aviso si no hay festivos ni presenciales
                if (ctx.Festivos.Count == 0 && ctx.Presenciales.Count == 0)
                {
                    traza.Advertencia("SinFestivos", $"Período {periodoSel.Nombre} sin festivos/presenciales: R5 y viáticos inactivos.");
                    Log($"[AVISO] Período {periodoSel.Nombre} sin festivos/presenciales registrados: R5 y viáticos inactivos.");
                }

                // 7. Cargar Ocupación Externa (Sesiones de Producción y Staging Aprobado) - Fase 126
                DateTime fi = _ventanaResuelta != null ? _ventanaResuelta.FechaInicio.Date : DateTime.Today;
                DateTime ff = _ventanaResuelta != null ? _ventanaResuelta.FechaFin.Date : DateTime.Today.AddMonths(5);
                string fiStr = fi.ToString("yyyy-MM-dd");
                string ffStr = ff.ToString("yyyy-MM-dd");

                var ocupacionesProduccion = new List<(long IdSesion, int IdTercero, DateTime Fecha, FranjaHoraria Franja)>();
                var ocupacionesStaging = new List<(int IdTercero, DateTime Fecha, FranjaHoraria Franja)>();

                // Sesiones de Producción en el rango
                string sqlProdSes = @"
SELECT id AS IdSesion, idtercero AS IdTercero, fechasesion AS FechaSesion, horariosesion AS HorarioSesion
FROM sesiones
WHERE fechasesion BETWEEN @fi AND @ff
  AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL);";

                var prodSes = (await db.QueryAsync(new CommandDefinition(sqlProdSes, new { fi = fiStr, ff = ffStr, Sub = ctx.Subunidad }, cancellationToken: ct))).ToList();
                foreach (var s in prodSes)
                {
                    int idTercero = s.IdTercero != null ? Convert.ToInt32(s.IdTercero) : 0;
                    if (idTercero <= 0) continue;

                    if (DateTime.TryParse(s.FechaSesion?.ToString(), out DateTime fSes))
                    {
                        var frRes = HorarioSesionParser.TryParse(s.HorarioSesion?.ToString());
                        if (frRes.IsSuccess)
                        {
                            ocupacionesProduccion.Add((Convert.ToInt64(s.IdSesion), idTercero, fSes.Date, frRes.Value));
                        }
                    }
                }

                // Sesiones de Staging de Aulas/Corridas ya aprobadas
                string sqlAgSes = @"
SELECT s.id_aula AS IdAula, s.idtercero AS IdTercero, s.fechasesion AS FechaSesion, s.horariosesion AS HorarioSesion
FROM ag_sesiones s
INNER JOIN ag_aula a ON s.id_aula = a.id
WHERE a.estado = 4 AND s.fechasesion BETWEEN @fi AND @ff;";

                try
                {
                    var agSes = (await db.QueryAsync(new CommandDefinition(sqlAgSes, new { fi = fiStr, ff = ffStr }, cancellationToken: ct))).ToList();
                    foreach (var s in agSes)
                    {
                        // Staging previo del aula aprobada seleccionada no es producción ni ocupación externa.
                        if (ctx.Aulas.Any(a => a.Id == Convert.ToInt32(s.IdAula) && a.Estado == AgEstadoAula.Aprobado)) continue;
                        int idTercero = s.IdTercero != null ? Convert.ToInt32(s.IdTercero) : 0;
                        if (idTercero <= 0) continue;

                        if (DateTime.TryParse(s.FechaSesion?.ToString(), out DateTime fSes))
                        {
                            var frRes = HorarioSesionParser.TryParse(s.HorarioSesion?.ToString());
                            if (frRes.IsSuccess)
                            {
                                ocupacionesStaging.Add((idTercero, fSes.Date, frRes.Value));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[AgFormGeneradorHorarios] Ocupaciones de staging omitidas: {ex.Message}");
                }

                ctx.CargarOcupacionesIdentificadas(ocupacionesProduccion, ocupacionesStaging);

                // 8. Cargar AulaAsignaturas para todo el alcance de aulas (Fases 293 & 473)
                var idsAulas = ctx.Aulas != null ? ctx.Aulas.Select(a => a.Id).ToList() : new List<int>();
                if (idsAulas.Count > 0)
                {
                    const string sqlAulaAsig = @"
SELECT id_aula AS IdAula, id_asignatura AS IdAsignatura, exclusivo AS Exclusivo, prioridad AS Prioridad,
       horas_objetivo AS HorasObjetivo, orden_curricular AS OrdenCurricular, requerida AS Requerida,
       dias_patron AS DiasPatron, subunidad AS Subunidad, origen_patron AS OrigenPatron
FROM ag_aula_asignatura
WHERE id_aula IN @ids AND id_asignatura > 0;";
                    ctx.AulaAsignaturas = (await db.QueryAsync<AgAulaAsignatura>(
                        new CommandDefinition(sqlAulaAsig, new { ids = idsAulas }, cancellationToken: ct))).ToList();
                    traza.Sql("ag_aula_asignatura", sqlAulaAsig, new { ids = idsAulas });
                    traza.Conteo("AulaAsignaturas(Alcance)", ctx.AulaAsignaturas.Count);

                    // Fase 516: Log oportuno de asignaturas por aula tras su carga efectiva
                    LoggerHelper.Info($"[Cohortes] asigsPorAula=[{string.Join(" ; ", ctx.Aulas.Select(a => $"{a.Id}:{a.AulaRef}:{ctx.AulaAsignaturas.Count(x => x.IdAula == a.Id)}"))}] totalAulaAsig={ctx.AulaAsignaturas.Count}");
                }
                else
                {
                    ctx.AulaAsignaturas = new List<AgAulaAsignatura>();
                    traza.Conteo("AulaAsignaturas", 0, minimo: 0);
                }

                if (ctx.AulaAsignaturas.Count == 0)
                {
                    traza.Advertencia("SinAulaAsignaturas", "Sin ag_aula_asignatura para el alcance: se usará patrón por defecto (2 días). Ejecute 'Generar Aulas'.");
                }

                // P06 & Fase 1430: Filtrar pool docente por terceros seleccionados y diagnóstico de pool vacío
                if (_tercerosSeleccionadosIds != null && _tercerosSeleccionadosIds.Count > 0)
                {
                    int poolAntes = ctx.PoolDocentes?.Count ?? 0;
                    ctx.FiltrarPoolDocentesPorTerceros(_tercerosSeleccionadosIds);

                    if (poolAntes > 0 && (ctx.PoolDocentes == null || ctx.PoolDocentes.Count == 0))
                    {
                        traza.Advertencia("P06_FiltroTercerosVacio", "La selección de terceros no intersecta el pool docente del alcance.");
                        MarcarErrorPreparacion("P06_FiltroTercerosVacio");
                        MessageBox.Show(this,
                            "La selección de terceros no intersecta el pool docente disponible para el alcance actual.\nAmplíe la selección o revise asignaturas habilitadas.",
                            "Sin docentes disponibles", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return Result.Failure<AgContexto>("La selección de terceros no intersecta el pool docente disponible.");
                    }
                }

                // P06: Construir índices en memoria
                traza.Paso("P06_ConstruirIndices", "Construcción de índices relacionales y tablas hash en memoria");
                ctx.ConstruirIndices();
                traza.Contexto(ctx);

                // Fase 296 & Fase 642: Traza de cobertura docente y detección real de aulas huérfanas sin asignaturas
                int totalAsignaturas = ctx.Asignaturas?.Count ?? 0;
                int asignaturasConPool = ctx.Asignaturas?
                    .Count(a => ctx.PoolPorAsignatura != null &&
                                ctx.PoolPorAsignatura.TryGetValue(a.Id, out var pList) &&
                                pList.Any(p => p.Activo == 1)) ?? 0;
                var aulasHuerfanas = ctx.Aulas != null
                    ? ctx.Aulas
                        .Where(a => ctx.AulaAsignaturas == null || !ctx.AulaAsignaturas.Any(x => x.IdAula == a.Id))
                        .Select(a => a.AulaRef).ToList()
                    : new List<string>();
                traza.Dato("CoberturaDocente", $"{asignaturasConPool}/{totalAsignaturas}");
                traza.Dato("AulasSinAsignaturas", aulasHuerfanas.Count);
                if (aulasHuerfanas.Count > 0)
                    traza.Advertencia("AulasHuerfanasAlcance",
                        $"{aulasHuerfanas.Count} aula(s) sin ag_aula_asignatura: {string.Join(" ; ", aulasHuerfanas)}");

                // Fase 1194: Telemetría diagnóstica de aulas sin dimensiones (jornada/días/horas) persistidas
                var aulasSinDim = ctx.Aulas != null
                    ? ctx.Aulas.Where(a => string.IsNullOrWhiteSpace(a.DiasVentana) || string.IsNullOrWhiteSpace(a.Jornada) || a.HorasDia <= 0).Select(a => a.AulaRef).ToList()
                    : new List<string>();
                traza.Dato("AulasSinDimensiones", aulasSinDim.Count);
                if (aulasSinDim.Count > 0)
                    traza.Advertencia("AulasSinDimensionesAlcance",
                        $"{aulasSinDim.Count} aula(s) sin dimensiones en ag_aula: {string.Join(" ; ", aulasSinDim)}");

                traza.Fin(true, $"Contexto preparado exitosamente (IdAulaCorrida={ctx.IdAulaCorrida})");

                return Result.Success(ctx);
            }
        }

        /// <summary>
        /// Ejecuta el diagnóstico de factibilidad pre-vuelo validando la existencia de ag_aula, correspondencia
        /// exacta de AulaRef, catálogo de asignaturas activo (con diagnóstico contextual accionable, Fase 275)
        /// y capacidad horaria disponible frente a la demanda del pensum.
        /// </summary>
        private async Task EjecutarDiagnosticoPreVueloAsync()
        {
            // Fase 1727: Prevenir reentrada concurrente en diagnóstico pre-vuelo
            if (_diagnosticoEnCurso)
            {
                return;
            }
            _diagnosticoEnCurso = true;

            txtLogAG.Clear();
            Log("================================================================================");
            Log(" INICIANDO DIAGNÓSTICO DE FACTIBILIDAD PRE-VUELO (Algoritmo Genético)");
            Log("================================================================================");

            var traza = CrearTraza("PreVuelo");
            traza.Inicio();
            traza.Paso("P07_DiagnosticoPreVuelo", "Evaluación de factibilidad y consistencia pre-vuelo");

            // Fase 1716, Fase 1722, Fase 1727 & Fase 1729: Bloquear controles y acciones de perfil/árbol, detener temporizador y adquirir semáforo
            bool arbolHabilitado = treeContextoAcademico.Enabled;
            bool perfilesHabilitados = cmbArboles.Enabled;
            bool btnValidarHabilitado = btnValidarFactibilidad.Enabled;
            bool btnEjecutarHabilitado = btnEjecutarAG.Enabled;
            bool btnFiltrarHabilitado = btnFiltrarArbol.Enabled;
            bool btnCargarHabilitado = btnCargarArbol.Enabled;
            bool btnSobrescribirHabilitado = btnSobrescribirArbol.Enabled;

            treeContextoAcademico.Enabled = false;
            cmbArboles.Enabled = false;
            btnValidarFactibilidad.Enabled = false;
            btnEjecutarAG.Enabled = false;
            btnFiltrarArbol.Enabled = false;
            btnCargarArbol.Enabled = false;
            btnSobrescribirArbol.Enabled = false;
            _timerSyncDimensiones.Stop();
            await _semaforoSyncDimensiones.WaitAsync();

            try
            {
                var ctxRes = await ConstruirContextoAsync(CancellationToken.None);
                if (ctxRes.IsFailure)
                {
                    traza.Fallo("P07_ValidarContexto", new Exception(ctxRes.Error), AgCausaFallo.SinAsignaturas);
                    traza.Fin(false, ctxRes.Error);
                    Log($"[ERROR PRE-VUELO] {ctxRes.Error}");
                    MessageBox.Show(this, ctxRes.Error, "Diagnóstico Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _contextoActual = ctxRes.Value;
                if (!await PrepararCoberturaAntesDeCambiosAsync(_contextoActual, CancellationToken.None)) return;
                // Sólo Disponible alcanza la escritura heredada; aprobadas esperan ratificación T4.
                if (_dimensionesPendientes != null)
                {
                    var pend = _dimensionesPendientes.Value;
                    var resSync = await _aulaDimSync.SincronizarConfirmadoAsync(pend.AulaRef, pend.Jornada, pend.HorasDia, pend.DiasVentana, pend.Subunidad, CancellationToken.None);
                    if (resSync.IsFailure)
                    {
                        Log($"[ERROR PRE-VUELO] {resSync.Error}");
                        MessageBox.Show(this, resSync.Error, "Sincronización de dimensiones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _dimensionesPendientes = null;
                    var recarga = await ConstruirContextoAsync(CancellationToken.None);
                    if (recarga.IsFailure) { MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, recarga.Error); return; }
                    _contextoActual = recarga.Value;
                    if (!await PrepararCoberturaAntesDeCambiosAsync(_contextoActual, CancellationToken.None)) return;
                }

                var combinaciones = ObtenerCombinacionesSeleccionadas();
                Log($"[ALCANCE] Diagnóstico: {combinaciones.Count} combinaciones confirmadas.");

                // Fase 213 & 231: Diagnóstico pre-vuelo defensivo
                if (_contextoActual.Aulas == null || _contextoActual.Aulas.Count == 0)
                {
                    traza.Fallo("P07_ValidarAulas", null, AgCausaFallo.SinAulas);
                    traza.Fin(false, "No existe Aula (ag_aula) para la combinación seleccionada");
                    Log("[ERROR PRE-VUELO] No existe Aula (ag_aula) para la combinación seleccionada. Use 'Generar Aulas (N)'.");
                    MessageBox.Show("No existe Aula (ag_aula) para la combinación seleccionada. Use 'Generar Aulas (N)'.", "Diagnóstico Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fases 454, 472, 480 & Fase 645: Validación de anclaje multi-combinación con contexto de alcance y fallback de cohorte a homónimos
                if (combinaciones.Count > 1)
                {
                    var faltantes = combinaciones
                        .Where(c => !_contextoActual.Aulas.Any(a =>
                            string.Equals(a.AulaRef, c.AulaRef, StringComparison.OrdinalIgnoreCase) ||
                            (AgAulaFactory.ExtraerCohorte(c.AulaRef).Length > 0 &&
                             string.Equals(AgAulaFactory.ExtraerCohorte(a.AulaRef),
                                           AgAulaFactory.ExtraerCohorte(c.AulaRef), StringComparison.OrdinalIgnoreCase))))
                        .Select(c => $"{c.AulaRef} [{ModalidadTutoriaCatalog.AEtiqueta(c.Modalidad)}]").ToList();
                    if (faltantes.Count > 0)
                    {
                        var refsAlcance = _contextoActual.Aulas.Select(a => a.AulaRef).Distinct().ToList();
                        var err = $"Faltan aulas (ag_aula) para {faltantes.Count} combinación(es):\n• {string.Join("\n• ", faltantes)}\n" +
                                  $"Alcance: {refsAlcance.Count} aulas | sedes=[{string.Join(",", _contextoActual.SedesAlcance)}] " +
                                  $"programas=[{string.Join(",", _contextoActual.ProgramasAlcance)}]\n" +
                                  $"Refs del alcance: {string.Join(" ; ", refsAlcance)}\n" +
                                  "Acción: ejecute 'Generar Aulas (N)' con las hojas marcadas (multi-sede/multi-programa).";
                        traza.Fallo("P07_ValidarAnclajeAula", new Exception(err), AgCausaFallo.SinAnclajeAula);
                        traza.Fin(false, err);
                        Log($"[ERROR PRE-VUELO] {err}");
                        MessageBox.Show(this, err, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    traza.Dato("AnclajeMultiCombinacion", $"aulas cubiertas={combinaciones.Count}, IdAulaCorrida=0 (multi-grupo)");
                }
                else if (_contextoActual.IdAulaCorrida <= 0)
                {
                    string refEsp = combinaciones.Any() ? combinaciones.First().AulaRef : "seleccionada";
                    traza.Fallo("P07_ValidarAnclajeAula", null, AgCausaFallo.SinAnclajeAula);
                    traza.Fin(false, $"No existe Aula exacta para '{refEsp}'");
                    Log($"[ERROR PRE-VUELO] No existe Aula (ag_aula) exacta para la combinación '{refEsp}'. Use 'Generar Aulas (N)'.");
                    MessageBox.Show(this, $"No existe Aula (ag_aula) exacta para la combinación '{refEsp}'. Use 'Generar Aulas (N)'.", "Diagnóstico Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 275: Diagnóstico pre-vuelo accionable con idprograma, semestre y subunidad
                if (_contextoActual.Asignaturas == null || _contextoActual.Asignaturas.Count == 0)
                {
                    var comb = combinaciones.FirstOrDefault();
                    string semestreRama = comb?.Semestre ?? "";
                    string detalle = $"idprograma={_contextoActual.IdPrograma} semestre='{semestreRama}' subunidad='{_contextoActual.Subunidad}'";
                    traza.Fallo("P07_ValidarAsignaturas", null, AgCausaFallo.SinAsignaturas);
                    traza.Fin(false, $"Sin asignaturas para {detalle}");
                    Log($"[ERROR PRE-VUELO] Sin asignaturas para {detalle}. Verifique el catálogo 'asignaturas' (idprograma/semestre/subunidad) o cargue el pensum.");
                    MessageBox.Show($"El programa seleccionado no tiene asignaturas registradas ({detalle}).\nVerifique el catálogo de asignaturas o cargue el pensum correspondiente.", "Diagnóstico Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 293, 479 & Fase 641: Validar que las aulas del alcance solicitado tengan asignaturas vinculadas en ag_aula_asignatura
                var refsSolicitadas = new HashSet<string>(
                    combinaciones.Select(c => c.AulaRef), StringComparer.OrdinalIgnoreCase);
                var aulasGate = _contextoActual.Aulas != null
                    ? _contextoActual.Aulas.Where(a => refsSolicitadas.Count == 0 || refsSolicitadas.Contains(a.AulaRef)).ToList()
                    : new List<AgAula>();
                var aulasSinMaterias = aulasGate
                    .Where(a => _contextoActual.AulaAsignaturas == null || !_contextoActual.AulaAsignaturas.Any(x => x.IdAula == a.Id))
                    .Select(a => a.AulaRef).ToList();
                var huerfanasExcluidas = _contextoActual.Aulas != null
                    ? _contextoActual.Aulas.Where(a => !refsSolicitadas.Contains(a.AulaRef)
                        && (_contextoActual.AulaAsignaturas == null || !_contextoActual.AulaAsignaturas.Any(x => x.IdAula == a.Id)))
                        .Select(a => a.AulaRef).ToList()
                    : new List<string>();
                if (huerfanasExcluidas.Count > 0)
                    traza.Advertencia("AulasHuerfanasNoSolicitadas",
                        $"{huerfanasExcluidas.Count} aula(s) huérfanas fuera de la selección omitidas: {string.Join(" ; ", huerfanasExcluidas)}");

                if (aulasSinMaterias.Count > 0 || _contextoActual.AulaAsignaturas == null || _contextoActual.AulaAsignaturas.Count == 0)
                {
                    var msg = $"Sin ag_aula_asignatura para {aulasSinMaterias.Count} aula(s):\n• {string.Join("\n• ", aulasSinMaterias)}\nEjecute 'Generar Aulas (N)'.";
                    traza.Fallo("P07_ValidarAulaAsignaturas", new Exception(msg), AgCausaFallo.SinAulaAsignaturas);
                    traza.Fin(false, msg);
                    Log($"[ERROR PRE-VUELO] {msg}");
                    MessageBox.Show(this, msg, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // F655 / Fase 1113 / Fase 1114 / Fase 1116 / Fase 1192: Validación de ventanas y bloques de decisión acumulando aulas afectadas
                if (_contextoActual.Aulas != null && _contextoActual.Bloques != null)
                {
                    var aulasSinDim = new List<string>();
                    var aulasConError = new List<string>();
                    var msgsDetallados = new List<string>();

                    foreach (var aula in aulasGate)
                    {
                        var progV = _contextoActual.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma);
                        var nivelV = AgNivelAcademicoCatalogo.Desde(progV?.Modalidad);
                        int horasEfectivasV = (int)AgUnidadesSemanales.HorasSesionAula(aula, nivelV);

                        // Fase 1114: Bloqueo explícito en pre-vuelo si aula de pregrado tiene configuradas 8 h/día
                        if (aula.HorasDia == 8 && nivelV == AgNivelAcademico.Pregrado)
                        {
                            var errD4 = $"Aula '{aula.AulaRef}': 8 h/día no permitido en PREGRADO (unidad canónica 4 h). Seleccione 4h en el árbol y ejecute 'Generar Aulas'.";
                            // Fase 1120: Causa de fallo específica para incoherencia horas/día ↔ nivel
                            traza.Fallo("P07_HorasDiaNivel", new Exception(errD4), AgCausaFallo.HorasDiaInvalidoNivel);
                            traza.Fin(false, errD4);
                            Log($"[ERROR PRE-VUELO] {errD4}");
                            MessageBox.Show(this, errD4, "Validación Pre-Vuelo Horas/Día", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        var valVentana = AgValidadorVentana.Validar(
                            string.Empty, // Fase 1116: patrón por asignatura se valida en el motor
                            aula.DiasVentana,
                            aula.Jornada,
                            horasEfectivasV,
                            _contextoActual.Bloques);

                        if (valVentana.IsFailure)
                        {
                            var faltantes = AgMensajeDimensionesAula.Faltantes(aula);
                            string txtError = AgMensajeDimensionesAula.Construir(aula, valVentana.Error);
                            msgsDetallados.Add(txtError);

                            if (faltantes.Count > 0)
                            {
                                aulasSinDim.Add(aula.AulaRef);
                            }
                            else
                            {
                                aulasConError.Add(aula.AulaRef);
                            }
                        }
                    }

                    if (msgsDetallados.Count > 0)
                    {
                        const int MaxDetalle = 5;
                        string msgModal = string.Join("\n\n", msgsDetallados.Take(MaxDetalle));
                        if (msgsDetallados.Count > MaxDetalle)
                        {
                            msgModal += $"\n\n... y {msgsDetallados.Count - MaxDetalle} aula(s) más en la misma condición.";
                        }

                        string msgLog = msgModal.Replace("\r", " ").Replace("\n", " ¶ ");
                        var causa = aulasSinDim.Count > 0 ? AgCausaFallo.VentanaAulaVacia : AgCausaFallo.SinBloques;

                        traza.Fallo("P07_ValidarVentanaBloques", new Exception(msgLog), causa);
                        traza.Dato("AulasSinDimensiones", aulasSinDim.Count);
                        traza.Dato("AulaSinDimensionesDetalle", string.Join(",", aulasSinDim.Take(50)));
                        traza.Dato("AulasConErrorVentana", aulasConError.Count);
                        traza.Fin(false, msgLog);
                        Log($"[ERROR PRE-VUELO] {msgLog}");
                        MessageBox.Show(this, msgModal, "Validación Pre-Vuelo Ventana/Jornada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Fase 1233 / Fase 1261 / Fase 1739: pre-vuelo de escalado de ventana unificado
                var resEscaladoDiag = await EvaluarYAplicarEscaladoAsync(traza, CancellationToken.None);
                if (resEscaladoDiag == AgResultadoEscaladoUI.FalloAlAplicar)
                {
                    return;
                }

                // Fase 295, 296, 303 & 321: Validar cobertura de pool docente (Consentimiento de Vacante)
                int totalAsig = _contextoActual.Asignaturas.Count;
                var asignaturasSinPool = _contextoActual.Asignaturas
                    .Where(a => !(_contextoActual.PoolPorAsignatura != null &&
                                  _contextoActual.PoolPorAsignatura.TryGetValue(a.Id, out var pl) &&
                                  pl.Any(p => p.Activo == 1)))
                    .ToList();
                int asigSinPool = asignaturasSinPool.Count;
                int asigConPool = totalAsig - asigSinPool;

                traza.Dato("CoberturaDocente", $"{asigConPool}/{totalAsig}");

                if (asigSinPool > 0)
                {
                    string listaSin = AgFormatoAsignatura.BulletList(asignaturasSinPool);
                    traza.Dato("AsignaturasSinDocente", listaSin.Replace("\n", " "));
                    traza.Causa(AgCausaFallo.SinPoolDocente, $"{asigSinPool}/{totalAsig} → VACANTE");
                    Log($"[PRE-VUELO] {asigSinPool}/{totalAsig} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSin}");

                    var dialogRes = MessageBox.Show(this,
                        $"Hay {asigSinPool} de {totalAsig} asignaturas sin docente habilitado en el pool:\n• {listaSin}\n\n" +
                        "Se marcarán como VACANTE (es_vacante=1) con penalización suave y podrá resolverlas después en el Panel de Vacantes.\n\n" +
                        "¿Continuar con la optimización?",
                        "Diagnóstico Pre-Vuelo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (dialogRes != DialogResult.Yes)
                    {
                        traza.Fin(false, "Cancelado por cobertura incompleta");
                        Log($"[DIAGNÓSTICO PRE-VUELO] Cancelado por el usuario debido a {asigSinPool}/{totalAsig} asignaturas sin pool.");
                        _contextoActual?.AsignaturasVacantesConsentidas?.Clear();
                        return;
                    }

                    _contextoActual.AsignaturasVacantesConsentidas = new HashSet<int>(asignaturasSinPool.Select(a => a.Id));
                    Log($"[DIAGNÓSTICO PRE-VUELO] Aceptado por el usuario: {asigSinPool}/{totalAsig} asignaturas pasarán a VACANTE.");
                }
                else
                {
                    _contextoActual?.AsignaturasVacantesConsentidas?.Clear();
                }

                var factibilidadResult = await _validadorVentana.ValidarFactibilidadAsync(_contextoActual, CancellationToken.None);

                if (factibilidadResult.IsFailure)
                {
                    traza.Fallo("P07_ValidarFactibilidad", new Exception(factibilidadResult.Error), AgCausaFallo.VentanaInfactible);
                    traza.Fin(false, factibilidadResult.Error);
                    Log($"[ERROR] Fallo al validar: {factibilidadResult.Error}");
                    return;
                }

                var inf = factibilidadResult.Value;
                traza.Dato("CursosAProgramar", inf.TotalCursosProgramar);
                traza.Dato("HorasRequeridas", $"{inf.TotalHorasRequeridas}h");
                traza.Dato("CapacidadHorariaAulas", $"{inf.TotalCapacidadHorariaAulas}h");
                traza.Dato("DocentesDisponibles", inf.TotalDocentesDisponibles);

                Log($"• Cursos a programar: {inf.TotalCursosProgramar}");
                Log($"• Horas académicas requeridas: {inf.TotalHorasRequeridas}h");
                Log($"• Capacidad máxima de aulas en la ventana: {inf.TotalCapacidadHorariaAulas}h");
                Log($"• Docentes disponibles en catálogo: {inf.TotalDocentesDisponibles}");

                if (inf.Advertencias.Count > 0)
                {
                    Log("\n[ADVERTENCIAS DETECTADAS]");
                    foreach (var adv in inf.Advertencias) Log($"  ⚠ {adv}");
                }

                if (inf.BloqueosCriticos.Count > 0)
                {
                    traza.Resultado(false, $"{inf.BloqueosCriticos.Count} bloqueos críticos detectados");
                    traza.Fin(false, "Bloqueos críticos de factibilidad");
                    Log("\n[BLOQUEOS CRÍTICOS (Inviable)]");
                    foreach (var blk in inf.BloqueosCriticos) Log($"  ❌ {blk}");
                    MessageBox.Show("El diagnóstico detectó bloqueos críticos de factibilidad. Revise el registro.", "Inviable", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    traza.Resultado(true, "Diagnóstico pre-vuelo superado exitosamente");
                    traza.Fin(true, "Factible");
                    Log("\n✓ DIAGNÓSTICO EXITOSO: El espacio de búsqueda y los recursos son matemáticamente viables.");
                    MessageBox.Show("Diagnóstico de factibilidad superado con éxito. Puede proceder a optimizar.", "Factibilidad OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                traza.Fallo("P07_EjecutarDiagnosticoPreVueloAsync", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                Log($"[EXCEPCIÓN] {ex.Message}");
            }
            finally
            {
                _semaforoSyncDimensiones.Release();
                treeContextoAcademico.Enabled = arbolHabilitado;
                cmbArboles.Enabled = perfilesHabilitados;
                btnValidarFactibilidad.Enabled = btnValidarHabilitado;
                btnEjecutarAG.Enabled = btnEjecutarHabilitado;
                btnFiltrarArbol.Enabled = btnFiltrarHabilitado;
                btnCargarArbol.Enabled = btnCargarHabilitado;
                btnSobrescribirArbol.Enabled = btnSobrescribirHabilitado;
                _diagnosticoEnCurso = false;
            }
        }

        private enum TipoPreparacionPlan
        {
            AlcanceInvalido, ErrorTecnico, Cancelada, NoVerificable, RatificacionPendiente,
            CandidatoListo, CoberturaInsuficiente, FlushFallido, Rechazado
        }

        private async Task<bool> PrepararCoberturaAntesDeCambiosAsync(AgContexto contexto, CancellationToken ct)
        {
            contexto.ManifiestoCobertura = null;
            contexto.PlanAprobadoSnapshot = null;
            contexto.ValidacionPlanAprobado = null;
            contexto.ReiniciarCandidatosCobertura();
            try
            {
                ct.ThrowIfCancellationRequested();
                var seleccion = ObtenerCombinacionesSeleccionadas();
                var ids = new List<int>();
                if (seleccion.Count == 0)
                    return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "SELECCION_EXPLICITA_REQUERIDA");
                foreach (var c in seleccion)
                {
                    var exactas = contexto.Aulas.Where(a => a.IdPeriodo == c.IdPeriodo && a.IdSede == c.IdSede
                        && a.IdPrograma == c.IdPrograma && string.Equals(a.AulaRef, c.AulaRef, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (exactas.Count != 1 || ids.Contains(exactas[0].Id))
                        return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "SELECCION_NO_EXACTA:" + c.AulaRef);
                    ids.Add(exactas[0].Id);
                }
                if (contexto.Aulas.Count != ids.Count)
                    return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "AULAS_FUERA_DE_SELECCION");
                // Fase 1850: Validar alcance ANTES de sincronizar dimensiones y proteger con semáforo
                if (_dimensionesPendientes.HasValue)
                {
                    var pend = _dimensionesPendientes.Value;
                    bool enAlcance = contexto.Aulas.Any(a =>
                        (a.Estado == AgEstadoAula.Disponible || contexto.PuedeGenerarAula(a.Id))
                        && string.Equals(a.AulaRef, pend.AulaRef, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(a.Subunidad, pend.Subunidad, StringComparison.OrdinalIgnoreCase));
                    if (!enAlcance)
                        return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "DIMENSION_PENDIENTE_FUERA_DE_ALCANCE");

                    _timerSyncDimensiones.Stop();
                    await _semaforoSyncDimensiones.WaitAsync(ct);
                    try
                    {
                        var resFlush = await _aulaDimSync.SincronizarConfirmadoAsync(
                            pend.AulaRef, pend.Jornada, pend.HorasDia, pend.DiasVentana,
                            pend.Subunidad, ct);
                        if (resFlush.IsFailure)
                            return MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, "ERROR_FLUSH_DIMENSIONES:" + resFlush.Error);
                        _dimensionesPendientes = null;
                    }
                    finally
                    {
                        _semaforoSyncDimensiones.Release();
                    }
                }
                // El límite es la fecha de fin de la ventana semestral seleccionada, no una fecha fija.
                var ventana = contexto.Ventana;
                if (ventana == null || ventana.FechaFin.Date < ventana.FechaInicio.Date)
                    return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "VENTANA_SEMESTRAL_INVALIDA");
                List<int> territoriales;
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    territoriales = (await db.QueryAsync<int>(new CommandDefinition(
                        "SELECT id FROM territoriales WHERE UPPER(TRIM(subunidad))=UPPER(TRIM(@Sub)) ORDER BY id",
                        new { Sub = contexto.Subunidad }, cancellationToken: ct))).ToList();
                }
                if (territoriales.Count != 1)
                    return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, "TERRITORIAL_AUSENTE_O_AMBIGUA");
                var manifiesto = new AgConstructorManifiestoCobertura().Construir(contexto, ids, territoriales[0],
                    contexto.Subunidad, ventana.FechaInicio, ventana.FechaFin);
                if (manifiesto.IsFailure) return MostrarFalloPlan(TipoPreparacionPlan.AlcanceInvalido, manifiesto.Error);
                contexto.ManifiestoCobertura = manifiesto.Value;
                if (!manifiesto.Value.Relaciones.Any(r => r.EstadoAula == (int)AgEstadoAula.Aprobado))
                {
                    // Sin lectura canónica obligatoria para Disponible; conserva su escalado/generación actuales.
                    contexto.ValidacionPlanAprobado = new AgValidacionPlanAprobado(null,
                        manifiesto.Value.Relaciones.GroupBy(r => r.Clave.IdAula).Select(g => new AgValidacionAulaPlan(g.Key,
                            AgEstadoValidacionPlan.NoRequerida, g.Select(r => new AgValidacionRelacionPlan(r, 0, 0, new string[0])))), new string[0]);
                    return true;
                }
                var lectura = await _planAprobadoRepository.ObtenerAsync(manifiesto.Value, ct);
                ct.ThrowIfCancellationRequested();
                if (lectura.IsFailure)
                    return MostrarFalloPlan(lectura.Error == "LECTURA_PLAN_CANCELADA" ? TipoPreparacionPlan.Cancelada
                        : TipoPreparacionPlan.ErrorTecnico, lectura.Error);
                // El lector puede descubrir extras únicamente productivas; no se aceptan cambios en
                // demanda, ventana ni claves AG ya declaradas al ligar su huella antes del consentimiento.
                if (lectura.Value == null || lectura.Value.Manifiesto.Inicio != manifiesto.Value.Inicio
                    || lectura.Value.Manifiesto.Fin != manifiesto.Value.Fin
                    || !lectura.Value.Manifiesto.Relaciones.Select(r => r.Clave.Valor)
                        .SequenceEqual(manifiesto.Value.Relaciones.Select(r => r.Clave.Valor))
                    || manifiesto.Value.RelacionesExtra.Any(k => !lectura.Value.Manifiesto.EsExtra(k)))
                    return MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, "LECTURA_PLAN_MANIFIESTO_NO_COINCIDE");
                contexto.ManifiestoCobertura = lectura.Value.Manifiesto;
                contexto.PlanAprobadoSnapshot = lectura.Value;
                var validacion = new AgValidadorPlanAprobado().Validar(lectura.Value, contexto);
                contexto.ValidacionPlanAprobado = validacion;
                foreach (var d in validacion.Diagnosticos) Log("[PLAN CANÓNICO] " + d);
                // Decisión humana 2026-11-15 «B: extras como ocupación fija»: las extras declaradas no
                // bloquean; se conservan inmutables y ocupan aula/docente sin sumar demanda requerida.
                foreach (var extra in lectura.Value.RelacionesExtra)
                    Log($"[PLAN CANÓNICO] extra aulas={string.Join(",", extra.ClavesCandidatas.Select(k => k.IdAula))} vinculacion={extra.IdVinculacion} sesiones={extra.Sesiones.Count} conservada=inmutable sin_replanificar");
                foreach (var aula in validacion.Aulas.Where(a => a.Estado != AgEstadoValidacionPlan.NoRequerida))
                {
                    Log($"[PLAN CANÓNICO] aula={aula.IdAula} estado={aula.Estado} esperadas={aula.HorasEsperadas} preservadas={aula.HorasPreservadas} déficit={aula.Deficit} exceso={aula.Exceso} fuera={aula.HorasFueraVentana}");
                    foreach (var r in aula.Relaciones)
                    {
                        // El resumen por aula ya quedó registrado; el detalle sólo se emite cuando exige revisión.
                        if (r.Diagnosticos.Count == 0 && r.Deficit <= 0.001 && r.Exceso <= 0.001) continue;
                        Log($"[PLAN CANÓNICO] aula={aula.IdAula} asignatura={r.Clave.IdAsignatura} esperadas={r.HorasEsperadas} válidasSinRatificar={r.HorasValidasEnVentana} déficit={r.Deficit} exceso={r.Exceso} fuera={r.HorasFueraVentana}");
                        foreach (var d in r.Diagnosticos) Log("[PLAN CANÓNICO] " + d);
                    }
                }
                // Ofrecer todas las aulas; ninguna confirmación habilita todavía la generación/promoción T5/T6.
                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Validando, "SIMULANDO_PROPUESTAS_APROBADAS",
                    "Estado: Evaluando escalado informativo de aulas aprobadas.");
                var propuestas = new AgSimuladorEscaladoAprobadas().Evaluar(contexto, ct);
                var aprobadas = contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Aprobado).OrderBy(a => a.Id).ToList();
                // Regla general: una aprobada sin fila productiva para su cohorte no tiene origen que
                // conservar; su demanda es obligatoria y se planifica como nueva. Sólo las aulas con origen
                // reconciliado pasan por ratificación, escalado de conservación o bloqueo.
                var pendientesProgramacion = aprobadas.Where(a => !contexto.TieneOrigenConservable(a.Id)).ToList();
                var conservables = aprobadas.Where(a => contexto.TieneOrigenConservable(a.Id)).ToList();
                foreach (var aula in pendientesProgramacion)
                {
                    Log($"[PLAN CANÓNICO] aula={aula.Id} estado=PendienteDeProgramar origen=AUSENTE acción=planificar");
                }
                if (_confirmadorPlanAprobado == null)
                    return MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, "PUBLICACION_NO_DISPONIBLE");
                var intentoRatificacion = _intentoOptimizacion ?? Guid.NewGuid();
                var terminales = new List<string>();
                var terminal = TipoPreparacionPlan.CandidatoListo;
                foreach (var aula in conservables)
                {
                    ct.ThrowIfCancellationRequested();
                    if (_aulasConfirmadasEnIntento.Contains(aula.Id))
                    {
                        terminales.Add("aula=" + aula.Id + ":PASO_YA_CONFIRMADO_EN_INTENTO");
                        terminal = TipoPreparacionPlan.Rechazado;
                        continue;
                    }
                    var propuesta = propuestas.SingleOrDefault(p => p.IdAula == aula.Id);
                    var diagnostico = validacion.Aulas.Single(v => v.IdAula == aula.Id);
                    if (diagnostico.Estado == AgEstadoValidacionPlan.PlanNoVerificable)
                    {
                        terminales.Add("aula=" + aula.Id + ":PLAN_NO_VERIFICABLE:" +
                            AgDescriptorDiagnosticoPlan.Describir(diagnostico, validacion.Diagnosticos));
                        terminal = TipoPreparacionPlan.Rechazado;
                        continue;
                    }
                    if (diagnostico.Deficit > 0.001 && propuesta == null)
                    {
                        terminales.Add("aula=" + aula.Id + ":I5_COBERTURA_INSUFICIENTE: sin escalón verificable");
                        terminal = TipoPreparacionPlan.CoberturaInsuficiente;
                        continue;
                    }
                    string detalle = propuesta == null
                        ? "Plan canónico completo: sólo ratificación, sin propuesta de cambio de días.\n" +
                            string.Join("\n", diagnostico.Relaciones.Select(r => r.Clave.Valor + " " + r.HorasEsperadas + "h"))
                        : ConstruirDetallePropuestaAprobada(propuesta);
                    int extrasConservadas = lectura.Value.RelacionesExtra.Count(r => r.ClavesCandidatas.Count == 1
                        && r.ClavesCandidatas[0].IdAula == aula.Id);
                    if (extrasConservadas > 0)
                        detalle += "\nExtras conservadas: " + extrasConservadas +
                            " relación(es) no requerida(s) que se conservan inmutables y ocupan aula y docente; " +
                            "no se replantean ni suman demanda requerida.";
                    Log(detalle);
                    string huellaSql = await _confirmadorPlanAprobado.CapturarConfiguracionAsync(contexto, aula.Id, propuesta, ct);
                    if (huellaSql == null)
                    {
                        terminales.Add("aula=" + aula.Id + ":CONFIGURACION_O_PUBLICACION_NO_DISPONIBLE");
                        terminal = TipoPreparacionPlan.FlushFallido;
                        continue;
                    }
                    var archivada = await _confirmadorPlanAprobado.BuscarArchivadoAsync(contexto, aula.Id, propuesta, huellaSql, ct);
                    if (archivada != null)
                        detalle += "\nOPERACION_ARCHIVADA:" + archivada.Operacion +
                            "; ratifique de nuevo el mismo efecto para recuperar publicación. No se autoriza otro paso.";
                    using (var formulario = new AgFormConfirmacionPlanAprobado(aula.Id, lectura.Value, propuesta, detalle, archivada))
                    {
                        EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.RequiereDecision,
                            "RATIFICACION_PENDIENTE", "Estado: Ratificación y paso candidato pendientes de confirmación.");
                        if (formulario.ShowDialog(this) != DialogResult.OK)
                        {
                            terminales.Add("aula=" + aula.Id + ":I5_REVISION_RECHAZADA; sin escrituras");
                            terminal = TipoPreparacionPlan.Rechazado;
                            continue;
                        }
                        var confirmada = await _confirmadorPlanAprobado.ConfirmarAsync(contexto, propuesta, aula.Id,
                            archivada?.Intento ?? intentoRatificacion, archivada?.Operacion ?? Guid.NewGuid(), huellaSql, formulario.Motivo,
                            formulario.Ratifica, formulario.AutorizaPaso, formulario.LegadoConsentido, ct);
                        if (confirmada.Codigo != "CANDIDATO_PUBLICADO" && confirmada.Codigo != "REPLAY_PUBLICADO")
                        {
                            terminales.Add("aula=" + aula.Id + ":" + confirmada.Codigo);
                            terminal = confirmada.Codigo.StartsWith("FLUSH_FALLIDO", StringComparison.Ordinal)
                                ? TipoPreparacionPlan.FlushFallido : TipoPreparacionPlan.Rechazado;
                            continue;
                        }
                        _aulasConfirmadasEnIntento.Add(aula.Id);
                        contexto.RegistrarCandidatoCobertura(confirmada.Candidato);
                        int residual = propuesta == null ? 0 : ReevaluarCandidatoConfirmado(contexto, propuesta,
                            formulario.LegadoConsentido, ct);
                        if (residual > 0)
                        {
                            terminales.Add("aula=" + aula.Id + ":I5_COBERTURA_INSUFICIENTE: " + residual + " sesiones; siguiente paso sólo en otra corrida.");
                            terminal = TipoPreparacionPlan.CoberturaInsuficiente;
                        }
                        else terminales.Add("aula=" + aula.Id + ":" + (propuesta == null
                            ? "RATIFICACION_SOLO_CONSERVACION_PENDIENTE_T5" : "CANDIDATO_COMPLETO_PENDIENTE_T5"));
                    }
                }
                if (terminal == TipoPreparacionPlan.CandidatoListo && contexto.CandidatosCobertura.Count == conservables.Count)
                {
                    var conservadas = new AgProyectorSolucionCobertura().Conservar(contexto);
                    if (conservadas.IsFailure)
                        return MostrarFalloPlan(TipoPreparacionPlan.NoVerificable, conservadas.Error);
                    contexto.SesionesConservadas = conservadas.Value;
                    // Se omite autocruce productivo sólo por ID canónico; las demás ocupaciones permanecen.
                    contexto.ExcluirOcupacionesCanonicas(contexto.PlanAprobadoSnapshot.Relaciones
                        .Where(r => r.ClavesCandidatas.Count == 1 && contexto.CandidatosCobertura.Any(c => c.IdAula == r.ClavesCandidatas[0].IdAula))
                        .SelectMany(r => r.Sesiones).Select(s => s.IdSesion));
                    // En alcance exclusivamente conservado, comprobar la unión ahora; la ruta de generación
                    // validará después de materializar cuando haya relaciones por replanificar.
                    if (contexto.ManifiestoCobertura.Relaciones.All(r => conservadas.Value.Any(s => s.Clave.Valor == r.Clave.Valor)))
                    {
                        var cobertura = new AgValidadorCoberturaSeleccion().Validar(contexto.ManifiestoCobertura,
                            conservadas.Value, Array.Empty<AgSesionCobertura>(), contexto.IdTerceroVacante, contexto.Festivos, contexto.AsignaturasVacantesConsentidas);
                        if (cobertura.IsFailure) return MostrarFalloPlan(TipoPreparacionPlan.NoVerificable, cobertura.Error);
                    }
                    return true;
                }
                return MostrarFalloPlan(terminal, string.Join("; ", terminales));
            }
            catch (OperationCanceledException) { return MostrarFalloPlan(TipoPreparacionPlan.Cancelada, "LECTURA_PLAN_CANCELADA"); }
            catch (Exception ex) { return MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, "PREPARACION_PLAN_ERROR:" + ex.GetType().Name); }
        }

        /// <summary>Replanifica una sola vez el candidato aislado, con exactamente el LEGADO consentido.</summary>
        private static int ReevaluarCandidatoConfirmado(AgContexto origen, AgPropuestaEscaladoAprobada propuesta,
            IReadOnlyCollection<int> legado, CancellationToken ct)
        {
            var copia = AgEscaladorRanuraAula.ClonarContextoParaSimulacion(origen);
            var aula = copia.Aulas.Single(a => a.Id == propuesta.IdAula);
            aula.DiasVentana = propuesta.Hacia;
            copia.AulaAsignaturas = copia.AulaAsignaturas.Where(r => r.Requerida == 1).ToList();
            copia.ConstruirIndices();
            foreach (var detalle in propuesta.Relaciones.Where(r => r.RedistribuyeLegado
                && legado.Contains(r.Clave.IdAsignatura)))
            {
                var relacion = copia.AulaAsignaturaPorAulaAsig[(aula.Id, detalle.Clave.IdAsignatura)];
                relacion.OrigenPatron = "AUTOMATICO";
                relacion.DiasPatron = null;
            }
            double unidad = AgUnidadesSemanales.HorasSesionAula(aula,
                AgNivelAcademicoCatalogo.Desde(copia.Programas.Single(p => p.Id == aula.IdPrograma).Modalidad));
            var patrones = copia.AulaAsignaturaPorAulaAsig.Values.Where(r => r.IdAula == aula.Id
                && !string.Equals(r.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(r.DiasPatron)).ToDictionary(r => r.IdAsignatura, r => r.DiasPatron);
            var resultado = new AgPlanificadorOleadas().ConstruirEvaluacion(AgEvaluadorDemandaVentana.Demandas(copia, aula),
                aula.DiasVentana, copia.Ventana.FechaInicio, copia.Ventana.FechaFin, unidad,
                copia.EsFestivo, copia.Bloques, aula.Jornada, patrones, ct);
            if (resultado.IsFailure) throw new InvalidOperationException("REEVALUACION_CANDIDATO:" + resultado.Error);
            return propuesta.Relaciones.Sum(r => Math.Max(0, (int)Math.Ceiling(
                origen.ValidacionPlanAprobado.Aulas.Single(a => a.IdAula == aula.Id).Relaciones
                    .Single(v => v.Clave.IdAsignatura == r.Clave.IdAsignatura).HorasEsperadas / unidad)
                - resultado.Value.Planes.Where(p => p.IdAsignatura == r.Clave.IdAsignatura).Sum(p => p.Sesiones.Count)));
        }

        internal static string ConstruirDetallePropuestaAprobada(AgPropuestaEscaladoAprobada propuesta)
        {
            var texto = new StringBuilder();
            texto.AppendLine($"Aula {propuesta.IdAula} — {propuesta.AulaRef}: {propuesta.Desde} → {(string.IsNullOrEmpty(propuesta.Hacia) ? "sin siguiente escalón" : propuesta.Hacia)}");
            string[] dias = { "", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo" };
            texto.AppendLine("Días añadidos: " + string.Join(", ", propuesta.DiasAnadidos.Select(d => dias[d])));
            texto.AppendLine($"Fecha final: {propuesta.FechaLimite:dd/MM/yyyy}; límite = fin de la ventana semestral seleccionada. Terminal: {propuesta.Terminal}.");
            texto.AppendLine("Sólo capacidad de calendario; no acredita disponibilidad docente ni ausencia de conflictos externos.");
            texto.AppendLine("No autoriza revisión, ratificación, redistribución ni aplicación. T4 pendiente; horario aprobado intacto.");
            if (propuesta.BusquedaAgotada) texto.AppendLine("Búsqueda acotada agotada: el residual no acredita imposibilidad.");
            foreach (var r in propuesta.Relaciones)
            {
                texto.AppendLine($"Relación: {r.Clave.Valor}; déficit canónico: {r.DeficitCanonicoHoras:0.##} h.");
                texto.AppendLine($"Patrón afectado {r.OrigenPatron}: '{r.PatronAntes}' → sin LEGADO '{r.PatronSinRedistribucion}' / con LEGADO '{r.PatronConRedistribucion}'. Redistribución LEGADO simulada: {r.RedistribuyeLegado}; FIJO intacto.");
                texto.AppendLine("Sesiones canónicas anteriores: " + string.Join(", ", r.SesionesOrigen.Select(s => $"{s.Fecha:dd/MM/yyyy} {s.HoraInicio}-{s.HoraFin} ({s.HorasLectivas:0.##}h)")));
                texto.AppendLine("Simulación actual: " + FechasPropuesta(r.Antes));
                texto.AppendLine($"Siguiente sin redistribución LEGADO: {FechasPropuesta(r.SinRedistribucion)}; residual {r.ResidualSinRedistribucion} sesiones.");
                texto.AppendLine($"Sesiones recuperadas sin LEGADO (ganancia neta): {r.GananciaNetaSinRedistribucion}. Fechas nuevas, incluidas reubicaciones: {FechasPropuesta(r.RecuperadasSinRedistribucion)}");
                texto.AppendLine($"Siguiente con redistribución LEGADO hipotética: {FechasPropuesta(r.ConRedistribucion)}; residual {r.ResidualConRedistribucion} sesiones.");
                texto.AppendLine($"Sesiones recuperadas con LEGADO (ganancia neta): {r.GananciaNetaConRedistribucion}. Fechas nuevas, incluidas reubicaciones: {FechasPropuesta(r.RecuperadasConRedistribucion)}");
            }
            return texto.ToString();
        }

        private static string FechasPropuesta(IEnumerable<AgSesionPropuestaAprobada> sesiones) =>
            string.Join(", ", sesiones.Select(s => $"{s.Fecha:dd/MM/yyyy} {s.Franja} ({s.Horas:0.##}h)"));

        private bool MostrarFalloPlan(TipoPreparacionPlan tipo, string detalle)
        {
            string codigo = tipo == TipoPreparacionPlan.NoVerificable ? "PLAN_NO_VERIFICABLE"
                : tipo == TipoPreparacionPlan.RatificacionPendiente ? "REQUIERE_RATIFICACION"
                : tipo == TipoPreparacionPlan.CandidatoListo ? "CANDIDATO_LISTO_PENDIENTE_T5"
                : tipo == TipoPreparacionPlan.CoberturaInsuficiente ? "I5_COBERTURA_INSUFICIENTE"
                : tipo == TipoPreparacionPlan.FlushFallido ? "FLUSH_FALLIDO"
                : tipo == TipoPreparacionPlan.Rechazado ? "REVISION_RECHAZADA" : tipo.ToString();
            EstablecerEstadoOptimizacion(tipo == TipoPreparacionPlan.Cancelada ? EstadoVisualOptimizacion.Cancelado
                : tipo == TipoPreparacionPlan.ErrorTecnico ? EstadoVisualOptimizacion.Error : EstadoVisualOptimizacion.Bloqueado,
                codigo + ":" + detalle, "Estado: " + codigo + ". Revise el diagnóstico.");
            Log("[PREPARACIÓN] " + codigo + ":" + detalle);
            MessageBox.Show(this, codigo + "\n" + detalle, "Preparación de cobertura", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        private enum AgResultadoEscaladoUI
        {
            SinCambios,
            AplicadoExitosamente,
            RechazadoPorUsuario,
            FalloAlAplicar
        }

        /// <summary>
        /// Evalúa si la demanda de las aulas excede la capacidad o satura los carriles de su ventana actual.
        /// Si detecta saturación, presenta un diálogo modal detallado con la ruta propuesta, ranuras antes/después,
        /// demanda requerida y demanda pendiente. Si el usuario confirma, aplica el escalado transaccionalmente,
        /// actualiza la memoria (aulas y patrones automáticos en ag_aula_asignatura), recarga las dimensiones del perfil activo
        /// en el árbol de forma segura y ejecuta una revalidación de factibilidad.
        /// </summary>
        private bool _escaladoEnCurso;

        private async Task<AgResultadoEscaladoUI> EvaluarYAplicarEscaladoAsync(AgDiagnosticoTraza traza, CancellationToken ct)
        {
            if (_escaladoEnCurso) { return AgResultadoEscaladoUI.FalloAlAplicar; }
            _escaladoEnCurso = true;
            try { return await EvaluarYAplicarEscaladoInternoAsync(traza, ct); }
            finally { _escaladoEnCurso = false; }
        }

        private async Task<AgResultadoEscaladoUI> EvaluarYAplicarEscaladoInternoAsync(AgDiagnosticoTraza traza, CancellationToken ct)
        {
            var escalador = new AgEscaladorRanuraAula(_connectionFactory);
            int idArbolActivo = cmbArboles.SelectedItem is AgArbol perfilActivo ? (int)perfilActivo.Id : 0;
            var contextoPropuesta = _contextoActual;

            try
            {
                // Paso 1: Evaluar plan fechado primero
                var decisionesEscalado = await Task.Run(() => escalador.Evaluar(contextoPropuesta, idArbolActivo, ct), ct);
                if (!decisionesEscalado.EsValido)
                {
                    MarcarErrorPreparacion("ESCALADO_EVALUACION_" + decisionesEscalado.Motivo);
                    MessageBox.Show(this,
                        $"No se pudo evaluar la propuesta inicial: {decisionesEscalado.Mensaje}\n\nMotivo: {decisionesEscalado.Motivo}\nAula: {decisionesEscalado.IdAulaRechazada}\nAsignatura: {decisionesEscalado.IdAsignaturaRechazada}\n\nDebe revisar la configuración.",
                        "Evaluación no aplicable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    Log($"[ESCALADO-WARN] Evaluación inicial rechazada: motivo={decisionesEscalado.Motivo} aula={decisionesEscalado.IdAulaRechazada} asig={decisionesEscalado.IdAsignaturaRechazada} mensaje={decisionesEscalado.Mensaje}");
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                }

                // Paso 2: Obtener candidatas exclusivamente desde el diagnóstico inicial sellado (ReclasificacionesPropuestas)
                var candidatosReclasificacion = decisionesEscalado
                    .Where(d => d != null && d.ReclasificacionesPropuestas != null)
                    .SelectMany(d => d.ReclasificacionesPropuestas)
                    .GroupBy(r => (r.IdAula, r.IdAsignatura))
                    .Select(g => g.First().Clone())
                    .ToList();

                bool hayDecisionesAplicables = decisionesEscalado.Any(d => d != null && d.Aplica);
                if (!hayDecisionesAplicables && candidatosReclasificacion.Count == 0)
                {
                    return AgResultadoEscaladoUI.SinCambios;
                }

                if (candidatosReclasificacion.Count > 0)
                {
                    using (var dlg = new FormAutorizarReclasificacion(candidatosReclasificacion))
                    {
                        MarcarDecisionPreparacion();
                        if (dlg.ShowDialog(this) != DialogResult.OK)
                        {
                            return AgResultadoEscaladoUI.RechazadoPorUsuario;
                        }
                        ReanudarValidacionPreparacion();
                    }
                }

                var autorizadas = candidatosReclasificacion.Where(c => c.Autorizado).ToList();
                if (autorizadas.Count > 0)
                {
                    var validacion = escalador.ValidarSelecciones(contextoPropuesta, autorizadas, candidatosReclasificacion, ct);
                    if (!validacion.EsValido)
                    {
                        MarcarErrorPreparacion("ESCALADO_SELECCION_" + validacion.Motivo);
                        MessageBox.Show(this,
                            $"No se pudo autorizar la reclasificación: {validacion.Mensaje}\n\nMotivo: {validacion.Motivo}\nAula: {validacion.IdAulaRechazada}\nAsignatura: {validacion.IdAsignaturaRechazada}\n\nDebe recalcular la propuesta para continuar.",
                            "Reclasificación no aplicable",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        Log($"[ESCALADO-WARN] Validación de reclasificación rechazada: motivo={validacion.Motivo} aula={validacion.IdAulaRechazada} asig={validacion.IdAsignaturaRechazada} mensaje={validacion.Mensaje}");
                        return AgResultadoEscaladoUI.FalloAlAplicar;
                    }
                }

                // Recalcular únicamente al simular cambios autorizados; de lo contrario reutilizar la evaluación inicial
                AgEscaladorRanuraAula.AgResultadoEvaluacionEscalado propuestaDefinitiva;
                if (autorizadas.Count > 0)
                {
                    propuestaDefinitiva = await Task.Run(() => escalador.Evaluar(contextoPropuesta, idArbolActivo, autorizadas, candidatosReclasificacion, ct), ct);
                    if (!propuestaDefinitiva.EsValido)
                    {
                        MarcarErrorPreparacion("ESCALADO_SIMULACION_" + propuestaDefinitiva.Motivo);
                        MessageBox.Show(this,
                            $"No se pudo aplicar la simulación de reclasificación: {propuestaDefinitiva.Mensaje}\n\nMotivo: {propuestaDefinitiva.Motivo}\nAula: {propuestaDefinitiva.IdAulaRechazada}\nAsignatura: {propuestaDefinitiva.IdAsignaturaRechazada}\n\nDebe recalcular la propuesta para continuar.",
                            "Simulación no aplicable",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        Log($"[ESCALADO-WARN] Simulación de reclasificación rechazada: motivo={propuestaDefinitiva.Motivo} aula={propuestaDefinitiva.IdAulaRechazada} asig={propuestaDefinitiva.IdAsignaturaRechazada} mensaje={propuestaDefinitiva.Mensaje}");
                        return AgResultadoEscaladoUI.FalloAlAplicar;
                    }
                }
                else
                {
                    propuestaDefinitiva = decisionesEscalado;
                }

                // Separar candidatos de decisiones aplicables; selección vacía devuelve SinCambios
                var decisionesAplicables = propuestaDefinitiva.Where(d => d != null && d.Aplica).ToList();
                if (decisionesAplicables.Count == 0)
                {
                    Log("[ESCALADO] Sin decisiones aplicables tras la selección (no hay escalado ni reclasificaciones autorizadas).");
                    return AgResultadoEscaladoUI.SinCambios;
                }

                // Paso 4 - Presentación del resultado definitivo al usuario
                var lineasDetalle = new List<string>();
                int totalAsigsReset = 0;
                foreach (var d in decisionesAplicables)
                {
                    var aulaObj = _contextoActual?.Aulas?.FirstOrDefault(a => string.Equals(a.AulaRef, d.AulaRef, StringComparison.OrdinalIgnoreCase));
                    var progObj = _contextoActual?.Programas?.FirstOrDefault(p => p.Id == aulaObj?.IdPrograma);
                    var nivelObj = AgNivelAcademicoCatalogo.Desde(progObj?.Modalidad);
                    double unidad = aulaObj != null ? AgUnidadesSemanales.HorasSesionAula(aulaObj, nivelObj) : 4.0;
                    string rotuloHacia = AgEscaleraVentanas.Rotular(d.Hacia, unidad);

                    int pendiente = d.DemandaPendiente;
                    string textoPendiente = pendiente > 0
                        ? $" [Demanda pendiente: {pendiente} sesiones; revisar restricciones o siguiente escalón]"
                        : " [Sin pendientes según el plan evaluado]";

                    int asigsAutoAula = d.IdAsignaturasBottleneck != null && d.IdAsignaturasBottleneck.Count > 0
                        ? d.IdAsignaturasBottleneck.Count
                        : (_contextoActual?.AulaAsignaturas != null && aulaObj != null
                            ? _contextoActual.AulaAsignaturas.Count(x => x.IdAula == aulaObj.Id && string.Equals(x.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                            : 0);
                    totalAsigsReset += asigsAutoAula;

                    string plazoAprobado = d.FechaLimite != default(DateTime) ? d.FechaLimite.ToString("yyyy-MM-dd") : "no definido";
                    string finProyectado = d.FechaFinProyectada.HasValue
                        ? d.FechaFinProyectada.Value.ToString("yyyy-MM-dd") : "sin proyección dentro del horizonte diagnóstico";
                    string linea = $"• {d.AulaRef}: {d.Desde} ({d.CapacidadAntes} ranuras) → {d.Hacia} [{rotuloHacia}] ({d.CapacidadDespues} ranuras); {d.SemanasRequeridas} semanas; plazo aprobado: {plazoAprobado}; fin proyectado {finProyectado}; ranuras ociosas {d.RanurasOciosas}; {d.Motivo}.{textoPendiente}";
                    if (d.ReclasificacionesAutorizadas != null && d.ReclasificacionesAutorizadas.Count > 0)
                    {
                        linea += $"\n   [Reclasificación autorizada: {string.Join(", ", d.ReclasificacionesAutorizadas.Select(r => $"{r.NombreAsignatura} ({r.DiasPatronAnterior} -> AUTOMATICO)"))}]";
                    }
                    lineasDetalle.Add(linea);
                }

                string detalleEsc = string.Join("\n", lineasDetalle);
                string sugerenciasResumen = string.Join(", ", decisionesAplicables
                    .Select(d =>
                    {
                        var aulaObj = _contextoActual?.Aulas?.FirstOrDefault(a => string.Equals(a.AulaRef, d.AulaRef, StringComparison.OrdinalIgnoreCase));
                        var progObj = _contextoActual?.Programas?.FirstOrDefault(p => p.Id == aulaObj?.IdPrograma);
                        var nivelObj = AgNivelAcademicoCatalogo.Desde(progObj?.Modalidad);
                        double unidad = aulaObj != null ? AgUnidadesSemanales.HorasSesionAula(aulaObj, nivelObj) : 4.0;
                        return $"{d.Desde} → {d.Hacia} ({AgEscaleraVentanas.Rotular(d.Hacia, unidad)})";
                    })
                    .Distinct());

                bool algunaPendiente = decisionesAplicables.Any(d => d.DemandaPendiente > 0);
                string advertenciaPendiente = algunaPendiente
                    ? "\n\nNOTA: Para una o más aulas, este escalado es un paso intermedio que amplía la ventana pero aún dejará demanda pendiente insatisfecha (se requiere revisar restricciones; otro escalón puede no resolverlas). Las fechas de semestre se mantienen fijas."
                    : "\n\nNOTA: Las fechas de inicio y fin de semestre se mantienen fijas.";

                string advertenciaReset = totalAsigsReset > 0
                    ? $"\n\nAVISO DE PATRONES: Se recalcularán los patrones AUTOMÁTICOS de {totalAsigsReset} asignatura(s) en las aulas afectadas para distribuirlas en los nuevos días."
                    : string.Empty;

                MarcarDecisionPreparacion();
                var respuesta = MessageBox.Show(this,
                    "Se detectaron aulas cuya demanda excede la capacidad o saturación de su ventana actual:\n\n" + detalleEsc +
                    advertenciaPendiente +
                    advertenciaReset +
                    "\n\n¿Desea confirmar y aplicar la propuesta definitiva (" + sugerenciasResumen + ")?",
                    "Escalado de ventana (Pre-Vuelo)", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                traza?.Dato("EscaladoSugerido", string.Join(" | ", decisionesAplicables.Select(d => $"{d.AulaRef}:{d.Desde}->{d.Hacia}")));

                if (respuesta != DialogResult.Yes)
                {
                    Log("[ESCALADO] Propuesta de escalado rechazada por el usuario.");
                    return AgResultadoEscaladoUI.RechazadoPorUsuario;
                }

                ReanudarValidacionPreparacion();
                // Paso 5 - Guardado del mismo lote sellado sin recalcular
                var vigente = await ConstruirContextoAsync(ct);
                if (vigente.IsFailure)
                {
                    MarcarErrorPreparacion("ESCALADO_LECTURA_VIGENTE");
                    Log(vigente.Error);
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                }
                int perfilVigente = cmbArboles.SelectedItem is AgArbol perfilActual ? (int)perfilActual.Id : 0;
                if (perfilVigente != idArbolActivo)
                {
                    MarcarErrorPreparacion("ESCALADO_PERFIL_OBSOLETO");
                    MessageBox.Show(this, "Cambió el perfil. Recalcule la propuesta.", "Escalado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                }

                if (!await PrepararCoberturaAntesDeCambiosAsync(vigente.Value, ct))
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                int esperadas = decisionesAplicables.Count;
                var resultado = await escalador.AplicarDetalladoAsync(decisionesAplicables, idArbolActivo, vigente.Value, ct);
                if (!resultado.Exito || resultado.Confirmadas != esperadas)
                {
                    Log("[ESCALADO] " + resultado.Error);
                    MarcarErrorPreparacion("ESCALADO_APLICACION");
                    MessageBox.Show(this, string.IsNullOrEmpty(resultado.Error) ? "No se confirmó el escalado completo." : resultado.Error,
                        "Escalado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                }
                Log("[ESCALADO] Confirmadas=" + resultado.Confirmadas + "; modificadas=" + resultado.Modificadas);
                if (!string.IsNullOrEmpty(resultado.Advertencia)) { Log(resultado.Advertencia); }
                // La memoria se vuelve a cargar desde BD; nunca simular cambios de relaciones no confirmados.
                var recarga = await ConstruirContextoAsync(ct);
                if (recarga.IsFailure)
                {
                    MarcarErrorPreparacion("ESCALADO_RECARGA_CONTEXTO");
                    MessageBox.Show(this, "El escalado quedó guardado, pero no se pudo recargar el contexto: " + recarga.Error,
                        "Escalado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return AgResultadoEscaladoUI.FalloAlAplicar;
                }
                _contextoActual = recarga.Value;
                if (!await PrepararCoberturaAntesDeCambiosAsync(_contextoActual, ct))
                    return AgResultadoEscaladoUI.FalloAlAplicar;

                // Recargar dimensiones del perfil activo en el árbol
                if (idArbolActivo > 0)
                {
                    if (!await RecargarDimensionesPerfilActivoTrasEscaladoAsync(idArbolActivo, ct).ConfigureAwait(true))
                    {
                        MarcarErrorPreparacion("ESCALADO_RECARGA_DIMENSIONES");
                        MessageBox.Show(this,
                            "No se pudieron actualizar las dimensiones. Recargue el perfil.",
                            "Escalado de ventana", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return AgResultadoEscaladoUI.FalloAlAplicar;
                    }
                }

                // Re-validar fuera del hilo UI
                var reval = await Task.Run(() => _validadorVentana.ValidarFactibilidad(_contextoActual), ct).ConfigureAwait(true);
                if (reval.IsFailure || (reval.Value != null && !reval.Value.EsFactible))
                {
                    string msgReval = reval.IsFailure
                        ? reval.Error
                        : string.Join(" | ", reval.Value.BloqueosCriticos);
                    Log($"[AVISO RE-VALIDACIÓN] {msgReval}");
                    if (!algunaPendiente)
                    {
                        // Aviso no terminal: el flujo heredado continúa y el motor conserva sus guardas.
                        ReanudarValidacionPreparacion();
                        MessageBox.Show(this, msgReval, "Re-validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                return AgResultadoEscaladoUI.AplicadoExitosamente;
            }
            catch (ArgumentException ex)
            {
                MarcarErrorPreparacion("ESCALADO_ARGUMENTO_INVALIDO");
                MessageBox.Show(this,
                    $"Error al procesar el escalado o la reclasificación: {ex.Message}\n\nDebe recalcular la propuesta para continuar.",
                    "Reclasificación no aplicable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                Log($"[ESCALADO-WARN] Excepción controlada de validación de reclasificación: {ex.Message}");
                return AgResultadoEscaladoUI.FalloAlAplicar;
            }
        }

        private async Task IniciarOptimizacionAsync()
        {
            // Fase 1727: No permitir optimización si el diagnóstico está en curso
            if (_diagnosticoEnCurso || _intentoOptimizacion.HasValue)
            {
                return;
            }

            // Identidad local anterior a cualquier await; no autoriza revisar horarios aprobados.
            var intento = Guid.NewGuid();
            _intentoOptimizacion = intento;
            _operacionEnCurso = true; // Fase 1905: protección de ciclo de vida completo de la optimización y post-proceso
            _aceptaProgresoOptimizacion = false;
            _terminalVisualOptimizacion = false;
            _faseFinOptimizacion = null;
            _motivoFinOptimizacion = null;

            // Fase 1716, Fase 1727 & Fase 1739: Bloquear controles y acciones de perfil/árbol, detener temporizador y adquirir semáforo
            bool arbolHabilitado = treeContextoAcademico.Enabled;
            bool perfilesHabilitados = cmbArboles.Enabled;
            bool btnValidarHabilitado = btnValidarFactibilidad.Enabled;
            bool btnEjecutarHabilitado = btnEjecutarAG.Enabled;
            bool btnFiltrarHabilitado = btnFiltrarArbol.Enabled;
            bool btnCargarHabilitado = btnCargarArbol.Enabled;
            bool btnSobrescribirHabilitado = btnSobrescribirArbol.Enabled;

            treeContextoAcademico.Enabled = false;
            cmbArboles.Enabled = false;
            btnValidarFactibilidad.Enabled = false;
            btnEjecutarAG.Enabled = false;
            btnFiltrarArbol.Enabled = false;
            btnCargarArbol.Enabled = false;
            btnSobrescribirArbol.Enabled = false;
            btnCancelar.Enabled = true;
            progressBarAG.Value = 0;
            EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Preparando);

            _timerSyncDimensiones.Stop();
            _ctsEjecucion = new CancellationTokenSource();
            LoggerHelper.ReiniciarDedupe(); // Fase 594/598: Reinicio centralizado de dedupe por proceso
            AgCoherenciaGrupo.ReiniciarIncidencias(); // Fase 852: Limpiar incidencias entre corridas
            _ultimoResultado = null; // Fase 850: Reiniciar estado para evitar retener corrida previa tras fallos
            _idCorridaCargada = 0;   // Fase 850: Evita apuntar a corrida previa si falla la optimización
            _contextoActual = null;                  // Fase 856: Limpiar contexto previo para evitar resolución con catálogos obsoletos
            _aulasConfirmadasEnIntento.Clear();
            btnPromoverProduccion.Enabled = false;   // Fase 856: Deshabilitar promoción hasta que la nueva corrida sea exitosa
            var traza = CrearTraza("EjecucionCorridaAG"); // Fase 606 (antes: "OptimizacionAG")
            traza.Inicio();
            RegistrarEventoIntento("INICIO", intento, 0);
            traza.Paso("P08_IniciarOptimizacion", "Preparación de corrida y ejecución evolutiva del Algoritmo Genético");

            long idCorridaCreada = 0;
            AgCorrida corrida = null;
            bool corridaFinalizada = false;
            bool cancelacionDetectada = false;
            bool semaforoAdquirido = false;

            try
            {
                await _semaforoSyncDimensiones.WaitAsync(_ctsEjecucion.Token);
                semaforoAdquirido = true;
                _ctsEjecucion.Token.ThrowIfCancellationRequested();
                txtLogAG.Clear();
                Log("================================================================================");
                Log(" INICIANDO OPTIMIZACIÓN EVOLUTIVA (Algoritmo Genético)");
                Log("================================================================================");

                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Validando);
                var ctxRes = await ConstruirContextoAsync(_ctsEjecucion.Token);
                if (ctxRes.IsFailure)
                {
                    traza.Fallo("P08_ValidarContexto", new Exception(ctxRes.Error), AgCausaFallo.SinAsignaturas);
                    traza.Fin(false, ctxRes.Error);
                    Log($"[ERROR] {ctxRes.Error}");
                    MarcarErrorPreparacion("P08_ValidarContexto");
                    MessageBox.Show(this, ctxRes.Error, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _contextoActual = ctxRes.Value;
                if (!await PrepararCoberturaAntesDeCambiosAsync(_contextoActual, _ctsEjecucion.Token)) return;
                if (_dimensionesPendientes != null)
                {
                    var pend = _dimensionesPendientes.Value;
                    var resSync = await _aulaDimSync.SincronizarConfirmadoAsync(pend.AulaRef, pend.Jornada, pend.HorasDia, pend.DiasVentana, pend.Subunidad, _ctsEjecucion.Token);
                    if (resSync.IsFailure)
                    {
                        Log($"[ERROR PRE-VUELO] {resSync.Error}");
                        MarcarErrorPreparacion("P08_SincronizarDimensiones");
                        MessageBox.Show(this, resSync.Error, "Sincronización de dimensiones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _dimensionesPendientes = null;
                    var recarga = await ConstruirContextoAsync(_ctsEjecucion.Token);
                    if (recarga.IsFailure) { MostrarFalloPlan(TipoPreparacionPlan.ErrorTecnico, recarga.Error); return; }
                    _contextoActual = recarga.Value;
                    if (!await PrepararCoberturaAntesDeCambiosAsync(_contextoActual, _ctsEjecucion.Token)) return;
                }

                var combinaciones = ObtenerCombinacionesSeleccionadas();
                Log($"[ALCANCE] Optimización: {combinaciones.Count} combinaciones confirmadas.");

                // Fase 213 & 231: Pre-vuelo defensivo antes de ejecutar AG
                if (_contextoActual.Aulas == null || _contextoActual.Aulas.Count == 0)
                {
                    string msg = "No existe Aula (ag_aula) para la combinación seleccionada. Use 'Generar Aulas (N)'.";
                    MarcarErrorPreparacion("P08_ValidarAulas");
                    traza.Fallo("P08_ValidarAulas", null, AgCausaFallo.SinAulas);
                    traza.Fin(false, msg);
                    Log($"[ERROR] {msg}");
                    MessageBox.Show(msg, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fases 454, 472, 480 & Fase 645: Validación de anclaje multi-combinación con contexto de alcance y fallback de cohorte a homónimos
                if (combinaciones.Count > 1)
                {
                    var faltantes = combinaciones
                        .Where(c => !_contextoActual.Aulas.Any(a =>
                            string.Equals(a.AulaRef, c.AulaRef, StringComparison.OrdinalIgnoreCase) ||
                            (AgAulaFactory.ExtraerCohorte(c.AulaRef).Length > 0 &&
                             string.Equals(AgAulaFactory.ExtraerCohorte(a.AulaRef),
                                           AgAulaFactory.ExtraerCohorte(c.AulaRef), StringComparison.OrdinalIgnoreCase))))
                        .Select(c => $"{c.AulaRef} [{ModalidadTutoriaCatalog.AEtiqueta(c.Modalidad)}]").ToList();
                    if (faltantes.Count > 0)
                    {
                        var semestresCombinaciones = combinaciones
                            .Select(c => AgFormatoAsignatura.ParsearSemestreGrupo(c.Semestre).Semestre)
                            .Where(s => !string.IsNullOrWhiteSpace(s) && s != "SIN ASIGNAR")
                            .Distinct().ToList();
                        var err = $"Faltan aulas (ag_aula) para {faltantes.Count} combinación(es):\n• {string.Join("\n• ", faltantes)}\n" +
                                  $"Alcance cargado: {_contextoActual.Aulas.Count} aulas, semestres {string.Join(",", semestresCombinaciones)}.\n" +
                                  "Ejecute 'Generar Aulas (N)' para el período objetivo.";
                        MarcarErrorPreparacion("P08_ValidarAnclajeAula");
                        traza.Fallo("P08_ValidarAnclajeAula", new Exception(err), AgCausaFallo.SinAnclajeAula);
                        traza.Fin(false, err);
                        Log($"[ERROR] {err}");
                        MessageBox.Show(this, err, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    traza.Dato("AnclajeMultiCombinacion", $"aulas cubiertas={combinaciones.Count}, IdAulaCorrida=0 (multi-grupo)");
                }
                else if (_contextoActual.IdAulaCorrida <= 0)
                {
                    string refEsp = combinaciones.Any() ? combinaciones.First().AulaRef : "seleccionada";
                    string msg = $"No existe Aula (ag_aula) exacta para la combinación '{refEsp}'. Use 'Generar Aulas (N)' antes de optimizar.";
                    MarcarErrorPreparacion("P08_ValidarAnclajeAula");
                    traza.Fallo("P08_ValidarAnclajeAula", null, AgCausaFallo.SinAnclajeAula);
                    traza.Fin(false, msg);
                    Log($"[ERROR] {msg}");
                    MessageBox.Show(this, msg, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 275: Diagnóstico pre-vuelo accionable con idprograma, semestre y subunidad
                if (_contextoActual.Asignaturas == null || _contextoActual.Asignaturas.Count == 0)
                {
                    var comb = combinaciones.FirstOrDefault();
                    string semestreRama = comb?.Semestre ?? "";
                    string detalle = $"idprograma={_contextoActual.IdPrograma} semestre='{semestreRama}' subunidad='{_contextoActual.Subunidad}'";
                    string msg = $"Sin asignaturas para {detalle}. Verifique el catálogo 'asignaturas' (idprograma/semestre/subunidad) o cargue el pensum.";
                    MarcarErrorPreparacion("P08_ValidarAsignaturas");
                    traza.Fallo("P08_ValidarAsignaturas", null, AgCausaFallo.SinAsignaturas);
                    traza.Fin(false, msg);
                    Log($"[ERROR] {msg}");
                    MessageBox.Show(msg, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 293, 479 & Fase 641: Validar asignaturas vinculadas en las aulas del alcance solicitado
                var refsSolicitadasOpt = new HashSet<string>(
                    combinaciones.Select(c => c.AulaRef), StringComparer.OrdinalIgnoreCase);
                var aulasGateOpt = _contextoActual.Aulas != null
                    ? _contextoActual.Aulas.Where(a => refsSolicitadasOpt.Count == 0 || refsSolicitadasOpt.Contains(a.AulaRef)).ToList()
                    : new List<AgAula>();
                var aulasSinMateriasOpt = aulasGateOpt
                    .Where(a => _contextoActual.AulaAsignaturas == null || !_contextoActual.AulaAsignaturas.Any(x => x.IdAula == a.Id))
                    .Select(a => a.AulaRef).ToList();
                var huerfanasExcluidasOpt = _contextoActual.Aulas != null
                    ? _contextoActual.Aulas.Where(a => !refsSolicitadasOpt.Contains(a.AulaRef)
                        && (_contextoActual.AulaAsignaturas == null || !_contextoActual.AulaAsignaturas.Any(x => x.IdAula == a.Id)))
                        .Select(a => a.AulaRef).ToList()
                    : new List<string>();
                if (huerfanasExcluidasOpt.Count > 0)
                    traza.Advertencia("AulasHuerfanasNoSolicitadas",
                        $"{huerfanasExcluidasOpt.Count} aula(s) huérfanas fuera de la selección omitidas: {string.Join(" ; ", huerfanasExcluidasOpt)}");

                if (aulasSinMateriasOpt.Count > 0 || _contextoActual.AulaAsignaturas == null || _contextoActual.AulaAsignaturas.Count == 0)
                {
                    var msg = $"Sin ag_aula_asignatura para {aulasSinMateriasOpt.Count} aula(s):\n• {string.Join("\n• ", aulasSinMateriasOpt)}\nEjecute 'Generar Aulas (N)'.";
                    MarcarErrorPreparacion("P08_ValidarAulaAsignaturas");
                    traza.Fallo("P08_ValidarAulaAsignaturas", new Exception(msg), AgCausaFallo.SinAulaAsignaturas);
                    traza.Fin(false, msg);
                    Log($"[ERROR] {msg}");
                    MessageBox.Show(this, msg, "Validación Pre-Vuelo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // F655 / Fase 1113 / Fase 1114 / Fase 1116 / Fase 1192: Validación de ventanas y bloques de decisión acumulando aulas afectadas
                if (_contextoActual.Aulas != null && _contextoActual.Bloques != null)
                {
                    var aulasSinDim = new List<string>();
                    var aulasConError = new List<string>();
                    var msgsDetallados = new List<string>();

                    foreach (var aula in aulasGateOpt)
                    {
                        var progV = _contextoActual.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma);
                        var nivelV = AgNivelAcademicoCatalogo.Desde(progV?.Modalidad);
                        int horasEfectivasV = (int)AgUnidadesSemanales.HorasSesionAula(aula, nivelV);

                        // Fase 1114: Bloqueo explícito en pre-vuelo si aula de pregrado tiene configuradas 8 h/día
                        if (aula.HorasDia == 8 && nivelV == AgNivelAcademico.Pregrado)
                        {
                            var errD4 = $"Aula '{aula.AulaRef}': 8 h/día no permitido en PREGRADO (unidad canónica 4 h). Seleccione 4h en el árbol y ejecute 'Generar Aulas'.";
                            // Fase 1120: Causa de fallo específica para incoherencia horas/día ↔ nivel
                            MarcarErrorPreparacion("P08_HorasDiaNivel");
                            traza.Fallo("P08_HorasDiaNivel", new Exception(errD4), AgCausaFallo.HorasDiaInvalidoNivel);
                            traza.Fin(false, errD4);
                            Log($"[ERROR] {errD4}");
                            MessageBox.Show(this, errD4, "Validación Pre-Vuelo Horas/Día", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        var valVentana = AgValidadorVentana.Validar(
                            string.Empty, // Fase 1116: patrón por asignatura se valida en el motor
                            aula.DiasVentana,
                            aula.Jornada,
                            horasEfectivasV,
                            _contextoActual.Bloques);

                        if (valVentana.IsFailure)
                        {
                            var faltantes = AgMensajeDimensionesAula.Faltantes(aula);
                            string txtError = AgMensajeDimensionesAula.Construir(aula, valVentana.Error);
                            msgsDetallados.Add(txtError);

                            if (faltantes.Count > 0)
                            {
                                aulasSinDim.Add(aula.AulaRef);
                            }
                            else
                            {
                                aulasConError.Add(aula.AulaRef);
                            }
                        }
                    }

                    if (msgsDetallados.Count > 0)
                    {
                        const int MaxDetalle = 5;
                        string msgModal = string.Join("\n\n", msgsDetallados.Take(MaxDetalle));
                        if (msgsDetallados.Count > MaxDetalle)
                        {
                            msgModal += $"\n\n... y {msgsDetallados.Count - MaxDetalle} aula(s) más en la misma condición.";
                        }

                        string msgLog = msgModal.Replace("\r", " ").Replace("\n", " ¶ ");
                        var causa = aulasSinDim.Count > 0 ? AgCausaFallo.VentanaAulaVacia : AgCausaFallo.SinBloques;

                        MarcarErrorPreparacion("P08_ValidarVentanaBloques");
                        traza.Fallo("P08_ValidarVentanaBloques", new Exception(msgLog), causa);
                        traza.Dato("AulasSinDimensiones", aulasSinDim.Count);
                        traza.Dato("AulaSinDimensionesDetalle", string.Join(",", aulasSinDim.Take(50)));
                        traza.Dato("AulasConErrorVentana", aulasConError.Count);
                        traza.Fin(false, msgLog);
                        Log($"[ERROR] {msgLog}");
                        MessageBox.Show(this, msgModal, "Validación Pre-Vuelo Ventana/Jornada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Fase 1739: pre-vuelo de escalado de ventana unificado en IniciarOptimizacionAsync
                var resEscaladoOpt = await EvaluarYAplicarEscaladoAsync(traza, _ctsEjecucion.Token);
                if (resEscaladoOpt == AgResultadoEscaladoUI.FalloAlAplicar)
                {
                    if (!_terminalVisualOptimizacion) MarcarErrorPreparacion("ESCALADO_NO_APLICABLE");
                    return;
                }
                if (resEscaladoOpt == AgResultadoEscaladoUI.RechazadoPorUsuario)
                {
                    MarcarDecisionPreparacion();
                    var cont = MessageBox.Show(this,
                        "Ha rechazado el escalado de ventana sugerido. La capacidad actual es insuficiente para la demanda y la optimización resultará con sesiones sin ubicar o violaciones.\n\n¿Desea continuar con la optimización de todos modos?",
                        "Escalado Rechazado", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (cont != DialogResult.Yes)
                    {
                        Log("[OPTIMIZACIÓN] Cancelada por el usuario tras rechazar el escalado.");
                        cancelacionDetectada = true;
                        EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Cancelado, "ESCALADO_CONTINUACION_RECHAZADA");
                        return;
                    }
                }

                ReanudarValidacionPreparacion();
                // Fase 295, 296, 303 & 321: Validar cobertura de pool docente (Consentimiento de Vacante)
                int totalAsigOpt = _contextoActual.Asignaturas.Count;
                var asignaturasSinPoolOpt = _contextoActual.Asignaturas
                    .Where(a => !(_contextoActual.PoolPorAsignatura != null &&
                                  _contextoActual.PoolPorAsignatura.TryGetValue(a.Id, out var pl) &&
                                  pl.Any(p => p.Activo == 1)))
                    .ToList();
                int asigSinPoolOpt = asignaturasSinPoolOpt.Count;
                int asigConPoolOpt = totalAsigOpt - asigSinPoolOpt;

                traza.Dato("CoberturaDocente", $"{asigConPoolOpt}/{totalAsigOpt}");

                if (asigSinPoolOpt > 0)
                {
                    string listaSinOpt = AgFormatoAsignatura.BulletList(asignaturasSinPoolOpt);
                    traza.Dato("AsignaturasSinDocente", listaSinOpt.Replace("\n", " "));
                    traza.Causa(AgCausaFallo.SinPoolDocente, $"{asigSinPoolOpt}/{totalAsigOpt} → VACANTE");
                    Log($"[OPTIMIZACIÓN] {asigSinPoolOpt}/{totalAsigOpt} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSinOpt}");

                    MarcarDecisionPreparacion();
                    var dialogRes = MessageBox.Show(this,
                        $"Hay {asigSinPoolOpt} de {totalAsigOpt} asignaturas sin docente habilitado en el pool:\n• {listaSinOpt}\n\n" +
                        "Se marcarán como VACANTE (es_vacante=1) con penalización suave y podrá resolverlas después en el Panel de Vacantes.\n\n" +
                        "¿Continuar con la optimización?",
                        "Validación Pre-Vuelo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (dialogRes != DialogResult.Yes)
                    {
                        traza.Fin(false, "Cancelado por cobertura incompleta");
                        Log($"[OPTIMIZACIÓN] Cancelado por el usuario debido a {asigSinPoolOpt}/{totalAsigOpt} asignaturas sin pool.");
                        cancelacionDetectada = true;
                        _contextoActual?.AsignaturasVacantesConsentidas?.Clear();
                        EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Cancelado, "POOL_CONTINUACION_RECHAZADA");
                        return;
                    }

                    ReanudarValidacionPreparacion();
                    _contextoActual.AsignaturasVacantesConsentidas = new HashSet<int>(asignaturasSinPoolOpt.Select(a => a.Id));
                    Log($"[OPTIMIZACIÓN] Aceptado por el usuario: {asigSinPoolOpt}/{totalAsigOpt} asignaturas pasarán a VACANTE.");
                }
                else
                {
                    _contextoActual?.AsignaturasVacantesConsentidas?.Clear();
                }

                // Captura final tras cualquier recarga de escalado; ninguna corrida pierde su manifiesto.
                if (_contextoActual.ManifiestoCobertura == null || _contextoActual.ValidacionPlanAprobado == null)
                    throw new InvalidOperationException("PREPARACION_COBERTURA_AUSENTE");
                // T1: una nueva corrida no purga staging, histórico ni auditoría de corridas anteriores.
                _ctsEjecucion.Token.ThrowIfCancellationRequested();

                // 1. Crear registro de corrida (Fase 219: anclado al aula de la corrida)
                // Corrida 92 - T6: persistir configuración reproducible en ParametrosJson (ventana, snapshot con demanda, patrones, bloques y festivos, versión, huella)
                var aulaCorrida = _contextoActual.Aulas?.FirstOrDefault(a => a.Id == _contextoActual.IdAulaCorrida)
                                  ?? _contextoActual.Aulas?.FirstOrDefault();
                string huellaContexto = null;
                string errorCaptura = null;
                try
                {
                    if (aulaCorrida == null)
                    {
                        throw new InvalidOperationException("No se encontró un aula configurada para la corrida.");
                    }
                    huellaContexto = AgHuellaPlanVentana.Calcular(_contextoActual, aulaCorrida);
                    if (string.IsNullOrWhiteSpace(huellaContexto))
                    {
                        throw new InvalidOperationException("Huella calculada vacía para el plan de ventana.");
                    }
                }
                catch (Exception ex)
                {
                    errorCaptura = ex.Message;
                    LoggerHelper.Error($"[AgFormGeneradorHorarios] Error al calcular huella de plan de ventana: {ex.Message}", ex);
                    throw new InvalidOperationException("SNAPSHOT_HUELLA_ERROR", ex);
                }

                var ventanaEfectiva = _contextoActual.Ventana ?? _ventanaResuelta;

                // Contenido recuperable: configuración completa de todas las aulas (incluyendo origen, estado y configuración temporal)
                var aulasLista = (_contextoActual.Aulas ?? new List<AgAula>()).Select(au => new Dictionary<string, object>
                {
                    { "id", au.Id },
                    { "idSede", au.IdSede },
                    { "idPeriodo", au.IdPeriodo },
                    { "idPrograma", au.IdPrograma },
                    { "codigo", au.Codigo },
                    { "nombre", au.Nombre },
                    { "capacidad", au.Capacidad },
                    { "tipoAula", au.TipoAula },
                    { "esVirtual", au.EsVirtual },
                    { "estado", (int)au.Estado },
                    { "caracteristicasJson", au.CaracteristicasJson },
                    { "semestre", au.Semestre },
                    { "grupo", au.Grupo },
                    { "modalidadVinculacion", au.ModalidadVinculacion },
                    { "fechaInicio", au.FechaInicio },
                    { "fechaFin", au.FechaFin },
                    { "semestreInicio", au.SemestreInicio },
                    { "semestreFin", au.SemestreFin },
                    { "ventanaOrigen", au.VentanaOrigen },
                    { "aulaRef", au.AulaRef },
                    { "subunidad", au.Subunidad },
                    { "jornada", au.Jornada },
                    { "horasDia", au.HorasDia },
                    { "diasVentana", au.DiasVentana }
                }).ToList();

                // Contenido recuperable: demanda efectiva de asignaturas (con HorasObjetivo y OrigenPatron calculados como el evaluador)
                var demandaLista = (_contextoActual.Asignaturas ?? new List<Asignatura>()).Select(a =>
                {
                    AgAulaAsignatura aa = null;
                    if (aulaCorrida != null && _contextoActual.AulaAsignaturaPorAulaAsig != null)
                    {
                        _contextoActual.AulaAsignaturaPorAulaAsig.TryGetValue((aulaCorrida.Id, a.Id), out aa);
                    }
                    if (aa == null && _contextoActual.AulaAsignaturasPorId != null)
                    {
                        _contextoActual.AulaAsignaturasPorId.TryGetValue(a.Id, out aa);
                    }

                    double horasObj = AgEvaluadorDemandaVentana.HorasObjetivo(a, aa);
                    string origPatron = aa?.OrigenPatron ?? "AUTOMATICO";
                    string diasPat = aa?.DiasPatron ?? string.Empty;

                    return new Dictionary<string, object>
                    {
                        { "idAsignatura", a.Id },
                        { "codigo", a.Codigo },
                        { "nombre", a.Nombre },
                        { "horas", a.Horas },
                        { "horasObjetivo", horasObj },
                        { "horasSemanales", a.HorasSemanales },
                        { "ordenCurricular", aa?.OrdenCurricular ?? 0 },
                        { "diasPatron", diasPat },
                        { "origenPatron", origPatron }
                    };
                }).ToList();

                // Contenido recuperable: relaciones por (IdAula, IdAsignatura) con HorasObjetivo evaluadas
                var relacionesAulaAsigLista = (_contextoActual.AulaAsignaturas ?? new List<AgAulaAsignatura>()).Select(aa =>
                {
                    var asigRel = _contextoActual.Asignaturas?.FirstOrDefault(a => a.Id == aa.IdAsignatura);
                    double horasObjRel = asigRel != null
                        ? AgEvaluadorDemandaVentana.HorasObjetivo(asigRel, aa)
                        : (aa.HorasObjetivo > 0 ? aa.HorasObjetivo : 64.0);

                    return new Dictionary<string, object>
                    {
                        { "idAula", aa.IdAula },
                        { "idAsignatura", aa.IdAsignatura },
                        { "horasObjetivo", aa.HorasObjetivo },
                        { "horasEfectivas", horasObjRel },
                        { "ordenCurricular", aa.OrdenCurricular },
                        { "diasPatron", aa.DiasPatron },
                        { "origenPatron", aa.OrigenPatron },
                        { "prioridad", aa.Prioridad },
                        { "exclusivo", aa.Exclusivo },
                        { "requerida", aa.Requerida },
                        { "subunidad", aa.Subunidad }
                    };
                }).ToList();

                // Contenido recuperable: patrones por asignatura
                var patronesDict = new Dictionary<string, string>();
                if (_contextoActual.Asignaturas != null)
                {
                    foreach (var a in _contextoActual.Asignaturas)
                    {
                        AgAulaAsignatura aa = null;
                        if (aulaCorrida != null && _contextoActual.AulaAsignaturaPorAulaAsig != null)
                        {
                            _contextoActual.AulaAsignaturaPorAulaAsig.TryGetValue((aulaCorrida.Id, a.Id), out aa);
                        }
                        if (aa == null && _contextoActual.AulaAsignaturasPorId != null)
                        {
                            _contextoActual.AulaAsignaturasPorId.TryGetValue(a.Id, out aa);
                        }
                        string pat = aa?.DiasPatron ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(pat))
                        {
                            patronesDict[a.Id.ToString()] = pat;
                        }
                    }
                }

                // Contenido recuperable: bloques horarios activos
                var bloquesLista = (_contextoActual.Bloques ?? new List<AgBloqueHorario>())
                    .Where(b => b.Activo == 1)
                    .Select(b => new Dictionary<string, object>
                    {
                        { "id", b.Id },
                        { "dia", b.DiaSemanaNumero },
                        { "franja", b.Franja },
                        { "horaInicio", b.HoraInicio },
                        { "horaFin", b.HoraFin },
                        { "duracionHoras", b.DuracionHoras },
                        { "activo", b.Activo },
                        { "subunidad", b.Subunidad }
                    }).ToList();

                // Contenido recuperable: festivos en la ventana
                var festivosLista = (_contextoActual.Festivos ?? new List<(DateTime Inicio, DateTime Fin)>())
                    .Select(f => new Dictionary<string, string>
                    {
                        { "inicio", f.Inicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) },
                        { "fin", f.Fin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) }
                    }).ToList();

                // Contenido recuperable: programas canónicos para snapshot sellado idéntico a la huella
                var programasLista = (_contextoActual.Programas ?? new List<Programa>()).Select(p => new Dictionary<string, object>
                {
                    { "id", p.Id },
                    { "nombrePrograma", p.NombrePrograma },
                    { "abreviatura", p.Abreviatura },
                    { "pensum", p.Pensum },
                    { "actual", p.Actual },
                    { "activo", p.Activo },
                    { "modalidad", p.Modalidad }
                }).ToList();

                // Contenido recuperable: asignaturas canónicas para snapshot sellado idéntico a la huella
                var asignaturasLista = (_contextoActual.Asignaturas ?? new List<Asignatura>()).Select(a => new Dictionary<string, object>
                {
                    { "id", a.Id },
                    { "idPrograma", a.IdPrograma },
                    { "semestre", a.Semestre },
                    { "numNivel", a.NumNivel },
                    { "codUnidad", a.CodUnidad },
                    { "codMateria", a.CodMateria },
                    { "uniTeorica", a.UniTeorica },
                    { "metodologia", a.Metodologia },
                    { "abreviatura", a.Abreviatura },
                    { "nucleo", a.Nucleo },
                    { "nombre", a.Nombre },
                    { "creditos", a.Creditos },
                    { "intHoraria", a.IntHoraria },
                    { "tipoAsignatura", a.TipoAsignatura },
                    { "tipoCredito", a.TipoCredito },
                    { "prerequisito", a.Prerequisito },
                    { "horas", a.Horas },
                    { "horasSemanales", a.HorasSemanales },
                    { "idSede", a.IdSede },
                    { "modalidad", a.Modalidad },
                    { "subunidad", a.Subunidad }
                }).ToList();

                var parametrosDict = new Dictionary<string, object>
                {
                    { "intentoId", intento.ToString("D") },
                    { "version", App.APP_VERSION },
                    { "vacantesConsentidas", _contextoActual.AsignaturasVacantesConsentidas != null ? _contextoActual.AsignaturasVacantesConsentidas.OrderBy(id => id).ToList() : new List<int>() },
                    { "manifiestoCobertura", _contextoActual.ManifiestoCobertura },
                    { "planCanonico", _contextoActual.PlanAprobadoSnapshot },
                    { "validacionPlan", _contextoActual.ValidacionPlanAprobado },
                    { "decisionesCobertura", _contextoActual.CandidatosCobertura },
                    { "sesionesConservadas", _contextoActual.SesionesConservadas },
                    { "huella", huellaContexto ?? string.Empty },
                    { "errorCaptura", errorCaptura },
                    { "ventana", new Dictionary<string, object>
                        {
                            { "idPeriodo", _contextoActual.IdPeriodo },
                            { "periodoNombre", ventanaEfectiva?.PeriodoNombre },
                            { "fechaInicio", ventanaEfectiva?.FechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) },
                            { "fechaFin", ventanaEfectiva?.FechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) },
                            { "origen", ventanaEfectiva != null ? ventanaEfectiva.Origen.ToString() : null },
                            { "observaciones", ventanaEfectiva?.Observaciones },
                            { "duracionDias", ventanaEfectiva?.DuracionDias ?? 0 },
                            { "diasVentana", combinaciones.Count == 1 ? aulaCorrida?.DiasVentana : null },
                            { "jornada", combinaciones.Count == 1 ? aulaCorrida?.Jornada : null }
                        }
                    },
                    { "snapshot", new Dictionary<string, object>
                        {
                            { "idPeriodo", _contextoActual.IdPeriodo },
                            { "idPrograma", _contextoActual.IdPrograma },
                            { "idAula", combinaciones.Count == 1 ? _contextoActual.IdAulaCorrida : 0 },
                            { "aulaRef", combinaciones.Count == 1 ? aulaCorrida?.AulaRef ?? string.Empty : string.Empty },
                            { "aulasSeleccionadas", _contextoActual.ManifiestoCobertura.Relaciones.Select(r => r.Clave.IdAula).Distinct().OrderBy(id => id).ToList() },
                            { "subunidad", _contextoActual.Subunidad ?? string.Empty },
                            { "totalAulas", _contextoActual.Aulas?.Count ?? 0 },
                            { "totalAsignaturas", _contextoActual.Asignaturas?.Count ?? 0 },
                            { "totalBloques", _contextoActual.Bloques?.Count ?? 0 },
                            { "aulas", aulasLista },
                            { "demanda", demandaLista },
                            { "relacionesAulaAsignatura", relacionesAulaAsigLista },
                            { "patrones", patronesDict },
                            { "bloques", bloquesLista },
                            { "festivos", festivosLista },
                            { "programas", programasLista },
                            { "asignaturas", asignaturasLista }
                        }
                    },
                    { "algoritmo", new Dictionary<string, object>
                        {
                            { "poblacion", (int)numPoblacion.Value },
                            { "generacionesMax", (int)numGeneraciones.Value },
                            { "tasaCruce", (double)numTasaCruce.Value },
                            { "tasaMutacion", (double)numTasaMutacion.Value },
                            { "elitismo", (int)numElitismo.Value }
                        }
                    }
                };

                string parametrosJson;
                try
                {
                    var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    serializer.MaxJsonLength = int.MaxValue;
                    parametrosJson = serializer.Serialize(parametrosDict);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[AgFormGeneradorHorarios] Error al serializar ParametrosJson: {ex.Message}", ex);
                    MarcarErrorPreparacion("SNAPSHOT_SERIALIZACION_ERROR");
                    throw new InvalidOperationException("SNAPSHOT_SERIALIZACION_ERROR", ex);
                }

                corrida = new AgCorrida
                {
                    IdPeriodo = _contextoActual.IdPeriodo,
                    IdPrograma = _contextoActual.IdPrograma,
                    IdAula = _contextoActual.IdAulaCorrida,
                    PoblacionTamano = (int)numPoblacion.Value,
                    GeneracionesMax = (int)numGeneraciones.Value,
                    TasaCruce = (double)numTasaCruce.Value,
                    TasaMutacion = (double)numTasaMutacion.Value,
                    Elitismo = (int)numElitismo.Value,
                    Subunidad = _contextoActual.Subunidad,
                    ParametrosJson = parametrosJson,
                    VacantesConsentidas = new HashSet<int>(_contextoActual.AsignaturasVacantesConsentidas ?? Enumerable.Empty<int>())
                };

                var idCorridaResult = await _scheduleRepository.InsertCorridaAsync(corrida, _ctsEjecucion.Token);
                if (idCorridaResult.IsFailure)
                {
                    traza.Fallo("P08_InsertCorridaAsync", new Exception(idCorridaResult.Error));
                    traza.Fin(false, idCorridaResult.Error);
                    Log($"[ERROR] No se pudo crear la corrida: {idCorridaResult.Error}");
                    MarcarErrorPreparacion("P08_InsertCorridaAsync");
                    return;
                }

                long idCorrida = idCorridaResult.Value;
                idCorridaCreada = idCorrida;
                corrida.Id = idCorrida;
                traza.Dato("IdCorrida", idCorrida);
                _contextoActual.IdCorrida = idCorrida; // Fase 585
                traza.Dato("Poblacion", corrida.PoblacionTamano);
                traza.Dato("GeneracionesMax", corrida.GeneracionesMax);

                // Fase 599: Resumen de inicio de corrida
                var periodoSel = cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo);
                var programaSel = cmbProgramas.Primario<Programa>() ?? (cmbProgramas.SelectedItem as Programa);
                var sedeSel = cmbSedes.Primario<Sede>() ?? (cmbSedes.SelectedItem as Sede);

                LoggerHelper.Resumen("InicioCorridaAG",
                    ("Periodo", periodoSel?.Nombre ?? _contextoActual.IdPeriodo.ToString()),
                    ("Programa", programaSel?.Nombre ?? _contextoActual.IdPrograma.ToString()),
                    ("Sede", sedeSel?.Nombre ?? _contextoActual.IdSede.ToString()),
                    ("Cohortes", _contextoActual.AulasAlcance?.Count ?? 0),
                    ("Poblacion", corrida.PoblacionTamano),
                    ("Generaciones", corrida.GeneracionesMax),
                    ("ModoVerboso", LoggerHelper.ModoVerboso));

                Log($"[CORRIDA #{idCorrida}] Población={corrida.PoblacionTamano}, Generaciones={corrida.GeneracionesMax}, Cruce={corrida.TasaCruce:P0}, Mut={corrida.TasaMutacion:P0}, Genes Esperados={_contextoActual.Asignaturas.Count}");

                bool soloConservacion = _contextoActual.ManifiestoCobertura.Relaciones.All(r =>
                    _contextoActual.SesionesConservadas.Any(s => s.Clave.Valor == r.Clave.Valor)) &&
                    _contextoActual.ManifiestoCobertura.Relaciones.All(r => r.EstadoAula == (int)AgEstadoAula.Aprobado);
                if (soloConservacion)
                {
                    var certificado = new AgValidadorCoberturaSeleccion().Validar(_contextoActual.ManifiestoCobertura,
                        _contextoActual.SesionesConservadas, Array.Empty<AgSesionCobertura>(), _contextoActual.IdTerceroVacante, _contextoActual.Festivos, _contextoActual.AsignaturasVacantesConsentidas);
                    if (certificado.IsFailure)
                    {
                        // Fase 1924 / Fase 1929: “Sólo conservación” sin excepción no controlada ante vacantes
                        if (certificado.Error == "SESION_SIN_DOCENTE")
                        {
                            MostrarFalloPlan(TipoPreparacionPlan.NoVerificable, certificado.Error);
                            return;
                        }
                        throw new InvalidOperationException(certificado.Error);
                    }
                    // No se crea población, individuo ni staging; la producción anterior no se toca.
                    corrida.FechaFinCorrida = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    corrida.Estado = AgEstadoCorrida.Completada;
                    var cierreConservado = await _scheduleRepository.UpdateCorridaAsync(corrida, _ctsEjecucion.Token);
                    if (cierreConservado.IsFailure) throw new InvalidOperationException(cierreConservado.Error);
                    corridaFinalizada = true;
                    _idCorridaCargada = idCorrida;
                    _ultimoResultado = new AgResultadoCorrida
                    {
                        IdCorrida = idCorrida,
                        Estado = AgEstadoCorrida.Completada,
                        Mensaje = "Sólo conservación: cobertura canónica validada; sin sesiones nuevas."
                    };
                    btnPromoverProduccion.Enabled = false; // No existe reemplazo ni promoción T6.
                    EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Completado, "SOLO_CONSERVACION",
                        "Estado: Sólo conservación; horario aprobado anterior intacto.");
                    await CargarModuloResultadosAsync(idCorrida);
                    btnPromoverProduccion.Enabled = false;
                    btnAprobarCorrida.Enabled = false;
                    tabControlPrincipal.SelectedTab = tabResultados;
                    return;
                }

                var progressReporter = new Progress<(int Generacion, double MejorFitness, int HardViolations, string Mensaje)>(p =>
                {
                    if (!PuedeRecibirProgresoOptimizacion(intento)) return;
                    int pct = Math.Min(100, (int)((p.Generacion / (double)corrida.GeneracionesMax) * 100));
                    progressBarAG.Value = pct;
                    lblStatusAG.Text = p.Mensaje;
                    if (p.Generacion % 10 == 0 || p.HardViolations == 0)
                    {
                        Log($"Gen {p.Generacion,3} | Fitness: {p.MejorFitness:F4} | Hard: {p.HardViolations}");
                    }
                });

                // Fase 598: Sub-etapa de optimización con AgDiagnosticoTraza propio
                var trazaOpt = CrearTraza("OptimizacionAG");
                trazaOpt.Inicio($"Poblacion={corrida.PoblacionTamano}, Generaciones={corrida.GeneracionesMax}, IdCorrida={idCorrida}");

                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Optimizando);

                // Fase 1895 & Fase 1900/1901: Offload del bucle intensivo de CPU del AG fuera del hilo STA de UI
                var contextoOpt = _contextoActual;
                var contexto = contextoOpt; // Fase 1901: instantánea inmutable para todo el tramo post-await
                var tokenOpt = _ctsEjecucion.Token;

                var tareaOpt = Task.Run(
                    () => _generadorHorarios.EjecutarOptimizacionAsync(
                            contextoOpt,
                            corrida.PoblacionTamano,
                            corrida.GeneracionesMax,
                            corrida.TasaCruce,
                            corrida.TasaMutacion,
                            corrida.Elitismo,
                            progressReporter,
                            tokenOpt),
                    tokenOpt);
                _tareaOptimizacion = tareaOpt;

                var optResult = await tareaOpt.ConfigureAwait(true);

                _aceptaProgresoOptimizacion = false;
                if (optResult.IsFailure)
                {
                    bool cancelada = _ctsEjecucion.IsCancellationRequested;
                    CerrarEstadoVisualOptimizacion(cancelada);
                    if (!cancelada) EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Error, "P08_EjecutarOptimizacionAsync");
                    await CerrarCorridaFallidaAsync(corrida, cancelada);
                    corridaFinalizada = true;
                    trazaOpt.Fin(false, optResult.Error);
                    traza.Fallo("P08_EjecutarOptimizacionAsync", new Exception(optResult.Error));
                    traza.Fin(false, optResult.Error);
                    Log($"[ERROR] Optimización fallida: {optResult.Error}");
                    MessageBox.Show(optResult.Error, "Error en Optimización", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                _ultimoResultado = optResult.Value;
                _ultimoResultado.IdCorrida = idCorrida;
                _idCorridaCargada = idCorrida;

                trazaOpt.Conteo("GeneracionesEjecutadas", _ultimoResultado.GeneracionesEjecutadas);
                trazaOpt.Conteo("HardViolations", _ultimoResultado.MejorIndividuo?.HardViolations ?? 0, 0);
                trazaOpt.Conteo("Genes", _ultimoResultado.MejorIndividuo?.Genes?.Count ?? 0);
                trazaOpt.Fin(true, $"Fitness={_ultimoResultado.MejorIndividuo?.Fitness ?? 0.0:F4}");

                // 2. Guardar Individuo Óptimo y Bitácoras
                if (_ultimoResultado.MejorIndividuo != null)
                {
                    _ultimoResultado.MejorIndividuo.IdCorrida = idCorrida;
                    var saveIndResult = await _scheduleRepository.SaveIndividuoAsync(_ultimoResultado.MejorIndividuo, _ctsEjecucion.Token);
                    if (saveIndResult.IsSuccess)
                    {
                        corrida.IdIndividuoOptimo = saveIndResult.Value;

                        // Fase 1564: verificación dura del detalle persistido en ag_violacion. Evita confundir
                        // "sin violaciones" con "detalle no persistido" (visto en las corridas #44 y #48).
                        var violPersistidas = await _scheduleRepository.GetViolacionesByIndividuoAsync(
                            saveIndResult.Value, _ctsEjecucion.Token);
                        int violEnDb = violPersistidas.IsSuccess ? violPersistidas.Value.Count : -1;
                        int violEnMemoria = _ultimoResultado.MejorIndividuo.Violaciones?.Count ?? 0;

                        if (violEnDb < 0)
                        {
                            LoggerHelper.Warning(
                                $"[AG] Corrida #{idCorrida}: no se pudo verificar ag_violacion ({violPersistidas.Error}).");
                        }
                        else if (violEnDb != violEnMemoria)
                        {
                            LoggerHelper.Warning(
                                $"[AG] Corrida #{idCorrida}: violaciones en memoria={violEnMemoria}, persistidas={violEnDb} " +
                                $"(HARD={_ultimoResultado.MejorIndividuo.HardViolations}). Revise ag_violacion.");
                        }
                        else
                        {
                            LoggerHelper.Info(
                                $"[AG] Corrida #{idCorrida}: {violEnDb} violación(es) persistidas en ag_violacion " +
                                $"(HARD={_ultimoResultado.MejorIndividuo.HardViolations}).");
                        }
                    }
                    else
                    {
                        LoggerHelper.Error(
                            $"[AG] Corrida #{idCorrida}: no se pudo guardar el individuo óptimo: {saveIndResult.Error}");
                    }
                }

                foreach (var b in _ultimoResultado.HistorialConvergencia)
                {
                    b.IdCorrida = idCorrida;
                    await _scheduleRepository.SaveBitacoraAsync(b, _ctsEjecucion.Token);
                }

                corrida.FechaFinCorrida = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                corrida.Estado = _ultimoResultado.Estado;
                corrida.GeneracionesEjecutadas = _ultimoResultado.GeneracionesEjecutadas;
                corrida.MejorFitness = _ultimoResultado.MejorIndividuo?.Fitness ?? 0.0;
                corrida.HardViolations = _ultimoResultado.MejorIndividuo?.HardViolations ?? 0;
                corrida.SoftViolations = _ultimoResultado.MejorIndividuo?.SoftViolations ?? 0.0;
                var updCorrida = await _scheduleRepository.UpdateCorridaAsync(corrida, _ctsEjecucion.Token);
                if (updCorrida.IsFailure)
                {
                    // Fase 1564: la cabecera desapareció (purga o eliminación concurrente). Sin ag_corrida no hay
                    // violaciones visibles ni promoción posible: se aborta ANTES de materializar staging.
                    traza.Fallo("P08_UpdateCorridaAsync", new Exception(updCorrida.Error));
                    traza.Fin(false, updCorrida.Error);
                    LoggerHelper.Error($"[AG] Corrida #{idCorrida}: {updCorrida.Error}");
                    CerrarEstadoVisualOptimizacion(_ctsEjecucion.IsCancellationRequested);
                    await CerrarCorridaFallidaAsync(corrida, _ctsEjecucion.IsCancellationRequested);
                    corridaFinalizada = true;
                    EstablecerEstadoOptimizacion(_ctsEjecucion.IsCancellationRequested
                            ? EstadoVisualOptimizacion.Cancelado : EstadoVisualOptimizacion.Error,
                        "P08_UpdateCorridaAsync", "Estado: Corrida sin cabecera en ag_corrida.");
                    MessageBox.Show(this,
                        $"No se pudo consolidar la cabecera de la corrida #{idCorrida}:\n{updCorrida.Error}\n\n" +
                        "La corrida no se materializará en staging para no dejar datos huérfanos.",
                        "Corrida inconsistente", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Materializando);
                // 3. Materializar en Staging (Fase 383 / Fase 404 / Fase 587 / Fase 598)
                bool stagingOk = true;
                string stagingError = null;
                int cantSes = 0;
                int cantVin = 0;
                double horasProyectadas = 0.0;

                if (_ultimoResultado.MejorIndividuo != null)
                {
                    // Fase 598: Sub-etapa de staging con AgDiagnosticoTraza propio
                    var trazaMat = CrearTraza("MaterializacionStaging");
                    trazaMat.Inicio($"IdCorrida={idCorrida}");

                    Log("\nMaterializando solución óptima en tablas de staging...");
                    var matRes = await _materializador.MaterializarEnStagingAsync(idCorrida, _ultimoResultado.MejorIndividuo, contexto, _ctsEjecucion.Token);
                    if (matRes.IsFailure)
                    {
                        stagingOk = false;
                        stagingError = matRes.Error;
                        trazaMat.Fin(false, matRes.Error);
                        traza.Fallo("P08_MaterializarStaging", new Exception(matRes.Error), AgCausaFallo.StagingVacio);
                        Log($"[ERROR] Fallo en staging: {matRes.Error}");
                        CerrarEstadoVisualOptimizacion(_ctsEjecucion.IsCancellationRequested);
                        if (!_ctsEjecucion.IsCancellationRequested)
                            EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Error, "P08_MaterializarStaging");
                        await CerrarCorridaFallidaAsync(corrida, _ctsEjecucion.IsCancellationRequested);
                        corridaFinalizada = true;
                        _ctsEjecucion.Token.ThrowIfCancellationRequested();
                        MessageBox.Show(this, $"La solución se calculó pero no se pudo materializar en staging:\n{matRes.Error}", "Fallo en Staging", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        await _sincronizador.SincronizarIndividuoAsync(corrida, _ultimoResultado.MejorIndividuo, _ctsEjecucion.Token);

                        var ses = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, CancellationToken.None);
                        var vin = await _scheduleRepository.GetVinculacionesStagingAsync(idCorrida, CancellationToken.None);
                        cantSes = ses.IsSuccess ? ses.Value.Count : -1;
                        cantVin = vin.IsSuccess ? vin.Value.Count : -1;
                        trazaMat.Conteo("SesionesStaging", cantSes, 1);
                        trazaMat.Conteo("VinculacionesStaging", cantVin, 1);
                        trazaMat.Fin(true, $"Sesiones={cantSes}, Vinculaciones={cantVin}");

                        traza.Conteo("SesionesStaging", cantSes, 1);
                        traza.Conteo("VinculacionesStaging", cantVin, 1);

                        horasProyectadas = ses.IsSuccess ? ses.Value.Sum(s => s.Horas) : 0;
                        Log($"✔ Staging: {cantVin} vinculaciones, {cantSes} sesiones, {horasProyectadas:0.#} h proyectadas.");
                    }
                }

                var sesiones = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, _ctsEjecucion.Token);
                int sesionesVacantes = sesiones.IsSuccess ? sesiones.Value.Count(s => s.EsVacante == 1) : 0;
                double horasVacantes = sesiones.IsSuccess ? sesiones.Value.Where(s => s.EsVacante == 1).Sum(s => s.Horas) : 0;

                // Fase 1813 & Fase 1815: Cobertura y factibilidad evaluadas por la compuerta centralizada
                var mejor = _ultimoResultado.MejorIndividuo;
                if (mejor != null)
                {
                    corrida.Vacantes = mejor.Vacantes;
                    // Persistir el plan fechado autorizado dentro de ParametrosJson como referencia independiente sellada
                    try
                    {
                        var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        serializer.MaxJsonLength = int.MaxValue;
                        Dictionary<string, object> dict = null;
                        if (!string.IsNullOrWhiteSpace(corrida.ParametrosJson))
                        {
                            dict = serializer.DeserializeObject(corrida.ParametrosJson) as Dictionary<string, object>;
                        }
                        if (dict == null) dict = new Dictionary<string, object>();

                        var planFechado = new List<Dictionary<string, object>>();
                        if (mejor.Genes != null)
                        {
                            foreach (var g in mejor.Genes)
                            {
                                var genDict = new Dictionary<string, object>
                                {
                                    { "idGen", g.Id },
                                    { "idAsignatura", g.IdAsignatura },
                                    { "idAula", g.IdAula },
                                    { "idBloqueHorario", g.IdBloqueHorario },
                                    { "fechas", g.FechasGen != null ? g.FechasGen.Select(d => d.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)).ToList() : new List<string>() }
                                };
                                planFechado.Add(genDict);
                            }
                        }
                        dict["planFechado"] = planFechado;
                        dict["vacantesConsentidas"] = contexto.AsignaturasVacantesConsentidas != null
                            ? contexto.AsignaturasVacantesConsentidas.OrderBy(id => id).ToList()
                            : new List<int>();
                        if (aulaCorrida != null)
                        {
                            dict["huella"] = AgHuellaPlanVentana.Calcular(contexto, aulaCorrida, planFechado: planFechado);
                        }
                        corrida.ParametrosJson = serializer.Serialize(dict);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("SNAPSHOT_PLAN_FECHADO_SERIALIZACION_ERROR", ex);
                    }
                }
                // Fase 1923 / Fase 1928: Diferenciar calidad del horario vs promovibilidad
                bool calidad = stagingOk && mejor != null && mejor.HardViolations == 0 && mejor.Vacantes == 0;
                bool promovible = stagingOk && AgCompuertaPromocion.EsPromovible(corrida, mejor, sesiones, contexto, permitirVacantes: true);
                if (stagingOk)
                {
                    if (sesiones.IsFailure)
                    {
                        stagingOk = false;
                        stagingError = "LECTURA_STAGING_ERROR:" + sesiones.Error;
                        LoggerHelper.Error($"[AG] Corrida #{idCorrida}: error al leer sesiones de staging: {sesiones.Error}");
                    }
                    else
                    {
                        var nuevas = new AgProyectorSolucionCobertura().Nuevas(contexto, mejor, sesiones.Value);
                        if (nuevas.IsFailure)
                        {
                            stagingOk = false;
                            stagingError = "PROYECCION_COBERTURA_ERROR:" + nuevas.Error;
                            LoggerHelper.Error($"[AG] Corrida #{idCorrida}: error en proyección de cobertura: {nuevas.Error}");
                        }
                        else
                        {
                            // Fase 1897/1898/1920/1925/1930: Captura detallada de vacantes de cobertura
                            var detalleVacantes = new List<AgDetalleVacanteCobertura>();
                            var cobertura = new AgValidadorCoberturaSeleccion().Validar(
                                contexto.ManifiestoCobertura,
                                contexto.SesionesConservadas,
                                nuevas.Value,
                                contexto.IdTerceroVacante,
                                contexto.Festivos,
                                contexto.AsignaturasVacantesConsentidas,
                                detalleVacantes);

                            if (cobertura.IsFailure)
                            {
                                stagingOk = false;
                                stagingError = "COBERTURA_NO_CERTIFICADA:" + cobertura.Error;
                                MarcarErrorPreparacion("P08_Cobertura_" + cobertura.Error);
                                var noConsentidas = detalleVacantes.Where(d => !d.Consentida).ToList();
                                LoggerHelper.Warning($"[AG] Cobertura no certificada: {cobertura.Error}. " +
                                    $"Vacantes={detalleVacantes.Count} (materiasConsentidas={contexto.AsignaturasVacantesConsentidas?.Count ?? 0}, sesionesConsentidas={detalleVacantes.Count(d => d.Consentida)}). " +
                                    string.Join("; ", detalleVacantes.Take(10)
                                        .Select(d => $"{d.IdentidadLegible}|asig={d.IdAsignatura}|motivo={d.MotivoVacante}|horas={d.Horas:0.#}")));
                                Log($"[ERROR] Cobertura no certificada: {cobertura.Error}. " +
                                    $"Vacantes consentidas={contexto?.AsignaturasVacantesConsentidas?.Count ?? 0}. " +
                                    "Verifique pool docente o retire asignaturas sin docente.");
                                traza.Causa(AgCausaFallo.CoberturaNoCertificada, cobertura.Error);
                            }
                            else if (detalleVacantes.Count > 0)
                            {
                                // Fase 1920/1921/1925: Cobertura certificada con vacantes declaradas (resultado parcial válido)
                                LoggerHelper.Info($"[AG] Corrida #{idCorrida}: cobertura certificada con {detalleVacantes.Count} vacantes declaradas. " +
                                    string.Join("; ", detalleVacantes.Take(10)
                                        .Select(d => $"{d.IdentidadLegible}|asig={d.IdAsignatura}|motivo={d.MotivoVacante}|horas={d.Horas:0.#}")));
                                Log($"[INFO] Cobertura completada con {detalleVacantes.Count} vacante(s) declarada(s). Se requiere justificación para promover.");
                            }
                        }
                    }
                }

                // FASE 1878/1881/1882: Si la solución no es certificable, purgar el staging huérfano (nunca ante fallos de lectura transitorios)
                bool noCertificada = !stagingOk
                    && (stagingError?.StartsWith("COBERTURA_NO_CERTIFICADA:", StringComparison.Ordinal) == true
                        || stagingError?.StartsWith("PROYECCION_COBERTURA_ERROR:", StringComparison.Ordinal) == true);

                if (noCertificada)
                {
                    var limpieza = await _scheduleRepository.LimpiarStagingAsync(idCorrida, CancellationToken.None);
                    if (limpieza.IsFailure)
                    {
                        LoggerHelper.Warning($"[AG] No se pudo purgar staging de corrida #{idCorrida}: {limpieza.Error}");
                    }
                    else
                    {
                        LoggerHelper.Info($"[AG] Staging no certificable purgado en corrida #{idCorrida}.");
                    }
                }

                corrida.Estado = stagingOk ? (calidad ? AgEstadoCorrida.Completada : AgEstadoCorrida.Completada) : AgEstadoCorrida.Error;
                var cierreCorrida = await _scheduleRepository.UpdateCorridaAsync(corrida, _ctsEjecucion.Token);
                if (cierreCorrida.IsFailure) throw new InvalidOperationException(cierreCorrida.Error);
                corridaFinalizada = true;

                if (stagingOk)
                {
                    EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Completado,
                        calidad ? "HORARIO_FACTIBLE" : "RESULTADO_PARCIAL",
                        calidad ? "Estado: Horario factible."
                            : $"Estado: Resultado parcial. HARD={mejor?.HardViolations ?? 0}; vacantes={mejor?.Vacantes ?? 0}; sesiones vacantes={sesionesVacantes}; horas vacantes={horasVacantes:0.#}.");
                }
                // T6 requiere sustitución transaccional con ratificación vigente; no habilitar promoción aquí.
                btnPromoverProduccion.Enabled = contexto.CandidatosCobertura.Count == 0 && promovible;

                // Fase 599: Resumen de fin de corrida
                LoggerHelper.Resumen("FinCorridaAG",
                    ("IdCorrida", corrida.Id),
                    ("Estado", stagingOk ? corrida.Estado : AgEstadoCorrida.Error),
                    ("Fitness", $"{corrida.MejorFitness:F4}"),
                    ("Hard", corrida.HardViolations),
                    ("Vinculaciones", cantVin),
                    ("Sesiones", cantSes),
                    ("Horas", $"{horasProyectadas:0.#}"));

                if (stagingOk)
                {
                    traza.Resultado(true, $"Optimización exitosa: Fitness={corrida.MejorFitness:F4}, HardViolations={corrida.HardViolations}");
                    traza.Fin(true, $"Corrida #{idCorrida} completada en {_ultimoResultado.GeneracionesEjecutadas} generaciones");

                    Log($"\n{_ultimoResultado.Mensaje}");
                    progressBarAG.Value = 100;
                }
                else
                {
                    traza.Resultado(false, $"Corrida #{idCorrida} sin staging: {stagingError}");
                    traza.Fin(false, $"Corrida #{idCorrida} finalizada con staging vacío");

                    Log($"\n[ERROR] Corrida #{idCorrida} sin staging: {stagingError}");
                    progressBarAG.Value = 100;
                    EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Error, "P08_MaterializarStaging",
                        "Estado: Finalizado con errores en staging.");
                }

                // 4. FASE 1878 & Fase 1907: Mostrar Resultados en Módulo R usando snapshot inmutable 'contexto'
                if (stagingOk)
                {
                    await CargarModuloResultadosAsync(idCorrida);
                    btnPromoverProduccion.Enabled = contexto.CandidatosCobertura.Count == 0 && promovible;
                    tabControlPrincipal.SelectedTab = tabResultados;
                }
                else
                {
                    btnPromoverProduccion.Enabled = false;
                }
            }
            catch (OperationCanceledException)
            {
                traza.Fin(false, "Cancelado por el usuario");
                Log("\n[CANCELADO] Ejecución abortada por el usuario.");
                cancelacionDetectada = true;
                CerrarEstadoVisualOptimizacion(true);
                // Un fallo de presentación no invalida un terminal ya persistido.
            }
            catch (Exception ex)
            {
                traza.Fallo("P08_EjecutarOptimizacionAsync", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                Log($"\n[EXCEPCIÓN NO CONTROLADA] {ex.Message}");
                CerrarEstadoVisualOptimizacion(_ctsEjecucion.IsCancellationRequested);
                // Conservar corridaFinalizada si la persistencia terminal ya tuvo éxito.
            }
            finally
            {
                _operacionEnCurso = false; // Fase 1905: fin de operación en curso
                _tareaOptimizacion = null; // Fase 1900: desasociar tarea finalizada
                _aceptaProgresoOptimizacion = false;
                if (!corridaFinalizada)
                {
                    bool cancelada = cancelacionDetectada || _ctsEjecucion.IsCancellationRequested;
                    if (!_terminalVisualOptimizacion) CerrarEstadoVisualOptimizacion(cancelada);
                    if (idCorridaCreada > 0)
                        await CerrarCorridaFallidaAsync(corrida, cancelada);
                }
                // Un único evento terminal por intento, también si nunca se insertó ag_corrida.
                RegistrarEventoIntento("FIN", intento, idCorridaCreada);
                if (semaforoAdquirido) _semaforoSyncDimensiones.Release();
                treeContextoAcademico.Enabled = arbolHabilitado;
                cmbArboles.Enabled = perfilesHabilitados;
                btnValidarFactibilidad.Enabled = btnValidarHabilitado;
                btnEjecutarAG.Enabled = btnEjecutarHabilitado;
                btnFiltrarArbol.Enabled = btnFiltrarHabilitado;
                btnCargarArbol.Enabled = btnCargarHabilitado;
                btnSobrescribirArbol.Enabled = btnSobrescribirHabilitado;
                btnCancelar.Enabled = false;

                _ctsEjecucion.Dispose();
                _ctsEjecucion = null;
                // Mantener la compuerta hasta terminar también el refresco asíncrono.
                if (idCorridaCreada > 0)
                {
                    try
                    {
                        await RefrescarComboCorridasAsync();
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Warning($"[AG] Error al refrescar el combo de corridas tras optimización: {ex.Message}");
                    }
                }
                _intentoOptimizacion = null;
            }
        }

        private bool PuedeRecibirProgresoOptimizacion(Guid intento)
        {
            return _intentoOptimizacion == intento && _aceptaProgresoOptimizacion &&
                !IsDisposed && !Disposing && IsHandleCreated;
        }

        private void CerrarEstadoVisualOptimizacion(bool cancelada)
        {
            // Catch/finally sólo aportan un respaldo: conservar fase y causa específica para FIN.
            // Cada intento limpia este terminal y su motivo antes del primer await.
            if (_terminalVisualOptimizacion) return;
            EstablecerEstadoOptimizacion(cancelada ? EstadoVisualOptimizacion.Cancelado : EstadoVisualOptimizacion.Error,
                cancelada ? "CANCELACION" : "FALLO_EJECUCION");

            // Fase 1901 / 1908: Restaurar controles de período, ventana, árbol y corridas bloqueados durante la optimización
            if (!IsDisposed && !Disposing)
            {
                cmbPeriodos.Enabled = true;
                chkManualVentana.Enabled = true;
                dtpVentanaInicio.Enabled = chkManualVentana.Checked;
                dtpVentanaFin.Enabled = chkManualVentana.Checked;
                if (cmbArboles != null) cmbArboles.Enabled = true;
                if (treeContextoAcademico != null) treeContextoAcademico.Enabled = true;
                if (btnFiltrarArbol != null) btnFiltrarArbol.Enabled = true;
                if (cmbSedes != null) cmbSedes.Enabled = true;
                if (cmbProgramas != null) cmbProgramas.Enabled = true;
                if (cmbCargarCorridas != null) cmbCargarCorridas.Enabled = true;
            }
        }

        // Los helpers compartidos con diagnóstico sólo afectan a un intento de optimización vigente.
        private void MarcarErrorPreparacion(string motivo)
        {
            if (!_intentoOptimizacion.HasValue || _terminalVisualOptimizacion) return;
            if (_ctsEjecucion?.IsCancellationRequested == true)
                CerrarEstadoVisualOptimizacion(true);
            else
                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Error, motivo);
        }

        private void MarcarDecisionPreparacion()
        {
            if (_intentoOptimizacion.HasValue)
                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.RequiereDecision);
        }

        private void ReanudarValidacionPreparacion()
        {
            if (_intentoOptimizacion.HasValue)
                EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Validando);
        }

        private void EstablecerEstadoOptimizacion(EstadoVisualOptimizacion estado, string motivo = null, string texto = null)
        {
            bool terminal = estado == EstadoVisualOptimizacion.Completado || estado == EstadoVisualOptimizacion.Bloqueado ||
                estado == EstadoVisualOptimizacion.Cancelado || estado == EstadoVisualOptimizacion.Error;
            if (terminal)
            {
                if (!_terminalVisualOptimizacion) _faseFinOptimizacion = _estadoVisualOptimizacion.ToString();
                _motivoFinOptimizacion = motivo;
            }
            _estadoVisualOptimizacion = estado;
            _terminalVisualOptimizacion = terminal;
            _aceptaProgresoOptimizacion = estado == EstadoVisualOptimizacion.Optimizando;
            if (IsDisposed || Disposing) return;
            if (terminal)
            {
                btnPromoverProduccion.Enabled = false;
                // Fase 1901 / 1908: Restaurar controles en estados terminales
                cmbPeriodos.Enabled = true;
                chkManualVentana.Enabled = true;
                dtpVentanaInicio.Enabled = chkManualVentana.Checked;
                dtpVentanaFin.Enabled = chkManualVentana.Checked;
                if (cmbArboles != null) cmbArboles.Enabled = true;
                if (treeContextoAcademico != null) treeContextoAcademico.Enabled = true;
                if (btnFiltrarArbol != null) btnFiltrarArbol.Enabled = true;
                if (cmbSedes != null) cmbSedes.Enabled = true;
                if (cmbProgramas != null) cmbProgramas.Enabled = true;
                if (cmbCargarCorridas != null) cmbCargarCorridas.Enabled = true;
            }
            else if (estado == EstadoVisualOptimizacion.Optimizando || estado == EstadoVisualOptimizacion.Materializando)
            {
                // Fase 1901 / 1908: Bloquear mutaciones de período, ventana, árbol, jerarquía y corrida durante optimización y materialización
                cmbPeriodos.Enabled = false;
                dtpVentanaInicio.Enabled = false;
                dtpVentanaFin.Enabled = false;
                chkManualVentana.Enabled = false;
                if (cmbArboles != null) cmbArboles.Enabled = false;
                if (treeContextoAcademico != null) treeContextoAcademico.Enabled = false;
                if (btnFiltrarArbol != null) btnFiltrarArbol.Enabled = false;
                if (cmbSedes != null) cmbSedes.Enabled = false;
                if (cmbProgramas != null) cmbProgramas.Enabled = false;
                if (cmbCargarCorridas != null) cmbCargarCorridas.Enabled = false;
            }
            Color color = Color.DarkBlue;
            switch (estado)
            {
                case EstadoVisualOptimizacion.Preparando: texto = texto ?? "Estado: Preparando..."; break;
                case EstadoVisualOptimizacion.Validando: texto = texto ?? "Estado: Validando..."; break;
                case EstadoVisualOptimizacion.RequiereDecision:
                    texto = texto ?? "Estado: Decisión pendiente de confirmación."; color = Color.DarkOrange; break;
                case EstadoVisualOptimizacion.Optimizando: texto = texto ?? "Estado: Optimizando..."; break;
                case EstadoVisualOptimizacion.Materializando: texto = texto ?? "Estado: Materializando..."; break;
                case EstadoVisualOptimizacion.Completado:
                    texto = texto ?? "Estado: Completado.";
                    color = motivo == "RESULTADO_PARCIAL" ? Color.DarkOrange : Color.DarkGreen; break;
                case EstadoVisualOptimizacion.Bloqueado:
                    texto = texto ?? "Estado: Bloqueado. Revise el diagnóstico."; color = Color.DarkOrange; break;
                case EstadoVisualOptimizacion.Cancelado:
                    texto = texto ?? "Estado: Cancelado por el usuario."; color = Color.DarkOrange; break;
                case EstadoVisualOptimizacion.Error:
                    texto = texto ?? "Estado: Error en ejecución."; color = Color.DarkRed; break;
            }
            lblStatusAG.Text = texto;
            lblStatusAG.ForeColor = color;
        }

        private void RegistrarEventoIntento(string evento, Guid intento, long idCorrida)
        {
            // LoggerHelper es el canal durable existente (cola a archivo), no txtLogAG ni autorización T4.
            // Sólo metadatos de control: no mensajes libres, snapshots ni datos personales.
            LoggerHelper.Info($"[AG-INTENTO] evento={evento} intentoId={intento:D} corridaId={idCorrida} " +
                $"fase={_faseFinOptimizacion ?? _estadoVisualOptimizacion.ToString()} " +
                $"estado={_estadoVisualOptimizacion} motivo={_motivoFinOptimizacion ?? "NINGUNO"}");
        }

        private async Task CerrarCorridaFallidaAsync(AgCorrida corrida, bool cancelada)
        {
            if (corrida == null || corrida.Id <= 0) return;
            corrida.Estado = cancelada ? AgEstadoCorrida.Cancelada : AgEstadoCorrida.Error;
            corrida.FechaFinCorrida = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            // Independiente del token de trabajo cancelado, pero con presupuesto propio de cierre.
            using (var ctsCierre = new CancellationTokenSource(TimeSpan.FromSeconds(5)))
            {
                try
                {
                    var cierre = await _scheduleRepository.UpdateCorridaAsync(corrida, ctsCierre.Token);
                    if (cierre.IsFailure)
                    {
                        string motivo = ctsCierre.IsCancellationRequested ? "Timeout de cierre (5 s)" : "Fallo de cierre";
                        LoggerHelper.Error($"[AG] {motivo} de la corrida #{corrida.Id} como {corrida.Estado}: {cierre.Error}");
                    }
                }
                catch (OperationCanceledException ex) when (ctsCierre.IsCancellationRequested)
                {
                    LoggerHelper.Error($"[AG] Timeout de cierre (5 s) de la corrida #{corrida.Id} como {corrida.Estado}", ex);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[AG] Excepción al cerrar la corrida #{corrida.Id} como {corrida.Estado}", ex);
                }
            }
        }

        private void CancelarOptimizacion()
        {
            _ctsEjecucion?.Cancel();
        }

        /// <summary>
        /// Carga y visualiza los resultados de la corrida especificada en la pestaña Módulo R (Fases 217, 222, 314, 1289 y 1557).
        /// Carga primero el grid de sesiones de staging y, si la corrida cuenta con individuo óptimo persistido,
        /// materializa sus métricas y violaciones. Si es staging sin resultado AG, muestra sesiones en modo solo lectura.
        /// </summary>
        /// <param name="idCorrida">Identificador de la corrida a visualizar.</param>
        /// <returns>True si la corrida se cargó y materializó exitosamente; false si abortó por falta de datos o error de staging.</returns>
        public async Task<bool> CargarModuloResultadosAsync(long idCorrida)
        {
            _idCorridaCargada = idCorrida;

            // Fase 314: Rehidratar _ultimoResultado desde SQLite si no existe en la sesión viva
            bool cabeceraAusente = false;
            if (_ultimoResultado?.MejorIndividuo == null)
            {
                var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
                if (corridaRes.IsFailure || corridaRes.Value == null)
                {
                    // Fase 1565: no se descarta el motivo real. Antes se asumía "sin cabecera" incluso ante
                    // errores de lectura (columnas ausentes, bloqueo, etc.), ocultando la causa al usuario.
                    cabeceraAusente = true;
                    LoggerHelper.Warning(
                        $"[AG] No se pudo leer la cabecera de la corrida #{idCorrida} en ag_corrida: " +
                        $"{(corridaRes.IsFailure ? corridaRes.Error : "sin filas")}");
                }
                else
                {
                    var corridaDb = corridaRes.Value;
                    AgIndividuo indDb = null;
                    if (corridaDb.IdIndividuoOptimo.HasValue && corridaDb.IdIndividuoOptimo.Value > 0)
                    {
                        var indRes = await _scheduleRepository.GetIndividuoByIdAsync(corridaDb.IdIndividuoOptimo.Value, CancellationToken.None);
                        if (indRes.IsSuccess)
                        {
                            indDb = indRes.Value;
                        }
                        else
                        {
                            LoggerHelper.Warning(
                                $"[AG] Corrida #{idCorrida}: no se pudo leer el individuo óptimo " +
                                $"#{corridaDb.IdIndividuoOptimo.Value}: {indRes.Error}");
                        }
                    }

                    _ultimoResultado = new AgResultadoCorrida
                    {
                        IdCorrida = idCorrida,
                        Estado = corridaDb.Estado,
                        GeneracionesEjecutadas = corridaDb.GeneracionesEjecutadas,
                        MejorIndividuo = indDb ?? new AgIndividuo
                        {
                            IdCorrida = idCorrida,
                            Fitness = corridaDb.MejorFitness,
                            HardViolations = corridaDb.HardViolations,
                            SoftViolations = corridaDb.SoftViolations
                        }
                    };
                }
            }

            // Fase 1557: Cargar siempre el grid de sesiones antes de evaluar la existencia del individuo óptimo
            if (!await CargarGridStagingAsync(idCorrida))
            {
                _idCorridaCargada = 0;
                _ultimoResultado = null;
                btnAprobarCorrida.Enabled = false;
                btnRecalcularViolaciones.Enabled = false;
                btnPromoverProduccion.Enabled = false;
                lblMejorFitness.Text = "Aptitud (Fitness): —";
                lblHardViolaciones.Text = "Restricciones Hard: —";
                lblHardViolaciones.ForeColor = Color.DarkOrange;
                lblVacantes.Text = "Vacantes: —";
                lblVacantes.ForeColor = Color.DarkOrange;
                dgvHorariosGenerados.DataSource = null;
                dgvViolaciones.DataSource = null;
                _todasLasVacantes = new List<AgVacanteGridDto>();
                dgvVacantes.DataSource = null;
                lblConteoVacantesTab.Text = "0 vacante(s)";
                lblAgModo.Text = "Modo AG · staging: ag_vinculaciones / ag_sesiones";
                return false;   // error de lectura o staging vacío (ya limpian estado)
            }

            if (_ultimoResultado?.MejorIndividuo == null)
            {
                // Fase 1565: staging sin resultado AG. Se distingue "cabecera ausente" (corrida huérfana,
                // reparable desde esta misma pantalla) de "cabecera presente sin individuo óptimo"
                // (requiere volver a ejecutar el AG).
                lblMejorFitness.Text = "Aptitud (Fitness): —";
                lblHardViolaciones.Text = cabeceraAusente
                    ? "Violaciones: — (staging huérfano: sin cabecera en ag_corrida)"
                    : "Restricciones Hard: — (sin individuo óptimo persistido)";
                lblHardViolaciones.ForeColor = Color.DarkOrange;
                lblVacantes.Text = "Vacantes: —";

                // El botón existente «Recalcular violaciones» es también la vía de reparación: reconstruye
                // cabecera e individuo desde el staging y recalcula el detalle en ag_violacion.
                btnRecalcularViolaciones.Enabled = true;
                btnPromoverProduccion.Enabled = false;
                btnAprobarCorrida.Enabled = false;

                lblAgModo.Text = cabeceraAusente
                    ? $"Modo AG · corrida #{idCorrida} · staging huérfano (sin ag_corrida): use «Recalcular violaciones» para repararla"
                    : $"Modo AG · corrida #{idCorrida} · sin individuo óptimo persistido (solo lectura)";

                LoggerHelper.Warning(cabeceraAusente
                    ? $"[AG] La corrida #{idCorrida} tiene staging pero no cabecera en ag_corrida (huérfana). " +
                      "«Recalcular violaciones» intentará reconstruir cabecera e individuo desde el staging."
                    : $"[AG] La corrida #{idCorrida} tiene cabecera en ag_corrida pero no individuo óptimo persistido.");

                return true;   // el grid ya está cargado con las sesiones
            }

            _idCorridaCargada = idCorrida;
            lblAgModo.Text = $"Modo AG · corrida #{idCorrida} · {AgModoDatos.Ag.VistaSesionesDetalle}";

            var ind = _ultimoResultado.MejorIndividuo;
            var traza = CrearTraza("CargarResultados");
            traza.Inicio($"idCorrida={idCorrida}");
            traza.Paso("P09_CargarResultados", $"Fitness={ind.Fitness:F4}, HardViolations={ind.HardViolations}, SoftViolations={ind.SoftViolations:F2}");

            lblMejorFitness.Text = $"Aptitud (Fitness): {ind.Fitness:F4}";
            var viols = ind.Violaciones ?? new List<AgViolacion>();
            int registrosHard = viols.Count(v => v.EsDura);
            int registrosSoft = viols.Count(v => !v.EsDura);
            lblHardViolaciones.Text =
                $"Registros HARD={registrosHard}, SOFT={registrosSoft}; " +
                $"infracciones HARD={ind.HardViolations}; penalización SOFT={ind.SoftViolations:0.##}.";
            lblHardViolaciones.ForeColor = ind.HardViolations == 0 ? Color.DarkGreen : Color.DarkRed;
            if (ind.Violaciones != null && ind.Violaciones.Count > 0 && viols.Count != ind.Violaciones.Count)
            {
                LoggerHelper.Warning("[AG] Cantidad de registros persistidos distinta de la evaluada.");
            }

            // Fase 320: Métrica y visualización de Vacantes
            lblVacantes.Text = $"Vacantes: {ind.Vacantes}";
            lblVacantes.ForeColor = ind.Vacantes == 0 ? Color.DarkGreen : Color.DarkOrange;

            // Regla histórica de la carga de resultados: sólo una corrida sin vacantes ni violaciones duras
            // habilita la promoción. T5/T6 añaden que una corrida con manifiesto o decisiones de cobertura
            // pendientes no habilita promoción desde esta ruta.
            bool sinPendientesHistoricos = ind.Vacantes == 0 && ind.HardViolations == 0;
            bool tieneAlcanceCobertura = _contextoActual?.ManifiestoCobertura != null
                || (_contextoActual?.CandidatosCobertura?.Count ?? 0) > 0;
            btnPromoverProduccion.Enabled = sinPendientesHistoricos && !tieneAlcanceCobertura;
            // Fase 1531: Habilitar recálculo de violaciones si la corrida tiene individuo persistido (cubre ruta AG y rehidratación)
            btnRecalcularViolaciones.Enabled = ind.Id > 0;
            btnAprobarCorrida.Enabled = ind.Id > 0;
            return true;
        }

        /// <summary>
        /// Carga en las grillas de resultados las sesiones de staging y violaciones de la corrida (Fases 217, 222 y 1557).
        /// Extraído en Fase 1557 para permitir visualizar corridas que posean sesiones materializadas aunque carezcan de individuo óptimo.
        /// </summary>
        /// <param name="idCorrida">Identificador de la corrida a cargar.</param>
        /// <returns>True si el staging contenía sesiones y se enlazaron con éxito; false si hubo error o estaba vacío.</returns>
        private async Task<bool> CargarGridStagingAsync(long idCorrida)
        {
            var traza = CrearTraza("CargarGridStaging");
            traza.Inicio($"idCorrida={idCorrida}");

            // Cargar Grid de Horarios Staging (Fases 217 y 222: 16 columnas auditables)
            var sesionesResult = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, CancellationToken.None);
            var vinculacionesResult = await _scheduleRepository.GetVinculacionesStagingAsync(idCorrida, CancellationToken.None);
            var vincsDict = vinculacionesResult.IsSuccess
                ? vinculacionesResult.Value.ToDictionary(v => v.Id, v => v)
                : new Dictionary<int, AgVinculacion>();

            traza.Conteo("SesionesStaging", sesionesResult.IsSuccess ? sesionesResult.Value.Count : 0);
            traza.Conteo("VinculacionesStaging", vincsDict.Count);
            traza.Conteo("Violaciones", _ultimoResultado?.MejorIndividuo?.Violaciones?.Count ?? 0, 0);

            if (sesionesResult.IsFailure)
            {
                // Fase 1290: Evitar pantalla mixta (métricas corrida N con grid corrida N-1) y deshabilitar promoción
                Log($"[ERROR] Error al consultar sesiones de staging para corrida #{idCorrida}: {sesionesResult.Error}");
                dgvHorariosGenerados.DataSource = null;
                _todasLasVacantes = new List<AgVacanteGridDto>();
                dgvVacantes.DataSource = null;
                lblConteoVacantesTab.Text = "0 vacante(s)";
                lblMejorFitness.Text = "Aptitud (Fitness): —";
                lblHardViolaciones.Text = "Restricciones Hard: —";
                lblVacantes.Text = "Vacantes: —";
                btnPromoverProduccion.Enabled = false;
                btnRecalcularViolaciones.Enabled = false; // Fase 1539: deshabilitar recálculo ante error de staging
                return false;
            }

            if (sesionesResult.Value.Count == 0 &&
                !(_contextoActual?.IdCorrida == idCorrida && _contextoActual.SesionesConservadas.Count > 0))
            {
                dgvHorariosGenerados.DataSource = null;
                _todasLasVacantes = new List<AgVacanteGridDto>(); // Fase 847: Limpiar vacantes previas
                dgvVacantes.DataSource = null;                    // Fase 847: Limpiar grid de vacantes
                lblConteoVacantesTab.Text = "0 vacante(s)";       // Fase 847: Reiniciar contador
                lblMejorFitness.Text = "Aptitud (Fitness): —";
                lblHardViolaciones.Text = "Restricciones Hard: —";
                lblVacantes.Text = "Vacantes: —";
                btnPromoverProduccion.Enabled = false;
                btnRecalcularViolaciones.Enabled = false; // Fase 1539: deshabilitar recálculo ante staging vacío
                Log($"[AVISO] La corrida #{idCorrida} no tiene sesiones materializadas en staging. No se muestran métricas para evitar datos obsoletos.");
                return false;
            }

            // Diagnóstico de individuo huérfano (Fase 1557): si no hay resultado AG, auditar si las sesiones referencian individuos
            if (_ultimoResultado?.MejorIndividuo == null)
            {
                var inds = sesionesResult.Value.Select(s => s.IdIndividuo).Where(i => i > 0).Distinct().ToList();
                if (inds.Count > 0)
                {
                    LoggerHelper.Warning($"[AG] Corrida #{idCorrida}: staging referencia individuo(s) [{string.Join(", ", inds)}] sin fila en ag_individuo.");
                }
            }

            if (sesionesResult.IsSuccess)
            {
                var lista = sesionesResult.Value.Select(s =>
                {
                    vincsDict.TryGetValue(s.IdVinculacion, out var v);
                    AgAula aula = null;
                    if (_contextoActual?.AulasPorId != null)
                    {
                        _contextoActual.AulasPorId.TryGetValue(s.IdAula, out aula);
                    }

                    Tercero doc = null;
                    if (_contextoActual?.DocentesPorId != null)
                    {
                        _contextoActual.DocentesPorId.TryGetValue(s.IdTercero, out doc);
                    }

                    int idPeriodo = v?.IdPeriodo ?? aula?.IdPeriodo ?? _contextoActual?.IdPeriodo ?? 0;
                    int idSede = v?.IdSede ?? aula?.IdSede ?? 0;
                    int idPrograma = v?.IdPrograma ?? aula?.IdPrograma ?? _contextoActual?.IdPrograma ?? 0;
                    int idAsignatura = v?.IdAsignatura ?? 0;

                    string periodoNom = _periodos.FirstOrDefault(p => p.Id == idPeriodo)?.Nombre ?? idPeriodo.ToString();
                    string sedeNom = _sedes.FirstOrDefault(sd => sd.Id == idSede)?.Ciudad ?? idSede.ToString();
                    string progNom = _programas.FirstOrDefault(pr => pr.Id == idPrograma)?.NombrePrograma ?? idPrograma.ToString();
                    string asigNom = (_contextoActual?.AsignaturasPorId.ContainsKey(idAsignatura) == true)
                        ? _contextoActual.AsignaturasPorId[idAsignatura].Nombre
                        : (idAsignatura > 0 ? $"Asignatura #{idAsignatura}" : "");

                    string semestre = v?.Semestre ?? aula?.Semestre ?? "";
                    string grupo = v?.Grupo ?? aula?.Grupo ?? "";
                    string modalidad = (aula != null && aula.EsVirtual == 1)
                        ? ModalidadTutoriaCatalog.Plataforma
                        : ModalidadTutoriaCatalog.NormalizarTutoria(v?.ModalidadVinculacion ?? aula?.ModalidadVinculacion);
                    string docenteNom = (s.EsVacante == 1 || s.IdTercero <= 0)
                        ? $"VACANTE ({s.MotivoVacante})"
                        : (doc?.Nombre ?? (_contextoActual?.DocentesPorId.ContainsKey(s.IdTercero) == true ? _contextoActual.DocentesPorId[s.IdTercero].Nombre : $"Docente #{s.IdTercero}"));
                    string aulaCod = aula?.Codigo ?? (_contextoActual?.AulasPorId.ContainsKey(s.IdAula) == true ? _contextoActual.AulasPorId[s.IdAula].Codigo : $"Aula #{s.IdAula}");
                    string aulaRef = aula?.AulaRef ?? "";

                    // Fase 461: Trazabilidad completa en grid Módulo R
                    return new
                    {
                        IdSesionGlobal = (long)s.Id,
                        IdCorrida = idCorrida,
                        IdGen = s.IdGen,
                        IdVinculacion = (long)s.IdVinculacion,
                        Origen = "NUEVA",
                        Fecha = s.FechaSesion.HasValue ? s.FechaSesion.Value.ToString("dd/MM/yyyy") : "S/F",
                        Dia = s.DiaSemanaTexto,
                        Horario = s.HorarioSesion,
                        Horas = s.Horas,
                        Periodo = periodoNom,
                        Sede = sedeNom,
                        Programa = progNom,
                        Semestre = semestre,
                        Grupo = grupo,
                        Modalidad = ModalidadTutoriaCatalog.AEtiqueta(modalidad),
                        Asignatura = asigNom,
                        Docente = docenteNom,
                        Aula = aulaCod,
                        AulaRef = aulaRef,
                        Viatica = s.Viatica ? "Sí" : "No",
                        Vacante = (s.EsVacante == 1 || s.IdTercero <= 0) ? "Sí" : "No",
                        Motivo = (s.EsVacante == 1 || s.IdTercero <= 0) ? (s.MotivoVacante ?? string.Empty) : "",
                        Candidatos = s.CandidatosPool
                    };
                }).ToList();

                // La unión es sólo una vista; no inserta ni duplica sesiones conservadas en staging.
                if (_contextoActual?.IdCorrida == idCorrida && _contextoActual.SesionesConservadas != null)
                {
                    foreach (var fija in _contextoActual.SesionesConservadas)
                    {
                        var clave = fija.Clave;
                        var aulaFija = _contextoActual.Aulas?.FirstOrDefault(a => a.Id == clave.IdAula);
                        var docenteFijo = _contextoActual.Docentes?.FirstOrDefault(d => d.Id == fija.IdTercero);
                        lista.Add(new
                        {
                            IdSesionGlobal = fija.Id,
                            IdCorrida = idCorrida,
                            IdGen = 0L,
                            IdVinculacion = 0L,
                            Origen = "CONSERVADA",
                            Fecha = fija.Fecha.ToString("dd/MM/yyyy"),
                            Dia = fija.Fecha.ToString("dddd"),
                            Horario = fija.Inicio.ToString(@"hh\:mm") + " A " + fija.Fin.ToString(@"hh\:mm"),
                            Horas = fija.Horas,
                            Periodo = _periodos.FirstOrDefault(p => p.Id == clave.IdPeriodo)?.Nombre ?? clave.IdPeriodo.ToString(),
                            Sede = _sedes.FirstOrDefault(sd => sd.Id == clave.IdSede)?.Ciudad ?? clave.IdSede.ToString(),
                            Programa = _programas.FirstOrDefault(pr => pr.Id == clave.IdPrograma)?.NombrePrograma ?? clave.IdPrograma.ToString(),
                            Semestre = clave.Semestre,
                            Grupo = clave.Grupo,
                            Modalidad = ModalidadTutoriaCatalog.AEtiqueta(clave.Modalidad),
                            Asignatura = _contextoActual.Asignaturas?.FirstOrDefault(a => a.Id == clave.IdAsignatura)?.Nombre ?? clave.IdAsignatura.ToString(),
                            Docente = docenteFijo?.Nombre ?? $"Docente #{fija.IdTercero}",
                            Aula = aulaFija?.Codigo ?? $"Aula #{clave.IdAula}",
                            AulaRef = aulaFija?.AulaRef ?? "",
                            Viatica = "No",
                            Vacante = "No",
                            Motivo = "",
                            Candidatos = 0
                        });
                    }
                    var m = _contextoActual.ManifiestoCobertura;
                    if (m != null)
                        Log($"[COBERTURA] aulas={m.Relaciones.Select(r => r.Clave.IdAula).Distinct().Count()} " +
                            $"relaciones={m.Relaciones.Count} conservadas={_contextoActual.SesionesConservadas.Select(s => s.Clave.IdAula).Distinct().Count()} " +
                            $"replanificadas={_contextoActual.CandidatosCobertura.Count(c => c.Desde != c.Hacia)} " +
                            $"esperadas={m.Relaciones.Sum(r => r.HorasEfectivas):0.##}h cubiertas={lista.Sum(x => x.Horas):0.##}h " +
                            $"vacantes={sesionesResult.Value.Count(s => s.EsVacante == 1)}");
                }
                dgvHorariosGenerados.DataSource = lista;

                // Fase 322: Poblar el Panel de Vacantes
                _todasLasVacantes = sesionesResult.Value
                    .Where(s => s.EsVacante == 1 || s.IdTercero <= 0)
                    .Select(s =>
                    {
                        vincsDict.TryGetValue(s.IdVinculacion, out var v);
                        AgAula aula = null;
                        if (_contextoActual?.AulasPorId != null)
                        {
                            _contextoActual.AulasPorId.TryGetValue(s.IdAula, out aula);
                        }

                        int idPrograma = v?.IdPrograma ?? aula?.IdPrograma ?? _contextoActual?.IdPrograma ?? 0;
                        int idAsignatura = v?.IdAsignatura ?? 0;
                        string progNom = _programas.FirstOrDefault(pr => pr.Id == idPrograma)?.NombrePrograma ?? idPrograma.ToString();
                        string asigNom = (_contextoActual?.AsignaturasPorId.ContainsKey(idAsignatura) == true)
                            ? _contextoActual.AsignaturasPorId[idAsignatura].Nombre
                            : (idAsignatura > 0 ? $"Asignatura #{idAsignatura}" : "");
                        string semestre = v?.Semestre ?? aula?.Semestre ?? "";
                        string aulaCod = aula?.Codigo ?? (_contextoActual?.AulasPorId.ContainsKey(s.IdAula) == true ? _contextoActual.AulasPorId[s.IdAula].Codigo : $"Aula #{s.IdAula}");
                        string aulaRef = aula?.AulaRef ?? "";

                        return new AgVacanteGridDto
                        {
                            IdSesion = s.Id,
                            IdVinculacion = s.IdVinculacion,
                            IdPrograma = idPrograma,
                            Programa = progNom,
                            IdAsignatura = idAsignatura,
                            Asignatura = asigNom,
                            Semestre = semestre,
                            MotivoVacante = AgMotivoVacanteExtensions.Desde(s.MotivoVacante),
                            CandidatosPool = s.CandidatosPool,
                            Fecha = s.FechaSesion.HasValue ? s.FechaSesion.Value.ToString("dd/MM/yyyy") : "S/F",
                            Dia = s.DiaSemanaTexto,
                            Horario = s.HorarioSesion,
                            Aula = aulaCod,
                            AulaRef = aulaRef
                        };
                    }).ToList();

                FiltrarYMostrarVacantes();
            }

            // Fase 1559: Si el individuo en memoria o rehidratado no tiene violaciones cargadas, consultarlas desde ag_violacion
            var indActual = _ultimoResultado?.MejorIndividuo;
            if (indActual != null && indActual.Id > 0 &&
                (indActual.Violaciones == null || indActual.Violaciones.Count == 0))
            {
                var violRes = await _scheduleRepository.GetViolacionesByIndividuoAsync(indActual.Id, CancellationToken.None);
                if (violRes.IsSuccess && violRes.Value.Count > 0)
                {
                    indActual.Violaciones = violRes.Value.ToList();
                }
            }

            if (indActual == null)
            {
                Log("[AG] Corrida sin resultado AG: no hay violaciones persistidas que mostrar.");
            }

            // Cargar Grid de Violaciones (Fase 460: ordenado por EsDura descendente, luego por Regla)
            var violaciones = (_ultimoResultado?.MejorIndividuo?.Violaciones ?? new List<AgViolacion>())
                .OrderByDescending(v => v.EsDura)
                .ThenBy(v => v.Regla.ToString())
                .Select(v => new
                {
                    Tipo = v.EsDura ? "HARD (Bloqueante)" : "SOFT (Ponderada)",
                    Regla = v.Regla.ToString(),
                    Mensaje = v.Mensaje
                }).ToList();

            dgvViolaciones.DataSource = violaciones;
            traza.Fin(true, $"Horarios cargados en grid ({sesionesResult.Value?.Count ?? 0} sesiones, {_todasLasVacantes.Count} vacantes)");
            return true;
        }

        /// <summary>
        /// Refresca el selector de corridas del Módulo R (Fase 1519, 1542, 1543, 1609, 1615, 1617, 1618, 1621, 1622, 1623, 1624):
        /// lista las corridas del período, programa y territorial activos con su volumen de staging y preselecciona la corrida cargada.
        /// Aplica cascada tolerante de filtros para no dejar vacío el selector si hay staging en la territorial.
        /// Protegido con guarda de reentrada, inhibición de mutación con popup abierto y procesamiento de solicitudes pendientes.
        /// </summary>
        private async Task RefrescarComboCorridasAsync(bool permitirCargaInicial = false)
        {
            if (cmbCargarCorridas == null) return;
            if (_refrescandoComboCorridas)
            {
                _cargaInicialPendiente |= permitirCargaInicial; // Fase 1624: acumular autorización si coincide con refresco en curso
                if (ClaveFiltroCorridas() != _claveFiltroEnVuelo) // Fase 1637: encolar solo si los filtros cambiaron realmente
                {
                    _solicitudPendienteRefrescoCombo = true;
                }
                return;
            }
            _refrescandoComboCorridas = true;
            _claveFiltroEnVuelo = ClaveFiltroCorridas(); // Fase 1637: registrar clave activa al inicio del refresco

            try
            {
                do
                {
                    permitirCargaInicial |= _cargaInicialPendiente; // Fase 1624: consumir permiso acumulado
                    _cargaInicialPendiente = false;
                    _solicitudPendienteRefrescoCombo = false;
                    if (IsDisposed || Disposing) return;

                    string claveFiltrosAlIniciar = ClaveFiltroCorridas();

                    var periodo = cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo);
                    var programa = cmbProgramas.Primario<Programa>() ?? (cmbProgramas.SelectedItem as Programa);

                    int? idPeriodo = periodo != null && periodo.Id > 0 ? (int?)periodo.Id : null;
                    int? idPrograma = programa != null && programa.Id > 0 ? (int?)programa.Id : null;
                    string subunidad = string.IsNullOrWhiteSpace(SyncContext.ActiveSubunidad) ? null : SyncContext.ActiveSubunidad;

                    var repo = new VinculacionQueryRepository(_connectionFactory, AgModoDatos.Ag);
                    var items = new List<CorridaStagingItem>();
                    string error = null;
                    bool ampliado = false;

                    // Fase 1542: cascada tolerante (período+programa -> período sin programa -> sin período ni programa)
                    foreach (var (fPeriodo, fPrograma) in AgCorridaSelector.CascadaFiltros(idPeriodo, idPrograma))
                    {
                        if (IsDisposed || Disposing) return;
                        var intento = await repo.GetCorridasStagingAsync(fPeriodo, fPrograma, subunidad, CancellationToken.None);
                        if (IsDisposed || Disposing) return;

                        if (intento.IsFailure) { error = intento.Error; break; }

                        items = intento.Value.Where(c => c.IdCorrida > 0).ToList();
                        if (items.Count > 0)
                        {
                            ampliado = fPeriodo != idPeriodo || fPrograma != idPrograma;
                            break;
                        }
                    }

                    if (error != null)
                    {
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] No se pudieron listar corridas AG: {error}");
                        continue;
                    }

                    // Fases 1618 & 1621: si la corrida activa no está en items filtrados, buscarla sin filtro y preservar marcador
                    CorridaStagingItem fueraDeFiltro = null;
                    if (_idCorridaCargada > 0 && !items.Any(c => c.IdCorrida == _idCorridaCargada))
                    {
                        var activaRes = await repo.GetCorridasStagingAsync(null, null, subunidad, CancellationToken.None);
                        fueraDeFiltro = activaRes.IsSuccess
                            ? activaRes.Value.FirstOrDefault(c => c.IdCorrida == _idCorridaCargada)
                            : null;
                        if (fueraDeFiltro != null)
                        {
                            items.Add(fueraDeFiltro);
                        }
                    }

                    if (ClaveFiltroCorridas() != claveFiltrosAlIniciar)
                    {
                        _solicitudPendienteRefrescoCombo = true;
                        continue;
                    }

                    // Fases 1615, 1622, 1623, 1631, 1634: cerrar temporalmente el popup si está abierto para mutar Items con seguridad
                    bool reabrirPopup = cmbCargarCorridas.DroppedDown;
                    if (reabrirPopup)
                    {
                        _cerrandoPopupInterno = true;
                        try { cmbCargarCorridas.DroppedDown = false; }
                        finally { _cerrandoPopupInterno = false; }
                    }

                    AgSeleccionModuloR seleccion = null;
                    _actualizandoComboCorridas = true;
                    try
                    {
                        // Fase 1543: limpiar enlace de datos antes de manipular Items
                        cmbCargarCorridas.DisplayMember = string.Empty;
                        cmbCargarCorridas.ValueMember = string.Empty;
                        cmbCargarCorridas.Items.Clear();
                        cmbCargarCorridas.SelectedIndex = -1;
                        cmbCargarCorridas.Text = string.Empty;

                        if (items.Count == 0)
                        {
                            if (_idCorridaCargada <= 0)
                            {
                                lblAgModo.Text = $"Modo AG · sin corridas en staging (período #{idPeriodo?.ToString() ?? "todos"}, programa #{idPrograma?.ToString() ?? "todos"}, territorial {subunidad ?? "todas"})";
                                btnAprobarCorrida.Enabled = false;
                                btnRecalcularViolaciones.Enabled = false;
                                btnPromoverProduccion.Enabled = false;
                            }
                            else
                            {
                                lblAgModo.Text = $"Modo AG · corrida cargada #{_idCorridaCargada} no visible con los filtros actuales";
                            }
                            Log("[AG] No hay corridas de staging para los filtros activos.");
                            continue;
                        }

                        if (ampliado && _idCorridaCargada <= 0)
                            lblAgModo.Text = "Modo AG · corridas ampliadas (sin filtro período/programa)";

                        foreach (var it in items) it.Etiqueta = AgCorridaSelector.FormatearEtiqueta(it);
                        // Fase 1621: sufijo aplicado tras el formateo para no sobrescribir la marca de fuera de filtro
                        if (fueraDeFiltro != null) fueraDeFiltro.Etiqueta += " (fuera de filtro)";

                        cmbCargarCorridas.DisplayMember = nameof(CorridaStagingItem.Etiqueta);
                        cmbCargarCorridas.ValueMember = nameof(CorridaStagingItem.IdCorrida);
                        foreach (var it in items) cmbCargarCorridas.Items.Add(it);

                        // Fase 1608: resolución segura del selector de corridas en Módulo R
                        seleccion = AgCorridaSelector.ResolverSeleccionModuloR(
                            items, _idCorridaCargada > 0 ? _idCorridaCargada : (long?)null);
                        cmbCargarCorridas.SelectedIndex = seleccion.IndiceSeleccionado;

                        if (seleccion.CorridaActivaNoEncontrada)
                        {
                            lblAgModo.Text = $"Modo AG · corrida cargada #{_idCorridaCargada} no visible con los filtros actuales";
                        }

                        _ultimoRefrescoComboUtc = DateTime.UtcNow; // Fase 1632 & 1635: registrar marca de tiempo estrictamente tras swap efectivo
                    }
                    finally
                    {
                        _actualizandoComboCorridas = false;
                        // Fases 1631, 1633, 1634: reabrir de forma diferida en el finally para cubrir también el camino items.Count == 0
                        if (reabrirPopup && !IsDisposed && !Disposing && IsHandleCreated)
                        {
                            ReabrirPopupDiferido();
                        }
                    }

                    // Fase 1608 & 1617: carga inicial automática solo si está explícitamente autorizada al abrir el formulario
                    if (permitirCargaInicial &&
                        seleccion != null &&
                        seleccion.RequiereCargaInicial &&
                        seleccion.ItemSeleccionado != null &&
                        seleccion.ItemSeleccionado.IdCorrida > 0 &&
                        seleccion.ItemSeleccionado.IdCorrida != _idCorridaCargada)
                    {
                        if (IsDisposed || Disposing) return;
                        await CargarCorridaDesdeComboAsync();
                        if (IsDisposed || Disposing) return;
                    }
                }
                while (_solicitudPendienteRefrescoCombo && !IsDisposed && !Disposing);
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgFormGeneradorHorarios] Error al refrescar el selector de corridas: {ex.Message}");
            }
            finally
            {
                _claveFiltroEnVuelo = null; // Fase 1637: liberar clave en vuelo
                _refrescandoComboCorridas = false;
            }
        }

        /// <summary>
        /// Clave que identifica la combinación de filtros actuales del selector de corridas (Fase 1544).
        /// </summary>
        private string ClaveFiltroCorridas()
        {
            var p = (cmbPeriodos.Primario<Periodo>() ?? cmbPeriodos.SelectedItem as Periodo)?.Id ?? 0;
            var g = (cmbProgramas.Primario<Programa>() ?? cmbProgramas.SelectedItem as Programa)?.Id ?? 0;
            return $"{p}|{g}|{SyncContext.ActiveSubunidad}";
        }

        private void RefrescarComboCorridasSeguro(bool permitirCargaInicial = false, bool omitirSiReciente = false)
        {
            // Fase 1632: deduplicación temporal para evitar sobrecarga de consultas y parpadeo de UI
            if (omitirSiReciente && !permitirCargaInicial &&
                (DateTime.UtcNow - _ultimoRefrescoComboUtc).TotalMilliseconds < 750)
            {
                return;
            }
            _ = RefrescarComboCorridasAsync(permitirCargaInicial);
        }

        /// <summary>
        /// Reabre asíncronamente el desplegable del selector inhibiendo la suscripción al evento DropDown (Fases 1631, 1633, 1634, 1636).
        /// Valida que el combo mantenga el foco para no robarlo indebidamente.
        /// </summary>
        private void ReabrirPopupDiferido()
        {
            _reabriendoPopupCombo++;
            try
            {
                BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // Fase 1636: sólo reabrir si el combo conserva el foco y nadie abrió otro control
                        if (!IsDisposed && !Disposing && IsHandleCreated && cmbCargarCorridas.Focused)
                        {
                            cmbCargarCorridas.DroppedDown = true;
                        }
                    }
                    catch (InvalidOperationException) { /* handle destruido durante el cierre */ }
                    finally
                    {
                        _reabriendoPopupCombo--;
                    }
                }));
            }
            catch (InvalidOperationException)
            {
                _reabriendoPopupCombo--;
            }
        }

        /// <summary>
        /// Carga en el Módulo R la corrida elegida en el selector (Fase 1519, 1616, 1625), rehidratando grid,
        /// métricas, vacantes y aprobando/habilitando la promoción según corresponda. Coalesce selecciones en vuelo.
        /// </summary>
        private async Task CargarCorridaDesdeComboAsync()
        {
            if (_actualizandoComboCorridas || _cargandoCorridaDesdeCombo)
            {
                // Fase 1625: no descartar selección del usuario ocurrida mientras otra carga está en vuelo
                if (!_actualizandoComboCorridas && cmbCargarCorridas?.SelectedItem is CorridaStagingItem pend
                    && pend.IdCorrida != _idCorridaCargada)
                {
                    _recargaCorridaPendiente = true;
                }
                return;
            }
            if (!(cmbCargarCorridas?.SelectedItem is CorridaStagingItem corrida) || corrida.IdCorrida <= 0) return;

            _cargandoCorridaDesdeCombo = true;

            try
            {
                Cursor = Cursors.WaitCursor;
                _ultimoResultado = null;   // Fase 1284: rehidratación limpia desde SQLite
                bool cargada = await CargarModuloResultadosAsync(corrida.IdCorrida);
                _ultimaCargaDesdeComboOk = cargada; // Fase 1629: registrar éxito de carga para despacho seguro

                if (!cargada)
                {
                    _idCorridaCargada = 0;
                    _ultimoResultado = null;
                    btnAprobarCorrida.Enabled = false;
                    btnRecalcularViolaciones.Enabled = false;
                    btnPromoverProduccion.Enabled = false;
                    // Fase 1555: no dejar métricas ni grids de la corrida previa cuando la carga no se completa
                    lblMejorFitness.Text = "Aptitud (Fitness): —";
                    lblHardViolaciones.Text = "Restricciones Hard: —";
                    lblVacantes.Text = "Vacantes: —";
                    dgvHorariosGenerados.DataSource = null;
                    dgvViolaciones.DataSource = null;
                    lblAgModo.Text = "Modo AG · staging: ag_vinculaciones / ag_sesiones";
                    Log($"[AG] La corrida #{corrida.IdCorrida} no se pudo cargar (sin sesiones materializadas o error de lectura).");
                    return;
                }

                lblAgModo.Text = $"Modo AG · corrida #{corrida.IdCorrida} · {AgModoDatos.Ag.VistaSesionesDetalle}";

                var ind = _ultimoResultado?.MejorIndividuo;
                btnAprobarCorrida.Enabled = ind != null && ind.Id > 0;
                // Fase 1565: el recálculo queda disponible también para corridas huérfanas (staging sin cabecera),
                // donde es la única vía de reparación sin volver a ejecutar el AG.
                bool esHuerfana = _ultimoResultado == null;
                btnRecalcularViolaciones.Enabled = (ind != null && ind.Id > 0) || esHuerfana;
                // Fase 1522, 1813 & 1815: compuerta centralizada estricta que rechaza lecturas fallidas y valida cobertura
                var corridaDbRes = await _scheduleRepository.GetCorridaByIdAsync(corrida.IdCorrida, CancellationToken.None);
                var sesionesRes = await _scheduleRepository.GetSesionesStagingAsync(corrida.IdCorrida, CancellationToken.None);
                var corridaEfectiva = corridaDbRes.IsSuccess ? corridaDbRes.Value : null;

                // Corrección [P1]: Asegurar contexto de la corrida cargada para evaluar factibilidad y compuerta de promoción
                if (corridaEfectiva != null && (_contextoActual == null || _contextoActual.IdPeriodo != corridaEfectiva.IdPeriodo || _contextoActual.IdPrograma != corridaEfectiva.IdPrograma))
                {
                    var ctxRes = await ConstruirContextoAsync(CancellationToken.None, corridaEfectiva.IdPeriodo, corridaEfectiva.IdPrograma, corridaEfectiva.Subunidad);
                    if (ctxRes.IsSuccess)
                    {
                        _contextoActual = ctxRes.Value;
                    }
                }

                // FASE 1871/1885: Restaurar asignaturas vacantes consentidas desde propiedad tipada o fallback ParametrosJson
                if (_contextoActual != null && corridaEfectiva != null)
                {
                    if (corridaEfectiva.VacantesConsentidas != null && corridaEfectiva.VacantesConsentidas.Count > 0)
                    {
                        _contextoActual.AsignaturasVacantesConsentidas = new HashSet<int>(corridaEfectiva.VacantesConsentidas);
                    }
                    else if (!string.IsNullOrWhiteSpace(corridaEfectiva.ParametrosJson))
                    {
                        try
                        {
                            var ser = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 512 * 1024 };
                            var dict = ser.DeserializeObject(corridaEfectiva.ParametrosJson) as Dictionary<string, object>;
                            if (dict != null && dict.TryGetValue("vacantesConsentidas", out var vacObj) && vacObj is System.Collections.IEnumerable list)
                            {
                                var ids = new HashSet<int>();
                                foreach (var item in list)
                                {
                                    if (item != null && int.TryParse(item.ToString(), out int id)) ids.Add(id);
                                }
                                _contextoActual.AsignaturasVacantesConsentidas = ids;
                                corridaEfectiva.VacantesConsentidas = ids;
                            }
                        }
                        catch (Exception exJson)
                        {
                            LoggerHelper.Warning($"[AG] Error al deserializar vacantes consentidas desde ParametrosJson: {exJson.Message}");
                        }
                    }
                }

                // Fase 1923: permitirVacantes: true para habilitar botón; la compuerta real es la justificación al promover
                btnPromoverProduccion.Enabled = corridaDbRes.IsSuccess && AgCompuertaPromocion.EsPromovible(corridaEfectiva, ind, sesionesRes, _contextoActual, permitirVacantes: true);
                LoggerHelper.Info($"[AG] Corrida #{corrida.IdCorrida} cargada: individuo #{ind?.Id ?? 0}, HARD={ind?.HardViolations ?? 0}, promover={btnPromoverProduccion.Enabled}.");
                Log($"[AG] Corrida #{corrida.IdCorrida} cargada desde el selector ({corrida.TotalSesiones} sesiones en staging).");
            }
            catch (OperationCanceledException)
            {
                _ultimaCargaDesdeComboOk = false;
            }
            catch (Exception ex)
            {
                _ultimaCargaDesdeComboOk = false;
                LoggerHelper.Error($"[AgFormGeneradorHorarios] Error al cargar la corrida #{corrida.IdCorrida} desde el selector", ex);
                MessageBox.Show(this, $"No se pudo cargar la corrida #{corrida.IdCorrida}:\n{ex.Message}",
                    "Cargar corrida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cargandoCorridaDesdeCombo = false;
                Cursor = Cursors.Default;
                // Fases 1625, 1629 & 1630: limpiar siempre flag pendiente y verificar IsHandleCreated antes de BeginInvoke
                bool despachar = _recargaCorridaPendiente && _ultimaCargaDesdeComboOk;
                _recargaCorridaPendiente = false;

                if (despachar && !IsDisposed && !Disposing && IsHandleCreated)
                {
                    try
                    {
                        BeginInvoke(new Action(() => { if (!IsDisposed && !Disposing) _ = CargarCorridaDesdeComboAsync(); }));
                    }
                    catch (InvalidOperationException) { /* handle destruido durante el cierre */ }
                }
            }
        }

        /// <summary>
        /// Repara una corrida huérfana (staging en ag_vinculaciones/ag_sesiones sin cabecera en ag_corrida
        /// ni individuo en ag_individuo) reutilizando los genes persistidos en ag_gen y el contexto del AG
        /// (Fase 1566). Se invoca desde «Recalcular violaciones», sin añadir controles nuevos a la interfaz.
        /// </summary>
        /// <param name="idCorrida">Corrida a reparar.</param>
        /// <returns>True si la corrida quedó reparada y recargada en el Módulo R.</returns>
        private async Task<bool> IntentarRepararCorridaHuerfanaAsync(long idCorrida)
        {
            var diagnostico = await _scheduleRepository.DetectarCorridaHuerfanaAsync(idCorrida, CancellationToken.None);
            if (diagnostico.IsFailure || !diagnostico.Value.PuedeRepararse)
            {
                LoggerHelper.Warning(
                    $"[AG-Reparacion] Corrida #{idCorrida} no reparable: " +
                    $"{(diagnostico.IsFailure ? diagnostico.Error : diagnostico.Value.ToString())}");
                return false;
            }

            var huerfana = diagnostico.Value;
            string periodoNom = _periodos.FirstOrDefault(p => p.Id == huerfana.IdPeriodo)?.Nombre
                ?? huerfana.IdPeriodo.ToString();
            string programaNom = _programas.FirstOrDefault(pr => pr.Id == huerfana.IdPrograma)?.NombrePrograma
                ?? huerfana.IdPrograma.ToString();

            var confirmar = MessageBox.Show(this,
                $"La corrida #{idCorrida} conserva {huerfana.TotalVinculaciones} vinculación(es) y " +
                $"{huerfana.TotalSesiones} sesión(es) en staging, pero perdió su cabecera en ag_corrida " +
                $"y su individuo (#{huerfana.IdIndividuo}).\n\n" +
                $"Se reconstruirá con período «{periodoNom}» y programa «{programaNom}» (datos del staging) " +
                "y se recalcularán las violaciones. No se vuelve a ejecutar el algoritmo genético.\n\n" +
                "Requiere que el árbol tenga seleccionadas las cohortes de esa corrida. ¿Continuar?",
                "Reparar corrida huérfana", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmar != DialogResult.Yes) return false;

            _contextoActual = null;
            var ctxRes = await ConstruirContextoAsync(
                CancellationToken.None, huerfana.IdPeriodo, huerfana.IdPrograma, huerfana.Subunidad);
            if (ctxRes.IsFailure)
            {
                MessageBox.Show(this,
                    "No se pudo reconstruir el contexto del AG para reparar la corrida:\n" + ctxRes.Error +
                    "\n\nVerifique el período, el programa, el semestre y las aulas del árbol.",
                    "Reparar corrida huérfana", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            _contextoActual = ctxRes.Value;


            var servicio = new AgReparacionCorridaService(
                _scheduleRepository, new AgEvaluadorAptitud(new AgValidadorReglas()));
            var res = await servicio.RepararAsync(idCorrida, _contextoActual, CancellationToken.None);
            if (res.IsFailure)
            {
                LoggerHelper.Warning($"[AG-Reparacion] corrida #{idCorrida} no reparada: {res.Error}");
                MessageBox.Show(this, "No se pudo reparar la corrida:\n" + res.Error,
                    "Reparar corrida huérfana", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            var r = res.Value;
            LoggerHelper.Info(
                $"[AG-Reparacion] Corrida #{idCorrida} reparada: individuo #{r.IdIndividuo}, " +
                $"HARD={r.Hard}, SOFT={r.Soft:0.##}, violaciones={r.Violaciones}, genes={r.GenesEvaluados}.");
            Log($"[AG-Reparacion] corrida #{idCorrida} reparada · individuo #{r.IdIndividuo} · " +
                $"HARD={r.Hard} · {r.ResumenPorRegla}");

            _ultimoResultado = null;
            await CargarModuloResultadosAsync(idCorrida);

            MessageBox.Show(this,
                $"Corrida #{idCorrida} reparada.\n\n" +
                $"• Cabecera reconstruida: período «{periodoNom}», programa «{programaNom}»\n" +
                $"• Individuo #{r.IdIndividuo} con {r.GenesEvaluados} gen(es)\n" +
                $"• Violaciones: HARD {r.Hard} · SOFT {r.Soft:0.##} ({r.Violaciones} en detalle)\n" +
                $"• Desglose: {r.ResumenPorRegla}\n\n" +
                "El detalle quedó persistido en ag_violacion. Apruebe la corrida para habilitar la promoción " +
                "(la compuerta de seguridad exige HARD = 0).",
                "Corrida reparada", MessageBoxButtons.OK,
                r.Hard == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            return true;
        }


        /// <summary>
        /// Recalcula las violaciones de la corrida cargada reutilizando el contexto del AG y persiste
        /// el desglose en ag_violacion (Fase 1527). No vuelve a ejecutar el algoritmo genetico.
        /// </summary>
        private async Task RecalcularViolacionesDesdeModuloRAsync()
        {
            long idCorrida = _ultimoResultado?.IdCorrida ?? _idCorridaCargada;
            if (idCorrida <= 0)
            {
                MessageBox.Show(this, "Cargue una corrida en el selector antes de recalcular sus violaciones.",
                    "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(this,
                    $"¿Recalcular las violaciones de la corrida #{idCorrida}?\n\nSe reevalúa la solución guardada con el contexto actual del AG y se reemplaza el detalle persistido en ag_violacion. No se vuelve a ejecutar el algoritmo genético.",
                    "Recalcular violaciones", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }

            btnRecalcularViolaciones.Enabled = false;
            var cursorPrevio = Cursor;
            Cursor = Cursors.WaitCursor;
            try
            {
                var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
                if (corridaRes.IsFailure)
                {
                    // Fase 1566: si la corrida conserva staging pero perdió su cabecera, se ofrece reconstruirla
                    // (ag_corrida + ag_individuo + ag_violacion) antes de reportar el error de lectura.
                    if (await IntentarRepararCorridaHuerfanaAsync(idCorrida))
                    {
                        return;
                    }

                    MessageBox.Show(this, $"No se pudo consultar la corrida #{idCorrida}:\n{corridaRes.Error}",
                        "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var corrida = corridaRes.Value;
                // Fase 1550: advertencia explícita ante corridas sin cabecera en ag_corrida (huérfanas)
                if (corrida.IdPeriodo <= 0 || corrida.IdPrograma <= 0)
                {
                    MessageBox.Show(this,
                        $"La corrida #{idCorrida} no tiene cabecera en ag_corrida (período/programa). " +
                        "Registre la corrida o recálculela desde el AG.",
                        "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 1546: reconstruir contexto forzando período, programa y subunidad de la corrida si no coincide
                if (_contextoActual == null ||
                    _contextoActual.IdPeriodo != corrida.IdPeriodo ||
                    _contextoActual.IdPrograma != corrida.IdPrograma)
                {
                    _contextoActual = null;
                    var ctxRes = await ConstruirContextoAsync(
                        CancellationToken.None, corrida.IdPeriodo, corrida.IdPrograma, corrida.Subunidad);
                    if (ctxRes.IsFailure)
                    {
                        MessageBox.Show(this,
                            "No se pudo reconstruir el contexto del AG para recalcular:\n" + ctxRes.Error +
                            "\n\nVerifique el período, el programa, el semestre y las aulas del árbol.",
                            "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _contextoActual = ctxRes.Value;
                }

                var servicio = new AgRecalculoViolacionesService(_scheduleRepository, new AgValidadorReglas());
                var res = await servicio.RecalcularAsync(idCorrida, _contextoActual, CancellationToken.None);
                if (res.IsFailure)
                {
                    LoggerHelper.Warning($"[AG-Recalculo] corrida #{idCorrida} no recalculada: {res.Error}");
                    MessageBox.Show(this, "No se pudieron recalcular las violaciones:\n" + res.Error,
                        "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var r = res.Value;
                LoggerHelper.Info($"[AG-Recalculo] corrida #{idCorrida}: HARD={r.Hard} SOFT={r.Soft:0.##} genes={r.GenesEvaluados} asignaturas={r.AsignaturasContexto} · {r.ResumenPorRegla}");
                Log($"[AG-Recalculo] corrida #{idCorrida} · HARD={r.Hard} SOFT={r.Soft:0.##} · {r.ResumenPorRegla}");

                // Fase 1530: Forzar rehidratación de _ultimoResultado desde SQLite para reflejar el recálculo
                _ultimoResultado = null;
                await CargarModuloResultadosAsync(idCorrida);
                var ind = _ultimoResultado?.MejorIndividuo;

                // Fase 1537 & Fase 1548: reflejar siempre en el panel el detalle recién persistido (ind puede ser null si staging quedó vacío)
                var violacionesLista = r.Violaciones?.ToList() ?? new List<AgViolacion>();
                if (ind != null)
                {
                    ind.Violaciones = violacionesLista;
                }
                dgvViolaciones.DataSource = violacionesLista
                    .OrderByDescending(v => v.EsDura)
                    .ThenBy(v => v.Regla.ToString())
                    .Select(v => new
                    {
                        Tipo = v.EsDura ? "HARD (Bloqueante)" : "SOFT (Ponderada)",
                        Regla = v.Regla.ToString(),
                        Mensaje = v.Mensaje
                    }).ToList();
                dgvViolaciones.Refresh();

                // Fase 1815: compuerta centralizada tras recálculo
                var corridaDbRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
                var sesionesRes = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, CancellationToken.None);
                var corridaEfectiva = corridaDbRes.IsSuccess ? corridaDbRes.Value : null;
                // Fase 1923: permitirVacantes: true para habilitar botón; la compuerta real es la justificación al promover
                btnPromoverProduccion.Enabled = corridaDbRes.IsSuccess && AgCompuertaPromocion.EsPromovible(corridaEfectiva, ind, sesionesRes, _contextoActual, permitirVacantes: true);

                MessageBox.Show(this,
                    $"Recálculo completado para la corrida #{idCorrida}.\n\n" +
                    $"Violaciones HARD: {r.Hard}\nViolaciones SOFT: {r.Soft:0.##}\n" +
                    $"Genes evaluados: {r.GenesEvaluados} · asignaturas en contexto: {r.AsignaturasContexto}\n\n" +
                    $"Desglose: {r.ResumenPorRegla}\n\n" +
                    "El detalle quedó persistido en ag_violacion y es el que verá la vista por fechas.",
                    "Recalcular violaciones", MessageBoxButtons.OK,
                    r.Hard == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AG-Recalculo] Error al recalcular violaciones de la corrida #{idCorrida}", ex);
                MessageBox.Show(this, "Error al recalcular las violaciones:\n" + ex.Message,
                    "Recalcular violaciones", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = cursorPrevio;
                // Fase 1547: no reactivar incondicionalmente; respetar guardas si no hay individuo óptimo
                btnRecalcularViolaciones.Enabled = _ultimoResultado?.MejorIndividuo?.Id > 0;
            }
        }

        /// <summary>
        /// Aprueba formalmente la corrida cargada y sus aulas para habilitar la compuerta de
        /// promoción a producción (Fase 1520). Reutiliza <see cref="AprobarCorridaYAulaAsync"/>.
        /// </summary>
        private async Task AprobarCorridaDesdeModuloRAsync()
        {
            long idCorrida = _ultimoResultado?.IdCorrida ?? _idCorridaCargada;
            if (idCorrida <= 0)
            {
                MessageBox.Show(this, "Cargue una optimización en el selector antes de aprobarla.",
                    "Aprobar corrida", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(this,
                    $"¿Aprobar formalmente la corrida #{idCorrida} y sus aulas?\n\nEsto habilita la promoción a producción.",
                    "Aprobar corrida", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }

            btnAprobarCorrida.Enabled = false;
            try
            {
                await AprobarCorridaYAulaAsync();

                // Fase 1522: verificar el estado real tras la aprobación y habilitar la promoción
                var verif = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
                if (verif.IsSuccess && verif.Value.Estado == AgEstadoCorrida.Aprobada)
                {
                    Log($"[AG] Corrida #{idCorrida} aprobada formalmente (aulas en estado Aprobado).");
                    LoggerHelper.Info($"[AG] Corrida #{idCorrida} aprobada (estado Aprobada). Promoción habilitada.");

                    var indActual = _ultimoResultado?.MejorIndividuo;
                    var sesionesRes = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, CancellationToken.None);
                    // Fase 1923: permitirVacantes: true para habilitar botón; la compuerta real es la justificación al promover
                    btnPromoverProduccion.Enabled = indActual != null && sesionesRes.IsSuccess && AgCompuertaPromocion.EsPromovible(verif.Value, indActual, sesionesRes, _contextoActual, permitirVacantes: true);
                }
                else
                {
                    string detalle = verif.IsFailure ? verif.Error : verif.Value.Estado.ToString();
                    LoggerHelper.Warning($"[AG] La corrida #{idCorrida} no quedó aprobada tras AprobarCorridaYAulaAsync (estado: {detalle}).");
                    MessageBox.Show(this, $"La corrida no quedó aprobada (estado: {detalle}).\nRevise el log para más detalle.",
                        "Aprobar corrida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgFormGeneradorHorarios] Error al aprobar la corrida #{idCorrida}", ex);
                MessageBox.Show(this, $"No se pudo aprobar la corrida:\n{ex.Message}",
                    "Aprobar corrida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnAprobarCorrida.Enabled = true;
            }
        }

        private async Task PromoverAProduccionAsync()
        {
            // T6: una corrida con manifiesto requiere reemplazo transaccional todavía no implementado.
            // Comprobar la cabecera persistida también tras reiniciar la aplicación, antes de cualquier escritura.
            long idParaPromover = _ultimoResultado?.IdCorrida > 0 ? _ultimoResultado.IdCorrida : _idCorridaCargada;
            if (idParaPromover > 0)
            {
                var alcancePersistido = await _scheduleRepository.GetCorridaByIdAsync(idParaPromover, CancellationToken.None);
                if (alcancePersistido.IsFailure || alcancePersistido.Value == null)
                {
                    MessageBox.Show(this, "No se pudo comprobar el alcance persistido de la corrida.",
                        "Promoción bloqueada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if ((alcancePersistido.Value.ParametrosJson ?? string.Empty).Contains("\"manifiestoCobertura\""))
                {
                    btnPromoverProduccion.Enabled = false;
                    MessageBox.Show(this, "Promoción bloqueada: falta reemplazo transaccional T6 de la corrida con manifiesto.",
                        "Promoción bloqueada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            if ((_ultimoResultado == null || _ultimoResultado.MejorIndividuo == null) && _idCorridaCargada > 0)
            {
                await CargarModuloResultadosAsync(_idCorridaCargada);
            }

            if (_ultimoResultado == null || _ultimoResultado.MejorIndividuo == null)
            {
                // Fase 1565: antes retornaba en silencio y el usuario percibía que "promover no funciona".
                LoggerHelper.Warning(
                    $"[AG] Promover: la corrida #{_idCorridaCargada} no tiene individuo óptimo rehidratado; se cancela la promoción.");
                MessageBox.Show(this,
                    $"La corrida #{_idCorridaCargada} no tiene individuo óptimo persistido.\n\n" +
                    "Pulse «Recalcular violaciones» para reconstruir la corrida (si conserva staging) " +
                    "o vuelva a ejecutar el AG.",
                    "Promover a producción", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var ind = _ultimoResultado.MejorIndividuo;
            var traza = CrearTraza("PromoverProduccion");
            traza.Inicio($"Corrida={_ultimoResultado.IdCorrida}, Individuo={ind.Id}");
            traza.Paso("P10_PromoverAProduccion", "Promoción de soluciones desde staging hacia tablas de producción");
            LoggerHelper.Info($"[AG] Promover: corrida #{_ultimoResultado.IdCorrida}, individuo #{ind.Id}, HARD={ind.HardViolations}, vacantes={ind.Vacantes}.");

            // Fase 1520 (G3): sin individuo óptimo persistido no hay nada que promover
            if (ind.Id <= 0)
            {
                traza.Resultado(false, "Individuo óptimo sin registro en ag_individuo");
                traza.Fin(false, "Individuo no rehidratado");
                MessageBox.Show(
                    $"La corrida #{_ultimoResultado.IdCorrida} no tiene individuo óptimo registrado.\n\nVuelva a ejecutar el AG o cargue otra corrida.",
                    "Compuerta de Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (ind.HardViolations > 0)
            {
                LoggerHelper.Warning($"[AG] Promoción bloqueada: individuo #{ind.Id} registra {ind.HardViolations} violación(es) HARD (requiere 0).");
                traza.Resultado(false, $"Bloqueado por {ind.HardViolations} violaciones duras activas");
                traza.Fin(false, "No se puede promover con restricciones HARD");
                MessageBox.Show(
                    $"No se puede promover a producción: la solución registra {ind.HardViolations} violación(es) HARD.\n\n" +
                    "Corrija la programación y vuelva a ejecutar el AG.\n" +
                    "Si el panel «Violaciones de Reglas» aparece vacío, el detalle no quedó persistido en ag_violacion; " +
                    "una nueva corrida lo regenerará.",
                    "Compuerta de Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string justificacionVacante = null;
            // Fases 320 y 330: Advertencia y justificación explícita si existen vacantes
            if (ind.Vacantes > 0)
            {
                var confirmVac = MessageBox.Show(
                    $"Esta solución contiene {ind.Vacantes} sesión(es) en estado VACANTE (sin docente asignado).\n" +
                    "Se promoverán a producción con docente vacío o sin asignar.\n\n" +
                    "¿Desea continuar con la promoción parcial a pesar de las vacantes activas?",
                    "Advertencia de Vacantes Activas",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (confirmVac != DialogResult.Yes) return;
                justificacionVacante = $"Promoción parcial autorizada con {ind.Vacantes} vacante(s) pendientes.";
            }

            var confirm = MessageBox.Show(
                "¿Está seguro de que desea promover esta programación a las tablas de PRODUCCIÓN?\n\n" +
                "Esto insertará las vinculaciones y sesiones correspondientes y sincronizará el pool docente.",
                "Confirmar Promoción a Producción",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            this.Cursor = Cursors.WaitCursor;

            try
            {
                // Gate estricto de promoción (Fase 117): Validar estados Aprobada y Aprobado
                var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(_ultimoResultado.IdCorrida, CancellationToken.None);
                if (corridaRes.IsFailure || corridaRes.Value.Estado != AgEstadoCorrida.Aprobada)
                {
                    string estadoActual = corridaRes.IsFailure ? $"error: {corridaRes.Error}" : corridaRes.Value.Estado.ToString();
                    LoggerHelper.Warning($"[AG] Promoción bloqueada: corrida #{_ultimoResultado.IdCorrida} en estado {estadoActual} (requiere Aprobada).");
                    traza.Fallo("P10_ValidarEstadoCorrida", null, AgCausaFallo.Ninguna);
                    traza.Fin(false, $"Corrida no aprobada formalmente ({estadoActual})");
                    MessageBox.Show(
                        $"La corrida #{_ultimoResultado.IdCorrida} no está aprobada (estado: {estadoActual}).\n\n" +
                        "Use el botón «✅ Aprobar corrida» y vuelva a intentar la promoción.",
                        "Compuerta de Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (corridaRes.Value.IdAula > 0)
                {
                    var aulaRes = await _scheduleRepository.GetAulaByIdAsync(corridaRes.Value.IdAula, CancellationToken.None);
                    if (aulaRes.IsFailure || (aulaRes.Value.Estado != AgEstadoAula.Aprobado && aulaRes.Value.Estado != AgEstadoAula.Disponible))
                    {
                        LoggerHelper.Warning($"[AG] Promoción bloqueada: aula #{corridaRes.Value.IdAula} no aprobada.");
                        traza.Fallo("P10_ValidarEstadoAula", null, AgCausaFallo.Ninguna);
                        traza.Fin(false, "Aula no aprobada");
                        MessageBox.Show("El aula asociada debe estar aprobada para permitir la promoción.", "Compuerta de Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Fase 1520 (G4): deshabilitar sólo cuando la promoción realmente arranca
                btnPromoverProduccion.Enabled = false;

                var resProm = await _promotor.PromoverAProduccionAsync(_ultimoResultado.IdCorrida, ind.Id, Environment.UserName, justificacionVacante, _contextoActual, CancellationToken.None);
                if (resProm.IsSuccess)
                {
                    var p = resProm.Value;
                    LoggerHelper.Info($"[AG] Promoción exitosa: corrida #{_ultimoResultado.IdCorrida} → {p.TotalVinculaciones} vinculaciones / {p.TotalSesiones} sesiones (registro #{p.Id}).");
                    traza.Dato("VinculacionesPromovidas", p.TotalVinculaciones);
                    traza.Dato("SesionesPromovidas", p.TotalSesiones);
                    traza.Resultado(true, $"Promoción exitosa: {p.TotalVinculaciones} vinculaciones, {p.TotalSesiones} sesiones");
                    traza.Fin(true, $"Promoción culminada exitosamente (Registro #{p.Id})");

                    MessageBox.Show(
                        $"Promoción exitosa a producción.\n\n• Vinculaciones creadas: {p.TotalVinculaciones}\n• Sesiones creadas: {p.TotalSesiones}\n• Registro ID: {p.Id}",
                        "Promoción Completada",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    traza.Fallo("P10_PromoverAProduccionAsync", new Exception(resProm.Error), AgCausaFallo.ExcepcionNoControlada);
                    traza.Fin(false, resProm.Error);
                    LoggerHelper.Error($"[AG] Promoción fallida (corrida #{_ultimoResultado.IdCorrida}, individuo #{ind.Id}): {resProm.Error}");
                    MessageBox.Show($"Error durante la promoción: {resProm.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                traza.Fallo("P10_PromoverAProduccionAsync", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error en promoción", ex);
                MessageBox.Show($"Ocurrió un error inesperado: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                // Fase 1520 (G4) & Fase 1815: el botón sólo se re-habilita si la solución sigue siendo promovible bajo compuerta centralizada
                var corridaDbRes = await _scheduleRepository.GetCorridaByIdAsync(_ultimoResultado.IdCorrida, CancellationToken.None);
                var sesionesRes = await _scheduleRepository.GetSesionesStagingAsync(_ultimoResultado.IdCorrida, CancellationToken.None);
                var corridaEfectiva = corridaDbRes.IsSuccess ? corridaDbRes.Value : null;
                // Fase 1923: permitirVacantes: true para habilitar botón; la compuerta real es la justificación al promover
                btnPromoverProduccion.Enabled = corridaDbRes.IsSuccess && AgCompuertaPromocion.EsPromovible(corridaEfectiva, ind, sesionesRes, _contextoActual, permitirVacantes: true);
            }
        }



        /// <summary>
        /// Filtra y actualiza las vacantes mostradas en <see cref="dgvVacantes"/> según el criterio del combo (Fase 322).
        /// </summary>
        private void FiltrarYMostrarVacantes()
        {
            if (_todasLasVacantes == null)
            {
                dgvVacantes.DataSource = null;
                lblConteoVacantesTab.Text = "0 vacante(s)";
                return;
            }

            string filtro = cmbMotivoVacante.SelectedItem?.ToString() ?? "(Todos)";
            var filtradas = (filtro == "(Todos)")
                ? _todasLasVacantes
                : _todasLasVacantes.Where(v => v.Motivo == filtro).ToList();

            dgvVacantes.DataSource = filtradas.ToList();
            lblConteoVacantesTab.Text = $"{filtradas.Count} de {_todasLasVacantes.Count} vacante(s)";
            ActualizarBotonContextualVacante();
        }

        /// <summary>
        /// Actualiza el texto y estado del botón contextual según el motivo de la vacante seleccionada (Fase 322).
        /// </summary>
        private void ActualizarBotonContextualVacante()
        {
            if (dgvVacantes.CurrentRow == null || dgvVacantes.CurrentRow.DataBoundItem == null)
            {
                btnAccionVacante.Text = "Configurar";
                btnAccionVacante.Enabled = false;
                return;
            }

            var item = dgvVacantes.CurrentRow.DataBoundItem as AgVacanteGridDto;
            if (item == null)
            {
                btnAccionVacante.Text = "Configurar";
                btnAccionVacante.Enabled = false;
                return;
            }

            btnAccionVacante.Enabled = true;
            switch (item.MotivoVacante)
            {
                case AgMotivoVacante.SIN_CANDIDATOS:
                case AgMotivoVacante.POOL_VACIO:
                    btnAccionVacante.Text = "Cargar pool";
                    break;
                case AgMotivoVacante.SIN_DISPONIBILIDAD:
                    btnAccionVacante.Text = "Mover franja";
                    break;
                case AgMotivoVacante.NO_CUMPLE_REQUISITOS:
                    btnAccionVacante.Text = "Revisar requisitos";
                    break;
                case AgMotivoVacante.FUERA_DE_VENTANA:
                case AgMotivoVacante.SIN_PLAN_OLEADA:
                    btnAccionVacante.Text = "Ajustar ventana";
                    break;
                case AgMotivoVacante.SIN_BLOQUE_NIVEL:
                    btnAccionVacante.Text = "Configurar bloques";
                    break;
                default:
                    btnAccionVacante.Text = "Configurar";
                    break;
            }
        }

        /// <summary>
        /// Ejecuta la acción contextual correspondiente al motivo de la vacante seleccionada (Fase 322).
        /// </summary>
        private async Task EjecutarAccionVacanteAsync()
        {
            if (dgvVacantes.CurrentRow == null || dgvVacantes.CurrentRow.DataBoundItem == null) return;
            var item = dgvVacantes.CurrentRow.DataBoundItem as AgVacanteGridDto;
            if (item == null) return;

            try
            {
                switch (item.MotivoVacante)
                {
                    case AgMotivoVacante.SIN_CANDIDATOS:
                    case AgMotivoVacante.POOL_VACIO:
                    case AgMotivoVacante.NO_CUMPLE_REQUISITOS:
                        var terceroAsig = new TerceroAsignatura
                        {
                            IdPrograma = item.IdPrograma,
                            IdAsignatura = item.IdAsignatura,
                            Semestre = item.Semestre
                        };
                        using (var form = new FormTerceroAsignaturas(0, terceroAsig, _connectionFactory))
                        {
                            form.ShowDialog(this);
                        }
                        break;

                    case AgMotivoVacante.SIN_DISPONIBILIDAD:
                        AbrirEnModoAg(item.IdVinculacion);
                        break;

                    case AgMotivoVacante.SIN_BLOQUE_NIVEL:
                        MessageBox.Show(
                            $"La asignatura '{item.Asignatura}' no cuenta con un bloque horario activo de la duración requerida para su nivel en los días asignados.\n\n" +
                            "Revise la configuración de bloques horarios en el catálogo institucional.",
                            "Sin Bloque de Nivel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        break;

                    case AgMotivoVacante.FUERA_DE_VENTANA:
                    case AgMotivoVacante.SIN_PLAN_OLEADA:
                        var opcVentana = MessageBox.Show(
                            $"La asignatura '{item.Asignatura}' presenta sesiones o fechas estimadas fuera de la ventana semestral activa.\n\n" +
                            "• Elija 'Sí' para abrir la sesión en Modo AG y ajustar manualmente sus fechas u horarios.\n" +
                            "• Elija 'No' para cambiar a la pestaña de Configuración (Módulo S) y revisar el rango de la ventana semestral.",
                            "Ajustar Ventana de Vacante",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (opcVentana == DialogResult.Yes)
                        {
                            AbrirEnModoAg(item.IdVinculacion);
                        }
                        else if (opcVentana == DialogResult.No)
                        {
                            tabControlPrincipal.SelectedTab = tabSimulacion;
                        }
                        break;

                    default:
                        AbrirEnModoAg(item.IdVinculacion);
                        break;
                }

                if (_idCorridaCargada > 0)
                {
                    await CargarModuloResultadosAsync(_idCorridaCargada);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al ejecutar acción contextual de vacante", ex);
                MessageBox.Show($"Error al ejecutar la acción de vacante: {ex.Message}", "Panel de Vacantes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Genera y persiste un encabezado de Aula (Módulo S) y asocia sus asignaturas correspondientes al semestre/programa.
        /// </summary>
        [Obsolete("Usar GenerarEncabezadosAulasAsync que propaga jornada, horas_dia y dias_ventana seleccionados.")]
        public async Task<Result<long>> GenerarEncabezadoAulaAsync(
            Periodo periodoSel, Sede sedeSel, Programa progSel,
            string semestre, string grupo, string modalidad,
            DateTime fechaInicioAula, DateTime fechaFinAula, CancellationToken ct)
        {
            if (sedeSel == null || periodoSel == null || progSel == null)
            {
                return Result.Failure<long>("Debe especificar Sede, Periodo y Programa.");
            }

            var aula = new AgAula
            {
                IdSede = sedeSel.Id,
                IdPeriodo = periodoSel.Id,
                IdPrograma = progSel.Id,
                Semestre = semestre,
                Grupo = grupo,
                ModalidadVinculacion = modalidad,
                FechaInicio = fechaInicioAula.ToString("yyyy-MM-dd"),
                FechaFin = fechaFinAula.ToString("yyyy-MM-dd"),
                SemestreInicio = _ventanaResuelta != null ? _ventanaResuelta.FechaInicio.ToString("yyyy-MM-dd") : fechaInicioAula.ToString("yyyy-MM-dd"),
                SemestreFin = _ventanaResuelta != null ? _ventanaResuelta.FechaFin.ToString("yyyy-MM-dd") : fechaFinAula.ToString("yyyy-MM-dd"),
                VentanaOrigen = _ventanaResuelta != null ? _ventanaResuelta.Origen.ToString() : "Directa",
                Codigo = $"{progSel.CodPrograma ?? progSel.Nombre}-{semestre}-{grupo}",
                Nombre = $"{progSel.Nombre} {semestre} {grupo}",
                Estado = AgEstadoAula.Disponible,
                AulaRef = AgAulaFactory.ConstruirAulaRef(periodoSel.Id, sedeSel.Id, progSel.Id, semestre, grupo, modalidad),
                Subunidad = SyncContext.ActiveSubunidad ?? string.Empty
            };

            var res = await _scheduleRepository.InsertAulaAsync(aula, ct);
            if (res.IsFailure) return res;

            long idAula = res.Value;
            using (var db = _connectionFactory())
            {
                if (db.State != ConnectionState.Open) db.Open();
                ct.ThrowIfCancellationRequested();
                string sqlAsig = "SELECT id FROM asignaturas WHERE idprograma = @idPrograma AND (semestre = @semestre OR @semestre = '')";
                var asigIds = (await db.QueryAsync<int>(new CommandDefinition(sqlAsig, new { idPrograma = progSel.Id, semestre }, cancellationToken: ct))).ToList();

                if (asigIds.Count > 0)
                {
                    await _scheduleRepository.InsertAulaAsignaturasAsync(idAula, asigIds, ct);
                }
            }

            return Result.Success(idAula);
        }

        #region Estructura y Cálculo Unificado de Columnas Nivel 6 (Ronda 204 - Fases 1040, 1041)

        /// <summary>
        /// Fuente única de verdad para la distribución horizontal proporcional de los radio buttons
        /// y cabeceras de dimensiones en el Nivel 6 (Modalidad).
        /// Distribución: Jornada 30%, Días Ventana 48%, Horas/Día 22% sobre el ancho disponible.
        /// </summary>
        internal struct AgColumnasNivel6
        {
            public int XBase;

            // Grupo A: Jornada (3 opciones: Mañana, Tarde, Noche)
            public int DxManana;
            public int DxTarde;
            public int DxNoche;
            public int DxSep1;

            // Grupo B: Días Ventana (5 opciones: Lun-Jue, Mar-Vie, Jue-Sab, Vie-Sab, Vie-Dom)
            public int DxLunJue;
            public int DxMarVie;
            public int DxJueSab;
            public int DxVieSab;
            public int DxVieDom;
            public int DxSep2;

            // Grupo C: Horas Día (2 opciones: 4h, 8h)
            public int DxHoras4;
            public int DxHoras8;
            public int DxFin;
        }

        private bool _alineandoCabeceras;

        private AgColumnasNivel6 CalcularColumnasNivel6(TreeNode n6 = null)
        {
            int xBase = ObtenerXBaseNivel6(n6);
            int anchoCliente = treeContextoAcademico.ClientSize.Width;
            int anchoDisp = Math.Max(684, anchoCliente - xBase - 150 - 12);

            // Reparto proporcional 30% / 48% / 22%
            int wJornada = Math.Max(186, (int)(anchoDisp * 0.30));
            int wDias = Math.Max(370, (int)(anchoDisp * 0.48));
            int wHoras = Math.Max(128, anchoDisp - wJornada - wDias);

            // Pasos internos con mínimos por control (62, 74, 64 px)
            int stepJ = Math.Max(62, (wJornada - 10) / 3);
            int stepD = Math.Max(74, (wDias - 10) / 5);
            int stepH = Math.Max(64, (wHoras - 10) / 2);

            var c = new AgColumnasNivel6
            {
                XBase = xBase,

                // Grupo A - Jornada
                DxManana = 150,
                DxTarde = 150 + stepJ,
                DxNoche = 150 + 2 * stepJ,
                DxSep1 = 150 + wJornada,

                // Grupo B - Días Ventana
                DxLunJue = 150 + wJornada + 10,
                DxMarVie = 150 + wJornada + 10 + stepD,
                DxJueSab = 150 + wJornada + 10 + 2 * stepD,
                DxVieSab = 150 + wJornada + 10 + 3 * stepD,
                DxVieDom = 150 + wJornada + 10 + 4 * stepD,
                DxSep2 = 150 + wJornada + wDias,

                // Grupo C - Horas Día
                DxHoras4 = 150 + wJornada + wDias + 10,
                DxHoras8 = 150 + wJornada + wDias + 10 + stepH,
                DxFin = 150 + wJornada + wDias + wHoras
            };

            return c;
        }

        #endregion

        /// <summary>
        /// Aprueba la corrida y el aula correspondiente para habilitar la compuerta de promoción a producción (Fase 314).
        /// </summary>
        public async Task AprobarCorridaYAulaAsync()
        {
            if (_ultimoResultado == null && _idCorridaCargada > 0)
            {
                await CargarModuloResultadosAsync(_idCorridaCargada);
            }

            long idCorrida = _ultimoResultado?.IdCorrida ?? _idCorridaCargada;
            if (idCorrida <= 0)
            {
                MessageBox.Show("No hay una corrida activa para aprobar.", "Aprobar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
                if (corridaRes.IsSuccess)
                {
                    var corrida = corridaRes.Value;
                    corrida.Estado = AgEstadoCorrida.Aprobada;
                    var updRes = await _scheduleRepository.UpdateCorridaAsync(corrida, CancellationToken.None);
                    if (updRes.IsFailure)
                    {
                        LoggerHelper.Error($"[AG] No se pudo aprobar la corrida #{corrida.Id}: {updRes.Error}");
                        MessageBox.Show(this, $"No se pudo aprobar la corrida #{corrida.Id}:\n{updRes.Error}",
                            "Aprobar corrida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    LoggerHelper.Info($"[AG] Corrida #{corrida.Id} marcada como Aprobada en ag_corrida.");

                    using (var db = _connectionFactory())
                    {
                        if (db.State != ConnectionState.Open) db.Open();

                        // Fase 543: Multi-aula / multi-sede: aprobar todas las aulas referenciadas en las sesiones de la corrida (y la singular si aplica)
                        var idsAulas = (await db.QueryAsync<int>(
                            @"SELECT DISTINCT s.id_aula 
                              FROM ag_sesiones s 
                              JOIN ag_vinculaciones v ON s.idvinculacion = v.id 
                              WHERE v.id_corrida = @IdCorrida AND s.id_aula > 0",
                            new { IdCorrida = corrida.Id })).ToList();

                        if (corrida.IdAula > 0 && !idsAulas.Contains(corrida.IdAula))
                        {
                            idsAulas.Add(corrida.IdAula);
                        }

                        if (idsAulas.Count > 0)
                        {
                            await db.ExecuteAsync("UPDATE ag_aula SET estado = @Estado WHERE id IN @Ids",
                                new { Estado = (int)AgEstadoAula.Aprobado, Ids = idsAulas });
                        }
                    }

                    MessageBox.Show("Corrida y Aula(s) aprobadas exitosamente (Compuerta de promoción habilitada).", "Aprobado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error aprobando corrida/aula", ex);
                MessageBox.Show($"Error al aprobar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Abre el formulario de edición de sesiones o vinculaciones en Modo AG (Staging).
        /// </summary>
        public void AbrirEnModoAg(int? idVinculacion = null)
        {
            try
            {
                using (var form = new FormSesionesAcademicas(_connectionFactory, null, idVinculacion, AgModoDatos.Ag))
                {
                    form.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al abrir en Modo AG", ex);
                MessageBox.Show($"Error al abrir en Modo AG: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAbrirModoAg_Click(object sender, EventArgs e)
        {
            AbrirEnModoAg(null);
        }

        /// <summary>
        /// Abre el formulario general de vinculaciones en Modo AG (Staging: ag_vinculaciones) desde Módulo R (Fase 298).
        /// Tras cerrar el diálogo, refresca los resultados de la corrida activa si existe.
        /// </summary>
        private async Task AbrirVinculacionesModoAgAsync()
        {
            try
            {
                var modo = AgModoDatos.Ag;
                var queryRepo = new VinculacionQueryRepository(_connectionFactory, modo);
                var sedeRepo = new SedeLookupRepository(_connectionFactory);
                var terceroRepo = new TerceroLookupRepository(_connectionFactory);
                var periodoRepo = new PeriodoLookupRepository(_connectionFactory);
                var programaRepo = new ProgramaLookupRepository(_connectionFactory);
                var listService = new VinculacionListService(queryRepo, sedeRepo, terceroRepo, periodoRepo, programaRepo);

                using (var f = new FormVinculaciones(listService, sedeRepo, terceroRepo, _connectionFactory,
                                                     ExcelDnaUtil.Application, modo,
                                                     _idCorridaCargada > 0 ? _idCorridaCargada : (long?)null))
                {
                    f.ShowDialog(this);
                }

                if (_idCorridaCargada > 0)
                {
                    bool recargada = await CargarModuloResultadosAsync(_idCorridaCargada);
                    if (!recargada)
                    {
                        _idCorridaCargada = 0;
                        _ultimoResultado = null;
                    }
                }
                await RefrescarComboCorridasAsync();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al abrir Vinculaciones (AG)", ex);
                MessageBox.Show($"Error al abrir Vinculaciones en Modo AG: {ex.Message}", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Abre el formulario de detalle de sesiones en árbol/grilla en Modo AG (Staging: vw_ag_sesiones_detalle) desde Módulo R (Fase 298).
        /// Consulta y proyecta las sesiones en formato canónico <see cref="SesionListItem"/> mediante <see cref="VinculacionQueryRepository"/>.
        /// </summary>
        private async Task AbrirDetalleSesionesModoAgAsync()
        {
            try
            {
                if (dgvHorariosGenerados.Rows.Count == 0 && _idCorridaCargada <= 0)
                {
                    MessageBox.Show(this, "Ejecute o cargue una corrida para auditar en Modo AG.", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var modo = AgModoDatos.Ag;
                var queryRepo = new VinculacionQueryRepository(_connectionFactory, modo);

                // Fase 534: si el alcance es multi-sede o multi-programa, no acotar a una sola para no ocultar sesiones
                int? idSede = (_contextoActual?.SedesAlcance != null && _contextoActual.SedesAlcance.Count == 1)
                    ? (int?)_contextoActual.SedesAlcance[0]
                    : ((_contextoActual?.SedesAlcance == null || _contextoActual.SedesAlcance.Count == 0) && _contextoActual?.IdSede > 0 ? (int?)_contextoActual.IdSede : null);

                int? idPeriodo = _contextoActual?.IdPeriodo > 0 ? (int?)_contextoActual.IdPeriodo : null;

                int? idPrograma = (_contextoActual?.ProgramasAlcance != null && _contextoActual.ProgramasAlcance.Count == 1)
                    ? (int?)_contextoActual.ProgramasAlcance[0]
                    : ((_contextoActual?.ProgramasAlcance == null || _contextoActual.ProgramasAlcance.Count == 0) && _contextoActual?.IdPrograma > 0 ? (int?)_contextoActual.IdPrograma : null);

                long? idCorridaDetalle = _ultimoResultado?.IdCorrida > 0 ? (long?)_ultimoResultado.IdCorrida
                                       : _idCorridaCargada > 0 ? (long?)_idCorridaCargada : null;

                CuadreCorrida cuadreEsperado = null;
                if (idCorridaDetalle.HasValue && idCorridaDetalle.Value > 0)
                {
                    // Fase 1931: detalle completo por IdCorrida, sin filtro implícito de programa/sede
                    var corridaFiltro = await _scheduleRepository.GetCorridaByIdAsync(idCorridaDetalle.Value, CancellationToken.None);
                    // sólo para trazar el programa principal; NO para filtrar
                    LoggerHelper.Info($"[AG] Detalle staging: corrida #{idCorridaDetalle.Value} alcance completo (IdPrograma cabecera=" +
                                      $"{(corridaFiltro.IsSuccess && corridaFiltro.Value != null ? corridaFiltro.Value.IdPrograma : 0)}; sin filtro de sede/programa).");

                    idSede = null;
                    idPeriodo = null;
                    idPrograma = null;

                    var resCuadre = await _scheduleRepository.GetCuadreCorridaAsync(idCorridaDetalle.Value, CancellationToken.None);
                    if (resCuadre.IsSuccess)
                    {
                        cuadreEsperado = resCuadre.Value;
                    }
                }

                var res = await queryRepo.GetSesionesAsync(
                    idSede,
                    null,
                    idPeriodo,
                    idPrograma,
                    null,
                    null,
                    null,
                    false,
                    null,
                    null,
                    -1,
                    0,
                    CancellationToken.None,
                    idCorridaDetalle);

                if (res.IsFailure || res.Value == null || res.Value.Count == 0)
                {
                    MessageBox.Show(this, $"No se encontraron sesiones registradas en la vista de detalle staging ({modo.VistaSesionesDetalle}).", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var listaItems = res.Value.ToList();
                // T6: origen visible sin duplicar sesiones ni escribir staging.
                foreach (var item in listaItems) item.OrigenSesion = "NUEVA";

                // Fase 1934: Visibilidad de programas homónimos en log de alcance
                if (idCorridaDetalle.HasValue && idCorridaDetalle.Value > 0)
                {
                    var progs = listaItems.Select(i => i.IdPrograma).Where(p => p > 0).Distinct().OrderBy(p => p).ToList();
                    LoggerHelper.Info($"[AG] Corrida #{idCorridaDetalle.Value} programas=[{string.Join(",", progs)}]");
                }

                using (var db = _connectionFactory())
                using (var f = new FormVerVinculaciones(listaItems, db, ExcelDnaUtil.Application, modo, cuadreEsperado))
                {
                    f.ShowDialog(this);
                }

                if (_idCorridaCargada > 0)
                {
                    await CargarModuloResultadosAsync(_idCorridaCargada);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al abrir Detalle de Sesiones (AG)", ex);
                MessageBox.Show($"Error al abrir Detalle de Sesiones en Modo AG: {ex.Message}", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Abre el formulario unitario de sesión académica en Modo AG (Staging) tomando la vinculación de la fila seleccionada (Fases 298, 313 y 314).
        /// Si la fila seleccionada no posee IdSesion válido o no resuelve su vinculación en ag_sesiones, notifica al usuario sin abrir el diálogo vacío.
        /// </summary>
        private async Task AbrirSesionModoAgAsync()
        {
            try
            {
                int? idVinc = await ObtenerIdVinculacionSeleccionadaAsync();
                if (!idVinc.HasValue || idVinc.Value <= 0)
                {
                    MessageBox.Show(this, "Seleccione una fila de resultados con IdSesion válido para editar en Modo AG.",
                                    "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                AbrirEnModoAg(idVinc);

                if (_idCorridaCargada > 0)
                {
                    await CargarModuloResultadosAsync(_idCorridaCargada);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al abrir Sesión en Modo AG", ex);
                MessageBox.Show($"Error al abrir Sesión en Modo AG: {ex.Message}", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Obtiene el IdVinculacion correspondiente a la fila actualmente seleccionada en <see cref="dgvHorariosGenerados"/> consultando ag_sesiones.
        /// </summary>
        private async Task<int?> ObtenerIdVinculacionSeleccionadaAsync()
        {
            if (dgvHorariosGenerados.CurrentRow == null) return null;

            try
            {
                // Fase 845: Resolver columna IdSesionGlobal o IdSesion de forma insensible a mayúsculas
                var cell = dgvHorariosGenerados.CurrentRow.Cells
                    .Cast<DataGridViewCell>()
                    .FirstOrDefault(c =>
                        string.Equals(c.OwningColumn.Name, "IdSesionGlobal", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(c.OwningColumn.Name, "IdSesion", StringComparison.OrdinalIgnoreCase));

                if (cell == null)
                {
                    LoggerHelper.Warning("[AgFormGeneradorHorarios] Grid sin columna IdSesionGlobal/IdSesion; no se puede resolver la vinculación.");
                    return null;
                }

                if (cell.Value != null && int.TryParse(cell.Value.ToString(), out int idSesion) && idSesion > 0)
                {
                    using (var db = _connectionFactory())
                    {
                        if (db.State != ConnectionState.Open) db.Open();
                        const string sql = "SELECT idvinculacion FROM ag_sesiones WHERE id = @IdSesion LIMIT 1;";
                        int? idVinc = await db.ExecuteScalarAsync<int?>(sql, new { IdSesion = idSesion });
                        return idVinc;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgFormGeneradorHorarios] No se pudo resolver IdVinculacion desde fila seleccionada: {ex.Message}");
            }

            return null;
        }

        private void Log(string mensaje)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(Log), mensaje);
                return;
            }

            string prefijo = (mensaje != null && mensaje.StartsWith("[")) ? string.Empty : $"[{DateTime.Now:HH:mm:ss}] ";
            txtLogAG.AppendText($"{prefijo}{mensaje}\r\n");
            txtLogAG.SelectionStart = txtLogAG.Text.Length;
            txtLogAG.ScrollToCaret();
            // Durable: LoggerHelper.Minimo es Info, así que Debug se descartaría y el detalle se perdería al cerrar.
            if (!string.IsNullOrWhiteSpace(mensaje)) LoggerHelper.Info(mensaje);
        }

        /// <summary>
        /// Crea una nueva instancia de <see cref="AgDiagnosticoTraza"/> cableada a la UI y respetando el estado de chkModoDebug (Fases 277 y 281).
        /// </summary>
        /// <param name="operacion">Nombre de la operación para correlación.</param>
        private AgDiagnosticoTraza CrearTraza(string operacion)
        {
            bool habilitado = chkModoDebug != null && chkModoDebug.Checked;
            LoggerHelper.ModoVerboso = habilitado;
            LoggerHelper.Minimo = habilitado ? NivelLog.Debug : NivelLog.Info; // Fase 605: sincronizar nivel mínimo con modo diagnóstico
            LoggerHelper.Info($"[AG-RANURAS-TMP] DIAGNOSTICO operacion={operacion} activo={habilitado} version={ConversorPDF.XLL.Core.Helpers.App.APP_VERSION}");
            return new AgDiagnosticoTraza(operacion, Log)
            {
                Habilitado = habilitado
            };
        }

        #region Métodos de Gestión del Árbol de 6 Niveles (Fases 157, 158, 160)

        /// <summary>
        /// Puebla el árbol jerárquico de 6 niveles (Periodo -> Sede -> Programa -> Semestre -> Grupo -> Modalidad).
        /// Carga inicialmente hasta nivel Programa y deja nodos dummy para carga bajo demanda (lazy load).
        /// </summary>
        private void PoblarArbolContextoAcademico()
        {
            var traza = CrearTraza("PoblarArbol");
            traza.Inicio();
            traza.Paso("P03_PoblarArbolContextoAcademico", "Construcción jerárquica del árbol de 6 niveles");

            treeContextoAcademico.BeginUpdate();
            treeContextoAcademico.Nodes.Clear();

            var periodosSelIds = cmbPeriodos.SeleccionadosIds();
            var periodos = AgFiltroArbolResolver.FiltrarPeriodos(_periodos, periodosSelIds).ToList();

            var sedesSelIds = cmbSedes.SeleccionadosIds();
            var sedes = AgFiltroArbolResolver.FiltrarSedes(_sedes, sedesSelIds).ToList();

            string tipoProgramaSel = cmbTipoPrograma.SelectedItem?.ToString() ?? AgTipoProgramaCatalog.Todos;
            var programasSelIds = cmbProgramas.SeleccionadosIds();
            var programas = AgFiltroArbolResolver.FiltrarProgramas(_programas, tipoProgramaSel, programasSelIds).ToList();

            foreach (var per in periodos)
            {
                string txtPer = per.Nombre;
                if (!string.IsNullOrWhiteSpace(per.PeriodoInicio) && !string.IsNullOrWhiteSpace(per.PeriodoFin))
                {
                    txtPer += $" ({per.PeriodoInicio} al {per.PeriodoFin})";
                }
                else
                {
                    txtPer += " [Sin fechas]";
                }

                var nodoPer = new TreeNode(txtPer)
                {
                    Tag = new AgNodoArbolKey(1, per.Id, per.Nombre, per.Nombre)
                };

                foreach (var sede in sedes)
                {
                    var nodoSede = new TreeNode(sede.Nombre)
                    {
                        Tag = new AgNodoArbolKey(2, sede.Id, sede.Ciudad, sede.Nombre)
                    };

                    foreach (var prog in programas)
                    {
                        string displayProg = !string.IsNullOrWhiteSpace(prog.Abreviatura)
                            ? $"{prog.Abreviatura} — {prog.NombrePrograma}"
                            : prog.NombrePrograma;

                        // Fase 643: Diferenciación visual de programas homónimos en el árbol
                        if (programas.Count(p => string.Equals(p.NombrePrograma?.Trim(), prog.NombrePrograma?.Trim(),
                                StringComparison.OrdinalIgnoreCase)) > 1)
                        {
                            displayProg += $" (ID {prog.Id})";
                        }

                        var nodoProg = new TreeNode(displayProg)
                        {
                            Tag = new AgNodoArbolKey(3, prog.Id, prog.Abreviatura, prog.NombrePrograma)
                        };

                        // Dummy node para lazy loading de semestres (Nivel 4)
                        nodoProg.Nodes.Add(new TreeNode("Cargando...") { Tag = null });
                        nodoSede.Nodes.Add(nodoProg);
                    }

                    nodoPer.Nodes.Add(nodoSede);
                }

                treeContextoAcademico.Nodes.Add(nodoPer);
            }

            foreach (TreeNode root in treeContextoAcademico.Nodes)
            {
                RefrescarEstadoRecursivo(root); // Fase 542
            }

            if (treeContextoAcademico.Nodes.Count > 0)
            {
                treeContextoAcademico.Nodes[0].Expand();

                // Si hay sede o programa seleccionados en los combos, expandir y revelar la rama correspondiente
                var sedeSel = cmbSedes.Primario<Sede>() ?? (cmbSedes.SelectedItem as Sede);
                var progSel = cmbProgramas.Primario<Programa>() ?? (cmbProgramas.SelectedItem as Programa);
                if (sedeSel != null || progSel != null)
                {
                    foreach (TreeNode nSede in treeContextoAcademico.Nodes[0].Nodes)
                    {
                        if (sedeSel != null && nSede.Tag is AgNodoArbolKey kSede && kSede.Id == sedeSel.Id)
                        {
                            nSede.Expand();
                            if (progSel != null)
                            {
                                foreach (TreeNode nProg in nSede.Nodes)
                                {
                                    if (nProg.Tag is AgNodoArbolKey kProg && kProg.Id == progSel.Id)
                                    {
                                        nProg.EnsureVisible();
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                nSede.EnsureVisible();
                            }
                            break;
                        }
                    }
                }
            }

            treeContextoAcademico.EndUpdate();
            ActualizarConteoCombinaciones();
            ActualizarCabecerasNivel6();
            traza.Dato("TotalNodosRaiz", treeContextoAcademico.Nodes.Count);
            traza.Fin(true, $"Árbol poblado ({treeContextoAcademico.Nodes.Count} nodos raíz)");
        }

        /// <summary>
        /// Carga diferida (lazy load) de niveles inferiores al expandir o marcar un nodo.
        /// Nivel 3 (Programa) -> Semestres I..X.
        /// Nivel 4 (Semestre) -> Grupos A..D.
        /// Nivel 5 (Grupo) -> Modalidades (Presencial, Plataforma, Mixta).
        /// </summary>
        private void CargarHijosLazy(TreeNode node)
        {
            if (node == null || !(node.Tag is AgNodoArbolKey key)) return;

            if (key.Nivel == 3)
            {
                if (node.Nodes.Count == 1 && node.Nodes[0].Tag == null)
                {
                    node.Nodes.Clear();
                    string[] semestres = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" };
                    foreach (var s in semestres)
                    {
                        var nSem = new TreeNode($"Semestre {s}")
                        {
                            Tag = new AgNodoArbolKey(4, 0, s, s)
                        };
                        nSem.Nodes.Add(new TreeNode("Cargando...") { Tag = null });
                        node.Nodes.Add(nSem);
                        RefrescarEstadoNodo(nSem); // Fase 542
                    }
                }
            }
            else if (key.Nivel == 4)
            {
                if (node.Nodes.Count == 1 && node.Nodes[0].Tag == null)
                {
                    node.Nodes.Clear();
                    string[] grupos = { "A", "B", "C", "D" };
                    foreach (var g in grupos)
                    {
                        var nGrp = new TreeNode($"Grupo {g}")
                        {
                            Tag = new AgNodoArbolKey(5, 0, g, g)
                        };
                        nGrp.Nodes.Add(new TreeNode("Cargando...") { Tag = null });
                        node.Nodes.Add(nGrp);
                        RefrescarEstadoNodo(nGrp); // Fase 542
                    }
                }
            }
            else if (key.Nivel == 5)
            {
                if (node.Nodes.Count == 1 && node.Nodes[0].Tag == null)
                {
                    node.Nodes.Clear();
                    int horasSugeridas = 4;
                    TreeNode nodoProg = node.Parent?.Parent;
                    if (nodoProg?.Tag is AgNodoArbolKey keyProg)
                    {
                        var prog = _programas.FirstOrDefault(p => p.Id == keyProg.Id);
                        horasSugeridas = AgTipoProgramaCatalog.HorasDiaSugeridas(prog?.Modalidad);
                    }

                    foreach (var par in ModalidadTutoriaCatalog.Pares)
                    {
                        var key6 = new AgNodoArbolKey(6, 0, par.Codigo, par.Etiqueta)
                        {
                            Jornada = "MANANA",
                            HorasDia = horasSugeridas,
                            DiasVentana = "MAR-VIE"
                        };
                        var nMod = new TreeNode(par.Etiqueta)
                        {
                            Tag = key6
                        };
                        node.Nodes.Add(nMod);
                        RefrescarEstadoNodo(nMod); // Fase 542
                    }
                }
            }
        }

        private void CargarTodoRecursivo(TreeNode node)
        {
            CargarHijosLazy(node);
            foreach (TreeNode child in node.Nodes)
            {
                CargarTodoRecursivo(child);
            }
        }

        #region Glifos de Estado y Radio de Árbol (Fase 542)

        private ImageList CrearImageListEstados()
        {
            var il = new ImageList();
            il.ImageSize = new Size(16, 16);
            il.ColorDepth = ColorDepth.Depth32Bit;

            // [0] CheckBox unchecked
            var bmp0 = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp0))
            {
                if (VisualStyleRenderer.IsSupported)
                    CheckBoxRenderer.DrawCheckBox(g, new Point(0, 0), CheckBoxState.UncheckedNormal);
                else
                    ControlPaint.DrawCheckBox(g, 0, 0, 16, 16, ButtonState.Normal);
            }
            il.Images.Add(bmp0);

            // [1] CheckBox checked
            var bmp1 = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp1))
            {
                if (VisualStyleRenderer.IsSupported)
                    CheckBoxRenderer.DrawCheckBox(g, new Point(0, 0), CheckBoxState.CheckedNormal);
                else
                    ControlPaint.DrawCheckBox(g, 0, 0, 16, 16, ButtonState.Checked);
            }
            il.Images.Add(bmp1);

            // [2] RadioButton unchecked
            var bmp2 = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp2))
            {
                if (VisualStyleRenderer.IsSupported)
                    RadioButtonRenderer.DrawRadioButton(g, new Point(0, 0), RadioButtonState.UncheckedNormal);
                else
                    ControlPaint.DrawRadioButton(g, 0, 0, 16, 16, ButtonState.Normal);
            }
            il.Images.Add(bmp2);

            // [3] RadioButton checked
            var bmp3 = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp3))
            {
                if (VisualStyleRenderer.IsSupported)
                    RadioButtonRenderer.DrawRadioButton(g, new Point(0, 0), RadioButtonState.CheckedNormal);
                else
                    ControlPaint.DrawRadioButton(g, 0, 0, 16, 16, ButtonState.Checked);
            }
            il.Images.Add(bmp3);

            return il;
        }

        private void RefrescarEstadoNodo(TreeNode n)
        {
            if (n == null || !(n.Tag is AgNodoArbolKey k)) return;
            n.StateImageIndex = k.Nivel == 6 ? (n.Checked ? 3 : 2) : (n.Checked ? 1 : 0);
        }

        private void RefrescarEstadoRecursivo(TreeNode n)
        {
            if (n == null) return;
            RefrescarEstadoNodo(n);
            foreach (TreeNode child in n.Nodes)
            {
                RefrescarEstadoRecursivo(child);
            }
        }

        private void TreeContextoAcademico_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node == null) return;

            if (!(e.Node.Tag is AgNodoArbolKey key) || key.Nivel != 6)
            {
                e.DrawDefault = true;
                return;
            }

            var g = e.Graphics;
            var bounds = e.Node.Bounds;
            int totalWidth = treeContextoAcademico.ClientSize.Width;
            var rowRect = new Rectangle(0, bounds.Y, Math.Max(totalWidth, bounds.Right + 500), bounds.Height);

            // Fondo de selección o fila
            bool isSelected = (e.State & TreeNodeStates.Selected) != 0;
            using (var brush = new SolidBrush(isSelected ? SystemColors.Highlight : treeContextoAcademico.BackColor))
            {
                g.FillRectangle(brush, rowRect);
            }

            // Texto de Modalidad
            Color textColor = isSelected ? SystemColors.HighlightText : treeContextoAcademico.ForeColor;
            using (var textBrush = new SolidBrush(textColor))
            using (var font = new Font(treeContextoAcademico.Font, FontStyle.Regular))
            {
                g.DrawString(e.Node.Text, font, textBrush, bounds.X, bounds.Y + (bounds.Height - font.Height) / 2);
            }

            // Fase 1041: Posiciones calculadas dinámicamente con struct AgColumnasNivel6 (sin offsets mágicos)
            // Fase 663: Radios seleccionados condicionados al estado Checked de la fila (modalidad activa)
            var c = CalcularColumnasNivel6();
            bool filaActiva = e.Node.Checked;
            int baseY = bounds.Y + (bounds.Height - 16) / 2;
            string jNorm = (key.Jornada ?? "").ToUpperInvariant();

            DibujarRadioButtonConTexto(g, bounds.X + c.DxManana, baseY, "Mañana", filaActiva && (jNorm == "MANANA" || jNorm == "MAÑANA"), textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxTarde, baseY, "Tarde", filaActiva && jNorm == "TARDE", textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxNoche, baseY, "Noche", filaActiva && jNorm == "NOCHE", textColor);
            DibujarSeparadorVertical(g, bounds.X + c.DxSep1, bounds);

            // Dimension 2: Días Ventana
            string dv = (key.DiasVentana ?? "").ToUpperInvariant();
            DibujarRadioButtonConTexto(g, bounds.X + c.DxLunJue, baseY, "Lun-Jue", filaActiva && dv == "LUN-JUE", textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxMarVie, baseY, "Mar-Vie", filaActiva && dv == "MAR-VIE", textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxJueSab, baseY, "Jue-Sab", filaActiva && dv == "JUE-SAB", textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxVieSab, baseY, "Vie-Sab", filaActiva && dv == "VIE-SAB", textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxVieDom, baseY, "Vie-Dom", filaActiva && dv == "VIE-DOM", textColor);
            DibujarSeparadorVertical(g, bounds.X + c.DxSep2, bounds);

            // Dimension 3: Horas por día
            int h = key.HorasDia;
            DibujarRadioButtonConTexto(g, bounds.X + c.DxHoras4, baseY, "4h", filaActiva && h == 4, textColor);
            DibujarRadioButtonConTexto(g, bounds.X + c.DxHoras8, baseY, "8h", filaActiva && h == 8, textColor);
        }

        private void DibujarSeparadorVertical(Graphics g, int x, Rectangle bounds)
        {
            using (var pen = new Pen(Color.FromArgb(160, 160, 160)))
            {
                g.DrawLine(pen, x, bounds.Y + 1, x, bounds.Bottom - 2);
            }
        }

        private void DibujarRadioButtonConTexto(Graphics g, int x, int y, string texto, bool activo, Color textColor)
        {
            var state = activo ? RadioButtonState.CheckedNormal : RadioButtonState.UncheckedNormal;
            if (VisualStyleRenderer.IsSupported)
            {
                RadioButtonRenderer.DrawRadioButton(g, new Point(x, y), state);
            }
            else
            {
                ControlPaint.DrawRadioButton(g, x, y, 14, 14, activo ? ButtonState.Checked : ButtonState.Normal);
            }

            using (var brush = new SolidBrush(textColor))
            using (var font = new Font("Segoe UI", 8.25F, FontStyle.Regular))
            {
                g.DrawString(texto, font, brush, x + 16, y + 1);
            }
        }

        private void ProcesarClickDimensionNivel6(TreeNode node, AgNodoArbolKey key, int dx)
        {
            if (node == null || key == null) return;

            // Fase 663: Si la modalidad no está seleccionada, marcarla y aplicar exclusividad entre hermanas
            if (!node.Checked)
            {
                node.Checked = true;
                AplicarExclusividadModalidad(node);
                TreeContextoAcademico_AfterCheck(this, new TreeViewEventArgs(node));
            }

            var c = CalcularColumnasNivel6();

            // Grupo A — Jornada
            if (dx >= c.DxManana && dx < c.DxTarde)
            {
                key.Jornada = "MANANA";
            }
            else if (dx >= c.DxTarde && dx < c.DxNoche)
            {
                key.Jornada = "TARDE";
            }
            else if (dx >= c.DxNoche && dx < c.DxSep1)
            {
                key.Jornada = "NOCHE";
            }
            // Grupo B — Días Ventana
            else if (dx >= c.DxLunJue && dx < c.DxMarVie)
            {
                key.DiasVentana = "LUN-JUE";
            }
            else if (dx >= c.DxMarVie && dx < c.DxJueSab)
            {
                key.DiasVentana = "MAR-VIE";
            }
            else if (dx >= c.DxJueSab && dx < c.DxVieSab)
            {
                key.DiasVentana = "JUE-SAB";
            }
            else if (dx >= c.DxVieSab && dx < c.DxVieDom)
            {
                key.DiasVentana = "VIE-SAB";
            }
            else if (dx >= c.DxVieDom && dx < c.DxSep2)
            {
                key.DiasVentana = "VIE-DOM";
            }
            // Grupo C — Horas Día
            else if (dx >= c.DxHoras4 && dx < c.DxHoras8)
            {
                key.HorasDia = 4;
            }
            else if (dx >= c.DxHoras8 && dx < c.DxFin)
            {
                key.HorasDia = 8;
            }

            // Fase 1193: Sincronizar en BD (ag_aula) si las dimensiones están completas
            if (!string.IsNullOrWhiteSpace(key.Jornada) && !string.IsNullOrWhiteSpace(key.DiasVentana) && key.HorasDia > 0)
            {
                string aulaRef = ConstruirAulaRefDeNodo(node);
                if (!string.IsNullOrWhiteSpace(aulaRef))
                {
                    // Fase 1731: Capturar subunidad activa al registrar edición para que SincronizarConfirmadoAsync no rechace aulas de subunidad no vacía
                    _dimensionesPendientes = (aulaRef, key.Jornada, key.HorasDia, key.DiasVentana, SyncContext.ActiveSubunidad ?? string.Empty);
                    _revisionDimensiones++; // Fase 1728: registrar nueva revisión para evitar descartes concurrentes
                    _timerSyncDimensiones.Stop();
                    _timerSyncDimensiones.Start();
                }
            }

            // Fase 1706 / Fase 1713: Refrescar tooltip de nodo nivel 6
            ActualizarToolTipNivel6(node, key);
        }

        private static readonly HashSet<string> EtiquetasVentanaRadio = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "LUN-JUE", "MAR-VIE", "JUE-SAB", "VIE-SAB", "VIE-DOM"
        };

        private void ActualizarToolTipNivel6(TreeNode node, AgNodoArbolKey key)
        {
            if (node == null)
            {
                return;
            }

            string dv = key?.DiasVentana?.Trim();
            if (!string.IsNullOrWhiteSpace(dv) && !EtiquetasVentanaRadio.Contains(dv))
            {
                node.ToolTipText = $"Ventana actual: {dv} ({AgDiasVentana.AIso(dv)})";
            }
            else
            {
                node.ToolTipText = null;
            }
        }

        private void RecorrerNodosNivel6(TreeNodeCollection nodes, Action<TreeNode, AgNodoArbolKey> accion)
        {
            if (nodes == null)
            {
                return;
            }

            foreach (TreeNode n in nodes)
            {
                if (n.Tag is AgNodoArbolKey k && k.Nivel == 6)
                {
                    accion(n, k);
                }

                RecorrerNodosNivel6(n.Nodes, accion);
            }
        }

        /// <summary>
        /// Fase 1706 / Fase 1713 / Fase 1716: Helper estrecho para recargar DTOs y dimensiones del perfil activo
        /// tras un escalado automático sin reiniciar el contexto ni rebarajar combos del formulario.
        /// Devuelve true si todas las claves requeridas fueron validadas y mapeadas exitosamente; false ante error o inconsistencia.
        /// </summary>
        private async Task<bool> RecargarDimensionesPerfilActivoTrasEscaladoAsync(int idArbolActivo, CancellationToken ct)
        {
            if (idArbolActivo <= 0)
            {
                return false;
            }

            bool prevRestaurando = _restaurandoPerfil;
            _restaurandoPerfil = true;
            try
            {
                var cargaRes = await _arbolService.CargarPerfilAsync(idArbolActivo, ct).ConfigureAwait(true);
                if (cargaRes.IsFailure)
                {
                    LoggerHelper.Warning($"[AgFormGeneradorHorarios] CargarPerfilAsync falló para perfil {idArbolActivo}: {cargaRes.Error}");
                    return false;
                }

                if (cargaRes.Value.Item2 == null)
                {
                    return false;
                }

                var dtosPorAulaRef = cargaRes.Value.Item2
                    .Where(d => !string.IsNullOrWhiteSpace(d.AulaRef))
                    .ToDictionary(d => d.AulaRef, StringComparer.OrdinalIgnoreCase);

                if (dtosPorAulaRef.Count == 0)
                {
                    LoggerHelper.Warning($"[AgFormGeneradorHorarios] Perfil {idArbolActivo} no contiene DTOs con AulaRef.");
                    return false;
                }

                // Fase 1716 & Fase 1721: Validar únicamente sobre nodos seleccionados (Checked) y sobre clave temporal antes de mutar
                var preparadas = new List<Tuple<TreeNode, AgNodoArbolKey, AgNodoArbolKey>>();
                bool falloMapeo = false;

                RecorrerNodosNivel6(treeContextoAcademico.Nodes, (nodo, key) =>
                {
                    if (!nodo.Checked)
                    {
                        return;
                    }

                    string aulaRef = ConstruirAulaRefDeNodo(nodo);
                    if (string.IsNullOrWhiteSpace(aulaRef))
                    {
                        return;
                    }

                    if (!dtosPorAulaRef.TryGetValue(aulaRef, out var dto))
                    {
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] DTO faltante para aulaRef '{aulaRef}' seleccionada en perfil {idArbolActivo}.");
                        falloMapeo = true;
                        return;
                    }

                    var temporal = new AgNodoArbolKey(6, 0, "", "");
                    var validacion = AgNodoArbolKeyMapper.AplicarDimensiones(temporal, dto);
                    if (validacion.IsFailure)
                    {
                        LoggerHelper.Warning($"[AgFormGeneradorHorarios] Dimensiones inválidas para aulaRef '{aulaRef}': {validacion.Error}");
                        falloMapeo = true;
                        return;
                    }

                    preparadas.Add(Tuple.Create(nodo, key, temporal));
                });

                if (falloMapeo || preparadas.Count == 0)
                {
                    return false;
                }

                // Aplicar las dimensiones solo si la validación previa de todas las seleccionadas fue completamente exitosa
                foreach (var p in preparadas)
                {
                    p.Item2.Jornada = p.Item3.Jornada;
                    p.Item2.HorasDia = p.Item3.HorasDia;
                    p.Item2.DiasVentana = p.Item3.DiasVentana;

                    ActualizarToolTipNivel6(p.Item1, p.Item2);
                }

                treeContextoAcademico.Invalidate();
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al recargar dimensiones tras escalado", ex);
                return false;
            }
            finally
            {
                _restaurandoPerfil = prevRestaurando;
            }
        }

        /// <summary>
        /// Resuelve la clave canónica natural AulaRef de un nodo de Nivel 6 recorriendo sus ancestros (Fase 1193).
        /// </summary>
        private string ConstruirAulaRefDeNodo(TreeNode n6)
        {
            if (n6 == null || !(n6.Tag is AgNodoArbolKey k6) || k6.Nivel != 6) return null;
            var n5 = n6.Parent;
            if (n5 == null || !(n5.Tag is AgNodoArbolKey k5) || k5.Nivel != 5) return null;
            var n4 = n5.Parent;
            if (n4 == null || !(n4.Tag is AgNodoArbolKey k4) || k4.Nivel != 4) return null;
            var n3 = n4.Parent;
            if (n3 == null || !(n3.Tag is AgNodoArbolKey k3) || k3.Nivel != 3) return null;
            var n2 = n3.Parent;
            if (n2 == null || !(n2.Tag is AgNodoArbolKey k2) || k2.Nivel != 2) return null;
            var n1 = n2.Parent;
            if (n1 == null || !(n1.Tag is AgNodoArbolKey k1) || k1.Nivel != 1) return null;

            return AgAulaFactory.ConstruirAulaRef(k1.Id, k2.Id, k3.Id, k4.Codigo, k5.Codigo, k6.Codigo);
        }

        /// <summary>
        /// Fases 661/662, 1041, 1047-1049: Actualiza la visibilidad y alineación milimétrica de las cabeceras
        /// del Nivel 6 en pnlCabecerasNivel6 con protección anti-reentrada y compensación de borde.
        /// </summary>
        private void ActualizarCabecerasNivel6()
        {
            if (_alineandoCabeceras) return;
            _alineandoCabeceras = true;
            try
            {
                TreeNode n6 = BuscarPrimerNodoNivel6(treeContextoAcademico.Nodes);
                bool hayVisible = n6 != null && n6.IsVisible;
                if (pnlCabecerasNivel6.Visible != hayVisible)
                {
                    pnlCabecerasNivel6.Visible = hayVisible;
                }
                if (!hayVisible) return;

                var c = CalcularColumnasNivel6(n6);
                const int BordeArbol = 1; // Compensación BorderStyle.FixedSingle (Fase 1049)
                lblColJornada.Left = c.XBase + c.DxManana - BordeArbol;
                sepNivel6A.Left = c.XBase + c.DxSep1 - BordeArbol;
                lblColDias.Left = c.XBase + c.DxLunJue - BordeArbol;
                sepNivel6B.Left = c.XBase + c.DxSep2 - BordeArbol;
                lblColHorasDia.Left = c.XBase + c.DxHoras4 - BordeArbol;
            }
            finally
            {
                _alineandoCabeceras = false;
            }
        }

        private bool HayNodoNivel6Visible()
        {
            TreeNode n = BuscarPrimerNodoNivel6(treeContextoAcademico.Nodes);
            return n != null && n.IsVisible;
        }

        private int ObtenerXBaseNivel6(TreeNode n6 = null)
        {
            TreeNode n = n6 ?? BuscarPrimerNodoNivel6(treeContextoAcademico.Nodes);
            int minPorDpi = (int)(treeContextoAcademico.ClientSize.Width * 0.09);
            if (n != null && n.IsVisible)
            {
                return Math.Max(n.Bounds.X, minPorDpi);
            }
            // Estimación determinista basada en 6 niveles de indentación (14 px) + espacio de icono de estado (~19 px)
            int est = (6 * treeContextoAcademico.Indent) + 19;
            return Math.Max(est, minPorDpi);
        }

        private TreeNode BuscarPrimerNodoNivel6(TreeNodeCollection nodes)
        {
            if (nodes == null) return null;
            foreach (TreeNode n in nodes)
            {
                if (n.Tag is AgNodoArbolKey k && k.Nivel == 6) return n;
                var sub = BuscarPrimerNodoNivel6(n.Nodes);
                if (sub != null) return sub;
            }
            return null;
        }

        #endregion

        private const string ModalidadPorDefecto = ModalidadTutoriaCatalog.Presencial; // Fase 528 y 789

        /// <summary>
        /// Aplica exclusividad (comportamiento tipo radio) entre hermanos de nivel 6 (Modalidad).
        /// </summary>
        private void AplicarExclusividadModalidad(TreeNode nodo)
        {
            if (nodo?.Parent == null) return;
            foreach (TreeNode hno in nodo.Parent.Nodes)
            {
                if (hno != nodo && hno.Tag is AgNodoArbolKey k && k.Nivel == 6)
                {
                    hno.Checked = false;
                    RefrescarEstadoNodo(hno); // Fase 542
                }
            }
            RefrescarEstadoNodo(nodo); // Fase 542
        }

        private void TreeContextoAcademico_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_isUpdatingTreeChecks) return;
            _isUpdatingTreeChecks = true;
            try
            {
                // Fase 528: exclusividad entre hermanos de nivel 6
                if (e.Node.Tag is AgNodoArbolKey k && k.Nivel == 6 && e.Node.Checked)
                {
                    AplicarExclusividadModalidad(e.Node);
                    // Fase 663: Invalidar filas de hermanas para actualizar render de radios
                    if (e.Node.Parent != null)
                    {
                        foreach (TreeNode hno in e.Node.Parent.Nodes)
                        {
                            var rh = hno.Bounds;
                            treeContextoAcademico.Invalidate(new Rectangle(0, rh.Y, treeContextoAcademico.ClientSize.Width, rh.Height));
                        }
                    }
                }

                CargarHijosLazy(e.Node);
                PropagarCheckHaciaAbajo(e.Node, e.Node.Checked);
                ActualizarEstadoPadre(e.Node);
                RefrescarEstadoNodo(e.Node); // Fase 542
            }
            finally
            {
                _isUpdatingTreeChecks = false;
                ActualizarConteoCombinaciones();
            }
        }

        private void PropagarCheckHaciaAbajo(TreeNode node, bool isChecked)
        {
            // Fase 528: si el nodo actual es de nivel 5 (Grupo), en sus hijos marcar SOLO la modalidad por defecto
            if (node.Tag is AgNodoArbolKey kn && kn.Nivel == 5 && node.Nodes.Count > 0)
            {
                var modDef = node.Nodes.Cast<TreeNode>().FirstOrDefault(n =>
                    n.Tag is AgNodoArbolKey k6 && k6.Nivel == 6 &&
                    string.Equals(k6.Codigo, ModalidadPorDefecto, StringComparison.OrdinalIgnoreCase))
                    ?? node.Nodes[0];

                foreach (TreeNode m in node.Nodes)
                {
                    m.Checked = isChecked && (m == modDef);
                    RefrescarEstadoNodo(m); // Fase 542
                }
                RefrescarEstadoNodo(node); // Fase 542
                return;
            }

            foreach (TreeNode child in node.Nodes)
            {
                CargarHijosLazy(child);

                // Fase 528: en nivel 5 (Grupo) marcar SOLO la modalidad por defecto
                if (child.Tag is AgNodoArbolKey kk && kk.Nivel == 5 && child.Nodes.Count > 0)
                {
                    var modDef = child.Nodes.Cast<TreeNode>().FirstOrDefault(n =>
                        n.Tag is AgNodoArbolKey k6 && k6.Nivel == 6 &&
                        string.Equals(k6.Codigo, ModalidadPorDefecto, StringComparison.OrdinalIgnoreCase))
                        ?? child.Nodes[0];

                    child.Checked = isChecked;
                    RefrescarEstadoNodo(child); // Fase 542
                    foreach (TreeNode m in child.Nodes)
                    {
                        m.Checked = isChecked && (m == modDef);
                        RefrescarEstadoNodo(m); // Fase 542
                    }
                    continue;
                }

                child.Checked = isChecked;
                RefrescarEstadoNodo(child); // Fase 542
                PropagarCheckHaciaAbajo(child, isChecked);
            }
            RefrescarEstadoNodo(node); // Fase 542
        }

        private void ActualizarEstadoPadre(TreeNode node)
        {
            var parent = node.Parent;
            if (parent == null) return;

            bool anyChecked = false;
            foreach (TreeNode sibling in parent.Nodes)
            {
                if (sibling.Checked) { anyChecked = true; break; }
            }

            if (parent.Checked != anyChecked)
            {
                parent.Checked = anyChecked;
            }

            RefrescarEstadoNodo(parent); // Fase 542
            ActualizarEstadoPadre(parent);
        }

        private List<AgContextoSeleccionDto> ObtenerCombinacionesSeleccionadas(List<string> descartes = null)
        {
            var resultado = new List<AgContextoSeleccionDto>();

            void Recorrer(TreeNode node, int idPeriodo, int idSede, int idPrograma, string semestre, string grupo)
            {
                if (node.Tag == null) return; // Omitir dummy nodes "Cargando..."

                if (node.Tag is AgNodoArbolKey key)
                {
                    switch (key.Nivel)
                    {
                        case 1: idPeriodo = key.Id; break;
                        case 2: idSede = key.Id; break;
                        case 3: idPrograma = key.Id; break;
                        case 4: semestre = key.Codigo; break;
                        case 5: grupo = key.Codigo; break;
                        case 6:
                            // Fase 215, Ronda 206 & Fase 1081: Validar combinaciones completas; acumular descartes para feedback visual
                            if (node.Checked && node.Nodes.Count == 0)
                            {
                                if (idPeriodo > 0 && idSede > 0 && idPrograma > 0 &&
                                    !string.IsNullOrWhiteSpace(semestre) &&
                                    !string.IsNullOrWhiteSpace(grupo) &&
                                    !string.IsNullOrWhiteSpace(key.Codigo))
                                {
                                    var resDto = AgSeleccionMapper.DesdeNodo(key, idPeriodo, idSede, idPrograma, semestre, grupo);
                                    if (resDto.IsSuccess)
                                    {
                                        var dto = resDto.Value;
                                        // Fase 250, 260 & 789: Marcar EsVirtual = 1 y Capacidad = 999 explícitamente si la modalidad es Plataforma
                                        if (ModalidadTutoriaCatalog.EsPlataforma(key.Codigo))
                                        {
                                            dto.EsVirtual = 1;
                                            dto.Capacidad = 999;
                                        }
                                        resultado.Add(dto);
                                    }
                                    else
                                    {
                                        if (descartes != null)
                                        {
                                            descartes.Add($"• {node.FullPath}: {resDto.Error}");
                                        }
                                        LoggerHelper.WarningLimitado(
                                            $"SeleccionInvalida|{key.Codigo}|{semestre}|{grupo}",
                                            $"[AgFormGeneradorHorarios] Combinación ignorada por mapper: {resDto.Error}");
                                    }
                                }
                            }
                            return;
                    }
                }

                foreach (TreeNode child in node.Nodes)
                {
                    Recorrer(child, idPeriodo, idSede, idPrograma, semestre, grupo);
                }
            }

            foreach (TreeNode root in treeContextoAcademico.Nodes)
            {
                Recorrer(root, 0, 0, 0, string.Empty, string.Empty);
            }

            // Fase 445, Fase 538 & Fase 545: Preservar IdPeriodo del árbol; solo asignar el del combo como fallback si la hoja carece de período
            int idPeriodoObjetivo = (cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo))?.Id ?? 0;
            var periodosEnHojas = resultado.Select(r => r.IdPeriodo).Where(p => p > 0).Distinct().ToList();
            if (periodosEnHojas.Count > 1)
            {
                LoggerHelper.Warning($"[AgFormGeneradorHorarios] Combinaciones seleccionadas pertenecen a múltiples períodos: [{string.Join(",", periodosEnHojas)}]. Se preservan períodos individuales.");
            }
            else
            {
                foreach (var dto in resultado)
                {
                    if (dto.IdPeriodo <= 0 && idPeriodoObjetivo > 0)
                    {
                        dto.IdPeriodo = idPeriodoObjetivo;
                    }
                }
            }

            // Deduplicar por clave canónica AulaRef
            return resultado
                .GroupBy(c => c.AulaRef)
                .Select(g => g.First())
                .ToList();
        }

        private void ActualizarConteoCombinaciones()
        {
            var combinaciones = ObtenerCombinacionesSeleccionadas();
            int count = combinaciones.Select(c => c.AulaRef).Distinct().Count();

            if (count == 0)
            {
                lblConteoCombinaciones.Text = "0 combinaciones seleccionadas";
                lblConteoCombinaciones.ForeColor = Color.DimGray;
                toolTipConteo?.SetToolTip(lblConteoCombinaciones, "0 combinaciones seleccionadas en el árbol.");
                btnGenerarAulas.Text = "Generar Aulas (0)";
                btnGenerarAulas.Enabled = false;
            }
            else if (count <= 5000)
            {
                lblConteoCombinaciones.Text = $"{count} combinaciones seleccionadas";
                lblConteoCombinaciones.ForeColor = Color.DarkGreen;
                toolTipConteo?.SetToolTip(lblConteoCombinaciones, $"{count} combinaciones académicas seleccionadas listas para generar aulas.");
                btnGenerarAulas.Text = $"Generar Aulas ({count})";
                btnGenerarAulas.Enabled = true;
            }
            else
            {
                lblConteoCombinaciones.Text = $"{count} combinaciones (¡ALERTA! Supera el límite de 5000)";
                lblConteoCombinaciones.ForeColor = Color.Red;
                toolTipConteo?.SetToolTip(lblConteoCombinaciones, $"{count} combinaciones seleccionadas. ¡ALERTA! Supera el límite máximo permitido de 5000.");
                btnGenerarAulas.Text = $"Límite Excedido ({count})";
                btnGenerarAulas.Enabled = false;
            }
        }

        private void ExpandirTodoElArbol()
        {
            treeContextoAcademico.BeginUpdate();
            try
            {
                foreach (TreeNode root in treeContextoAcademico.Nodes)
                {
                    CargarTodoRecursivo(root);
                    root.ExpandAll();
                }
            }
            finally
            {
                treeContextoAcademico.EndUpdate();
                ActualizarCabecerasNivel6();
            }
        }

        private void ContraerTodoElArbol()
        {
            treeContextoAcademico.CollapseAll();
            if (treeContextoAcademico.Nodes.Count > 0)
            {
                treeContextoAcademico.Nodes[0].Expand();
            }
        }

        private void MarcarTodoElArbol(bool check)
        {
            _isUpdatingTreeChecks = true;
            treeContextoAcademico.BeginUpdate();
            try
            {
                foreach (TreeNode root in treeContextoAcademico.Nodes)
                {
                    if (check) CargarTodoRecursivo(root);
                    root.Checked = check;
                    RefrescarEstadoNodo(root); // Fase 542
                    PropagarCheckHaciaAbajo(root, check);
                    RefrescarEstadoRecursivo(root); // Fase 542
                }
            }
            finally
            {
                treeContextoAcademico.EndUpdate();
                _isUpdatingTreeChecks = false;
                ActualizarConteoCombinaciones();
            }
        }

        private void InvertirSeleccionArbol()
        {
            _isUpdatingTreeChecks = true;
            treeContextoAcademico.BeginUpdate();
            try
            {
                void Invertir(TreeNode node)
                {
                    CargarHijosLazy(node);
                    if (node.Nodes.Count == 0 && node.Tag is AgNodoArbolKey k && k.Nivel == 6)
                    {
                        node.Checked = !node.Checked;
                        if (node.Checked)
                        {
                            AplicarExclusividadModalidad(node); // Fase 528
                        }
                        RefrescarEstadoNodo(node); // Fase 542
                    }
                    else
                    {
                        foreach (TreeNode child in node.Nodes)
                        {
                            Invertir(child);
                        }
                    }
                }

                foreach (TreeNode root in treeContextoAcademico.Nodes)
                {
                    Invertir(root);
                    ActualizarEstadoPadre(root);
                    RefrescarEstadoRecursivo(root); // Fase 542
                }
            }
            finally
            {
                treeContextoAcademico.EndUpdate();
                _isUpdatingTreeChecks = false;
                ActualizarConteoCombinaciones();
            }
        }

        #endregion

        #region Gestión de Perfiles de Selección de Árbol (Fase 240, 241)

        private async Task RefrescarComboArbolesAsync()
        {
            try
            {
                string sub = SyncContext.ActiveSubunidad ?? string.Empty;
                var res = await _arbolService.ListarPerfilesAsync(sub, CancellationToken.None);
                if (res.IsSuccess)
                {
                    _perfilesArbol = res.Value.ToList();
                    cmbArboles.DataSource = null;
                    cmbArboles.DataSource = _perfilesArbol;
                    cmbArboles.DisplayMember = "Nombre";
                    cmbArboles.ValueMember = "Id";
                    if (_perfilesArbol.Count > 0) cmbArboles.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al listar perfiles de árbol", ex);
            }
        }

        private async Task GuardarNuevoPerfilArbolAsync()
        {
            int idPeriodo = (cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo))?.Id ?? 0;
            // Fase 252: Validación temprana de período y ventana temporal
            if (_ventanaResuelta == null || idPeriodo <= 0)
            {
                MessageBox.Show(this, "Debe seleccionar un período con ventana temporal resuelta antes de guardar el perfil.", "Guardar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 1207: Obtener combinaciones seleccionadas evaluando descartes para evitar pérdidas silenciosas
            var descartes = new List<string>();
            var combinaciones = ObtenerCombinacionesSeleccionadas(descartes);
            if (descartes.Count > 0)
            {
                string detalle = string.Join("\n", descartes.Take(10));
                if (descartes.Count > 10) detalle += $"\n... y {descartes.Count - 10} más.";
                var confirmDescartes = MessageBox.Show(this,
                    $"Se encontraron {descartes.Count} combinaciones seleccionadas con configuración incompleta o inválida que serán omitidas:\n\n{detalle}\n\n¿Desea continuar guardando únicamente las {combinaciones.Count} combinaciones válidas?",
                    "Combinaciones Incompletas Descartadas",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (confirmDescartes != DialogResult.Yes) return;
            }

            if (combinaciones.Count == 0)
            {
                MessageBox.Show(this, "Debe seleccionar al menos una combinación (hoja) en el árbol académico antes de guardar un perfil.", "Guardar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 246: Sincronizar fechas manuales con _ventanaResuelta antes de guardar
            if (chkManualVentana.Checked)
            {
                dtpVentanaInicio.Value = dtpVentanaInicio.Value.Date;
                dtpVentanaFin.Value = dtpVentanaFin.Value.Date;
                await ActualizarVentanaSemestralAsync();
            }

            using (var dlg = new FormNombreArbol())
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                string nombre = dlg.NombrePerfil;

                // Fase 258: Prevenir pérdida silenciosa de perfil homónimo
                bool existe = _perfilesArbol.Any(p =>
                    string.Equals(p.Nombre?.Trim(), nombre?.Trim(), StringComparison.OrdinalIgnoreCase));
                if (existe)
                {
                    var confirmacion = MessageBox.Show(this,
                        $"Ya existe un perfil con el nombre '{nombre}'. ¿Desea sobrescribirlo con la selección y parámetros actuales?",
                        "Perfil Duplicado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirmacion != DialogResult.Yes)
                    {
                        return;
                    }
                }

                string sub = SyncContext.ActiveSubunidad ?? string.Empty;

                var res = await _arbolService.GuardarPerfilAsync(
                    nombre,
                    string.Empty,
                    idPeriodo,
                    _ventanaResuelta,
                    chkManualVentana.Checked,
                    (int)numPoblacion.Value,
                    (int)numGeneraciones.Value,
                    (double)numTasaCruce.Value,
                    (double)numTasaMutacion.Value,
                    (int)numElitismo.Value,
                    combinaciones,
                    Environment.UserName,
                    sub,
                    CancellationToken.None);

                if (res.IsSuccess)
                {
                    await RefrescarComboArbolesAsync();
                    SeleccionarPerfilEnCombo(res.Value);
                    MessageBox.Show(this, $"Perfil '{nombre}' guardado exitosamente con {combinaciones.Count} combinación(es).", "Perfil Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, res.Error, "Error al Guardar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task SobrescribirPerfilArbolAsync()
        {
            // Fase 1729: Inhibir acción si el diagnóstico pre-vuelo está en curso
            if (_diagnosticoEnCurso)
            {
                return;
            }

            if (!(cmbArboles.SelectedItem is AgArbol perfilSel))
            {
                MessageBox.Show(this, "Seleccione un perfil existente de la lista para sobrescribir.", "Sobrescribir Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int idPeriodo = (cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo))?.Id ?? perfilSel.IdPeriodo;
            // Fase 252: Validación temprana de período y ventana temporal
            if (_ventanaResuelta == null || idPeriodo <= 0)
            {
                MessageBox.Show(this, "Debe seleccionar un período con ventana temporal resuelta antes de sobrescribir el perfil.", "Sobrescribir Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 1207: Obtener combinaciones seleccionadas evaluando descartes para evitar pérdidas silenciosas
            var descartes = new List<string>();
            var combinaciones = ObtenerCombinacionesSeleccionadas(descartes);
            if (descartes.Count > 0)
            {
                string detalle = string.Join("\n", descartes.Take(10));
                if (descartes.Count > 10) detalle += $"\n... y {descartes.Count - 10} más.";
                var confirmDescartes = MessageBox.Show(this,
                    $"Se encontraron {descartes.Count} combinaciones seleccionadas con configuración incompleta o inválida que serán omitidas:\n\n{detalle}\n\n¿Desea continuar sobrescribiendo con únicamente las {combinaciones.Count} combinaciones válidas?",
                    "Combinaciones Incompletas Descartadas",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (confirmDescartes != DialogResult.Yes) return;
            }

            if (combinaciones.Count == 0)
            {
                MessageBox.Show(this, "Debe seleccionar al menos una combinación en el árbol académico antes de sobrescribir el perfil.", "Sobrescribir Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(this,
                $"¿Desea sobrescribir el perfil '{perfilSel.Nombre}' con la selección ({combinaciones.Count} combinaciones) y parámetros evolutivos actuales?",
                "Confirmar Sobrescritura", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            // Fase 246: Sincronizar fechas manuales con _ventanaResuelta antes de guardar
            if (chkManualVentana.Checked)
            {
                dtpVentanaInicio.Value = dtpVentanaInicio.Value.Date;
                dtpVentanaFin.Value = dtpVentanaFin.Value.Date;
                await ActualizarVentanaSemestralAsync();
            }

            string sub = SyncContext.ActiveSubunidad ?? string.Empty;

            var res = await _arbolService.GuardarPerfilAsync(
                perfilSel.Nombre,
                perfilSel.Descripcion,
                idPeriodo,
                _ventanaResuelta,
                chkManualVentana.Checked,
                (int)numPoblacion.Value,
                (int)numGeneraciones.Value,
                (double)numTasaCruce.Value,
                (double)numTasaMutacion.Value,
                (int)numElitismo.Value,
                combinaciones,
                Environment.UserName,
                sub,
                CancellationToken.None);

            if (res.IsSuccess)
            {
                await RefrescarComboArbolesAsync();
                SeleccionarPerfilEnCombo(res.Value);
                MessageBox.Show(this, $"Perfil '{perfilSel.Nombre}' sobrescrito exitosamente.", "Sobrescribir Perfil", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, res.Error, "Error al Sobrescribir Perfil", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task CargarPerfilArbolSeleccionadoAsync(bool mostrarAvisos = true)
        {
            // Fase 1729: Inhibir acción si el diagnóstico pre-vuelo está en curso
            if (_diagnosticoEnCurso)
            {
                return;
            }

            if (!(cmbArboles.SelectedItem is AgArbol perfilSel))
            {
                if (mostrarAvisos)
                {
                    MessageBox.Show(this, "Seleccione un perfil de la lista para cargar.", "Cargar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }

            var traza = CrearTraza("CargarPerfilArbol");
            traza.Inicio($"Perfil={perfilSel.Nombre}");

            try
            {
                var cargaRes = await _arbolService.CargarPerfilAsync(perfilSel.Id, CancellationToken.None);
                if (cargaRes.IsFailure)
                {
                    traza.Fallo("P12_CargarPerfil", new Exception(cargaRes.Error), AgCausaFallo.SinAulas);
                    traza.Fin(false, cargaRes.Error);
                    if (mostrarAvisos)
                    {
                        MessageBox.Show(this, cargaRes.Error, "Error al Cargar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                var perfil = cargaRes.Value.Item1;
                var dtos = cargaRes.Value.Item2;

                // Fase 253: Silenciar eventos durante la asignación masiva de controles para evitar carreras concurrentes
                _restaurandoPerfil = true;
                try
                {
                    // 1. Establecer Período si coincide
                    if (perfil.IdPeriodo > 0)
                    {
                        var perEncontrado = _periodos.FirstOrDefault(p => p.Id == perfil.IdPeriodo);
                        if (perEncontrado != null)
                        {
                            cmbPeriodos.Marcar(new[] { perEncontrado.Id });
                        }
                    }

                    // Fase 1580: Alinear el alcance del árbol (períodos, sedes, programas y tipo) con TODAS
                    // las combinaciones del perfil. Antes solo se sincronizaba el primer DTO, por lo que las
                    // combinaciones de otras sedes/programas activas se reportaban como "omitidas".
                    var alcancePerfil = AgPerfilAlcanceResolver.Resolver(dtos, _periodos, _sedes, _programas);

                    if (alcancePerfil.Periodos.Count > 0)
                    {
                        cmbPeriodos.Marcar(alcancePerfil.Periodos);
                    }

                    if (alcancePerfil.Sedes.Count > 0)
                    {
                        cmbSedes.Marcar(alcancePerfil.Sedes);
                    }

                    // El tipo se ajusta al mínimo que cubre los programas del perfil (TODOS si son mixtos)
                    bool suprimirPrevio = _suprimirAutoFiltro;
                    _suprimirAutoFiltro = true;
                    try
                    {
                        int idxTipo = cmbTipoPrograma.FindStringExact(alcancePerfil.TipoPrograma);
                        if (idxTipo >= 0 && idxTipo != cmbTipoPrograma.SelectedIndex)
                        {
                            cmbTipoPrograma.SelectedIndex = idxTipo;
                        }
                    }
                    finally
                    {
                        _suprimirAutoFiltro = suprimirPrevio;
                    }

                    // Recarga la lista de programas visibles según el tipo y marca los del perfil
                    if (alcancePerfil.Programas.Count > 0)
                    {
                        CargarProgramasSegunTipo(alcancePerfil.TipoPrograma, alcancePerfil.Programas);
                    }

                    if (alcancePerfil.CombosFueraDeCatalogo > 0)
                    {
                        LoggerHelper.Warning(
                            $"[AgFormGeneradorHorarios] Perfil '{perfilSel.Nombre}': {alcancePerfil.CombosFueraDeCatalogo} combinación(es) con período/sede/programa inexistente en catálogo.");
                    }

                    // 2. Establecer ventana temporal
                    if (perfil.ManualVentana == 1)
                    {
                        chkManualVentana.Checked = true;
                        if (DateTime.TryParse(perfil.FechaInicio, out DateTime fIni)) dtpVentanaInicio.Value = fIni;
                        if (DateTime.TryParse(perfil.FechaFin, out DateTime fFin)) dtpVentanaFin.Value = fFin;
                    }
                    else
                    {
                        chkManualVentana.Checked = false;
                    }
                }
                finally
                {
                    _restaurandoPerfil = false;
                }

                // Fase 245 & 253: Actualizar _ventanaResuelta una única vez y en orden determinista tras fijar todos los controles
                await ActualizarVentanaSemestralAsync();

                // Fase 1580: Repoblar el árbol con el alcance del perfil antes de restaurar las hojas,
                // evitando que combinaciones de sedes/programas activos queden fuera del árbol por el filtro previo.
                PoblarArbolContextoAcademico();
                _filtroArbolDesactualizado = false;
                ActualizarResumenFiltros();
                await GuardarFiltrosAsync(forzar: true);

                // 3. Establecer Parámetros Evolutivos
                if (perfil.Poblacion >= numPoblacion.Minimum && perfil.Poblacion <= numPoblacion.Maximum)
                    numPoblacion.Value = perfil.Poblacion;
                if (perfil.Generaciones >= numGeneraciones.Minimum && perfil.Generaciones <= numGeneraciones.Maximum)
                    numGeneraciones.Value = perfil.Generaciones;
                if ((decimal)perfil.TasaCruce >= numTasaCruce.Minimum && (decimal)perfil.TasaCruce <= numTasaCruce.Maximum)
                    numTasaCruce.Value = (decimal)perfil.TasaCruce;
                if ((decimal)perfil.TasaMutacion >= numTasaMutacion.Minimum && (decimal)perfil.TasaMutacion <= numTasaMutacion.Maximum)
                    numTasaMutacion.Value = (decimal)perfil.TasaMutacion;
                if (perfil.Elitismo >= numElitismo.Minimum && perfil.Elitismo <= numElitismo.Maximum)
                    numElitismo.Value = perfil.Elitismo;

                // 4. Asegurar selección en el árbol y registrar traza P12 (Fases 1209, 1217 y 1220)
                traza.Paso("P12_RestaurarDimensionesPerfil", $"Restauración de 3 dimensiones en {dtos.Count} combinaciones");
                traza.Dato("Combinaciones", dtos.Count);

                var resumenDims = string.Join("; ", dtos.Take(20).Select(d => $"{d.AulaRef}={d.Jornada}/{d.HorasDia}h/{d.DiasVentana}"));
                traza.Dato("Dimensiones", resumenDims);

                int dimsDefecto = AsegurarSubarbol(dtos, mostrarAvisos, out var omitidas);

                if (omitidas.Count > 0)
                {
                    LoggerHelper.Warning(
                        $"[AgFormGeneradorHorarios] Perfil '{perfil.Nombre}': {omitidas.Count} combinación(es) omitida(s) por período/sede/programa inexistente en catálogo: {string.Join(", ", omitidas)}");
                }

                if (dimsDefecto > 0)
                {
                    traza.Fallo("P12_RestaurarDimensionesPerfil",
                                new Exception($"Fallback aplicado en {dimsDefecto} combinación(es)"),
                                AgCausaFallo.HorasDiaInvalidoNivel);
                }

                traza.Fin(true, $"Perfil '{perfil.Nombre}' restaurado ({dtos.Count} combinación/es, fallback={dimsDefecto})");
                LoggerHelper.Info($"[AgFormGeneradorHorarios] Perfil '{perfil.Nombre}' cargado exitosamente ({dtos.Count} combinación/es, {dimsDefecto} con fallback). Dims: {resumenDims}");
            }
            catch (Exception ex)
            {
                traza.Fallo("P12_RestaurarDimensionesPerfil", ex, AgCausaFallo.ExcepcionNoControlada);
                traza.Fin(false, ex.Message);
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error al cargar perfil de árbol", ex);
                if (mostrarAvisos)
                {
                    MessageBox.Show(this, $"Error al cargar perfil: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task EliminarPerfilArbolSeleccionadoAsync()
        {
            if (!(cmbArboles.SelectedItem is AgArbol perfilSel))
            {
                MessageBox.Show(this, "Seleccione un perfil para eliminar.", "Eliminar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(this,
                $"¿Está seguro de que desea eliminar permanentemente el perfil '{perfilSel.Nombre}'?",
                "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            var res = await _arbolService.EliminarPerfilAsync(perfilSel.Id, CancellationToken.None);
            if (res.IsSuccess)
            {
                await RefrescarComboArbolesAsync();
                MessageBox.Show(this, $"Perfil '{perfilSel.Nombre}' eliminado exitosamente.", "Perfil Eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, res.Error, "Error al Eliminar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SeleccionarPerfilEnCombo(long idArbol)
        {
            for (int i = 0; i < cmbArboles.Items.Count; i++)
            {
                if (cmbArboles.Items[i] is AgArbol p && p.Id == idArbol)
                {
                    cmbArboles.SelectedIndex = i;
                    break;
                }
            }
        }

        private int AsegurarSubarbol(IReadOnlyCollection<AgContextoSeleccionDto> nodos, bool mostrarAvisos, out List<string> omitidas)
        {
            omitidas = new List<string>();
            if (nodos == null || nodos.Count == 0) return 0;

            _isUpdatingTreeChecks = true;
            treeContextoAcademico.BeginUpdate();
            int huerfanos = 0;
            int dimensionesDefecto = 0;
            TreeNode primerNodoRestaurado = null;

            string RefDe(AgContextoSeleccionDto d) => !string.IsNullOrWhiteSpace(d.AulaRef)
                ? d.AulaRef
                : $"{d.IdPeriodo}|{d.IdSede}|{d.IdPrograma}|{d.Semestre}|{d.Grupo}|{d.ModalidadEfectiva}";

            try
            {
                // 1. Limpiar todas las marcas actuales del árbol
                foreach (TreeNode root in treeContextoAcademico.Nodes)
                {
                    root.Checked = false;
                    DesmarcarRecursivo(root);
                }

                // 2. Para cada combinación del perfil, ubicar o asegurar la ruta jerárquica hasta la hoja
                foreach (var dto in nodos)
                {
                    // Nivel 1: Periodo
                    TreeNode nodoPer = treeContextoAcademico.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 1 && k.Id == dto.IdPeriodo);

                    if (nodoPer == null)
                    {
                        huerfanos++;
                        omitidas.Add(RefDe(dto));
                        continue;
                    }

                    // Nivel 2: Sede
                    TreeNode nodoSede = nodoPer.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 2 && k.Id == dto.IdSede);

                    if (nodoSede == null)
                    {
                        huerfanos++;
                        omitidas.Add(RefDe(dto));
                        continue;
                    }

                    // Nivel 3: Programa
                    TreeNode nodoProg = nodoSede.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 3 && k.Id == dto.IdPrograma);

                    if (nodoProg == null)
                    {
                        huerfanos++;
                        omitidas.Add(RefDe(dto));
                        continue;
                    }

                    CargarHijosLazy(nodoProg);

                    // Nivel 4: Semestre
                    TreeNode nodoSem = nodoProg.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 4 && string.Equals(k.Codigo, dto.Semestre, StringComparison.OrdinalIgnoreCase));

                    if (nodoSem == null)
                    {
                        nodoSem = new TreeNode($"Semestre {dto.Semestre}")
                        {
                            Tag = new AgNodoArbolKey(4, 0, dto.Semestre, dto.Semestre)
                        };
                        // Fase 251: Agregar nodo ficticio para que CargarHijosLazy reconozca el nodo como no cargado
                        nodoSem.Nodes.Add(new TreeNode("Cargando...") { Tag = null });
                        nodoProg.Nodes.Add(nodoSem);
                    }

                    CargarHijosLazy(nodoSem);

                    // Nivel 5: Grupo
                    TreeNode nodoGrp = nodoSem.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 5 && string.Equals(k.Codigo, dto.Grupo, StringComparison.OrdinalIgnoreCase));

                    if (nodoGrp == null)
                    {
                        nodoGrp = new TreeNode($"Grupo {dto.Grupo}")
                        {
                            Tag = new AgNodoArbolKey(5, 0, dto.Grupo, dto.Grupo)
                        };
                        // Fase 251: Agregar nodo ficticio para que CargarHijosLazy reconozca el nodo como no cargado
                        nodoGrp.Nodes.Add(new TreeNode("Cargando...") { Tag = null });
                        nodoSem.Nodes.Add(nodoGrp);
                    }

                    CargarHijosLazy(nodoGrp);

                    // Nivel 6: Modalidad
                    string modObjetivo = dto.ModalidadEfectiva;
                    TreeNode nodoMod = nodoGrp.Nodes.Cast<TreeNode>()
                        .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 6 && string.Equals(ModalidadTutoriaCatalog.NormalizarTutoria(k.Codigo), modObjetivo, StringComparison.OrdinalIgnoreCase));

                    if (nodoMod == null)
                    {
                        nodoMod = new TreeNode(ModalidadTutoriaCatalog.AEtiqueta(modObjetivo))
                        {
                            Tag = new AgNodoArbolKey(6, 0, modObjetivo, modObjetivo)
                        };
                        nodoGrp.Nodes.Add(nodoMod);
                    }

                    // Fase 1205: Restaurar dimensiones completas (Jornada, HorasDia, DiasVentana) en la clave del nodo
                    if (nodoMod.Tag is AgNodoArbolKey keyMod)
                    {
                        var resDim = AgNodoArbolKeyMapper.AplicarDimensiones(keyMod, dto);
                        if (resDim.IsFailure)
                        {
                            dimensionesDefecto++;
                        }
                        ActualizarToolTipNivel6(nodoMod, keyMod);
                    }

                    if (primerNodoRestaurado == null)
                    {
                        primerNodoRestaurado = nodoMod;
                    }

                    nodoMod.Checked = true;
                    AplicarExclusividadModalidad(nodoMod); // Fase 537: exclusividad al cargar perfil
                    ActualizarEstadoPadre(nodoMod);
                }

                foreach (TreeNode root in treeContextoAcademico.Nodes)
                {
                    RefrescarEstadoRecursivo(root); // Fase 542
                }
            }
            finally
            {
                treeContextoAcademico.EndUpdate();
                _isUpdatingTreeChecks = false;
                ActualizarConteoCombinaciones();

                // Fase 1206: Asegurar visibilidad del primer nodo restaurado, repintar e invalidar cabeceras
                if (primerNodoRestaurado != null)
                {
                    primerNodoRestaurado.EnsureVisible();
                }
                treeContextoAcademico.Invalidate();
                ActualizarCabecerasNivel6();
            }

            // Fase 1580: detalle accionable de las combinaciones realmente fuera de catálogo
            string detalleOmitidas = omitidas.Count > 0
                ? " Ej.: " + string.Join(", ", omitidas.Take(5)) + (omitidas.Count > 5 ? ", …" : string.Empty)
                : string.Empty;

            if (dimensionesDefecto > 0)
            {
                LoggerHelper.Warning($"[AgFormGeneradorHorarios] Fallback MAÑANA/4h/MAR-VIE en {dimensionesDefecto} combinación(es) con configuración incompleta.");
                if (mostrarAvisos)
                {
                    string aviso = huerfanos > 0
                        ? $"Se cargó el perfil.\n- Se omitieron {huerfanos} combinación(es) porque su período, sede o programa no existe en el catálogo activo.{detalleOmitidas}\n- Se aplicaron dimensiones por defecto en {dimensionesDefecto} combinación(es) incompletas."
                        : $"Se cargó el perfil. Se aplicaron dimensiones por defecto en {dimensionesDefecto} combinación(es) incompletas.";
                    MessageBox.Show(this, aviso, "Aviso de Carga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (huerfanos > 0)
            {
                LoggerHelper.Warning($"[AgFormGeneradorHorarios] Perfil cargado con {huerfanos} combinación(es) huérfana(s) omitida(s): {string.Join(", ", omitidas)}");
                if (mostrarAvisos)
                {
                    MessageBox.Show(this, $"Se cargó el perfil. Se omitieron {huerfanos} combinación(es) porque su período, sede o programa no existe en el catálogo activo.{detalleOmitidas}", "Aviso de Carga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            return dimensionesDefecto;
        }

        private void DesmarcarRecursivo(TreeNode node)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = false;
                DesmarcarRecursivo(child);
            }
        }

        #endregion

        /// <summary>
        /// Genera de forma transaccional y por lotes los registros en ag_aula correspondientes a las tuplas S seleccionadas (Fase 160).
        /// Incluye detección de duplicados mediante AulaRef y aplicación de tope de seguridad (N &lt;= 5000).
        /// Vincula de forma transaccional asignaturas en ag_aula_asignatura respetando el esquema canónico de asignaturas y V49 con fallback defensivo (Fases 218, 262, 264 y 266).
        /// Notifica la persistencia al vault cifrado mediante MarkDirty (Fase 268) y reporta explícitamente aulas huérfanas de asignaturas (Fase 275).
        /// </summary>
        private async Task GenerarEncabezadosAulasAsync()
        {
            // Fase 1081: Verificar descartes de configuración incompleta y alertar al usuario antes de proceder
            var descartes = new List<string>();
            var dtos = ObtenerCombinacionesSeleccionadas(descartes);
            if (descartes.Count > 0)
            {
                string detalle = string.Join("\n", descartes.Take(10));
                if (descartes.Count > 10) detalle += $"\n... y {descartes.Count - 10} más.";
                MessageBox.Show(
                    $"Se encontraron {descartes.Count} combinaciones seleccionadas con configuración incompleta o inválida:\n\n{detalle}\n\nPor favor configure jornada, ventana de días y horas/día antes de generar aulas.",
                    "Configuración Incompleta",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Fase 545: Validar que no haya mezcla de múltiples períodos
            var periodosGenerar = dtos.Select(d => d.IdPeriodo).Where(p => p > 0).Distinct().ToList();
            if (periodosGenerar.Count > 1)
            {
                MessageBox.Show($"Las combinaciones seleccionadas pertenecen a {periodosGenerar.Count} períodos académicos distintos (IDs: {string.Join(", ", periodosGenerar)}).\n\nNo es posible generar aulas simultáneamente para múltiples períodos. Por favor seleccione ramas correspondientes a un único período.", "Mezcla de Períodos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var expandidor = new AgSeleccionAulaExpandidor();
            const int TopeCombinaciones = 5000; // Fase 536: ampliado para soporte multi-sede/multi-programa
            var valRes = expandidor.ValidarYExpandir(dtos, TopeCombinaciones);

            if (valRes.IsFailure)
            {
                MessageBox.Show(valRes.Error, "Generación de Aulas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var lista = valRes.Value;
            var confirm = MessageBox.Show(
                $"¿Desea generar {lista.Count} encabezados de aula para las ramas seleccionadas?\n\nLa operación es transaccional y omitirá duplicados existentes.",
                "Confirmación de Generación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            var traza = CrearTraza("GenerarAulas");
            traza.Inicio($"combinaciones={lista.Count}");
            traza.Paso("P04_GenerarEncabezadosAulas", $"combinaciones={lista.Count}");

            try
            {
                int insertadas = 0;
                int omitidas = 0;
                int aseguradas = 0;
                int sinAsignaturas = 0;
                int sinHoras = 0;
                var porCombinacion = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                string sub = SyncContext.ActiveSubunidad ?? string.Empty;

                var periodoSelGenerar = cmbPeriodos.Primario<Periodo>() ?? (cmbPeriodos.SelectedItem as Periodo);
                if (periodoSelGenerar != null)
                {
                    // Fase 453: Traza de período objetivo en GenerarEncabezadosAulasAsync
                    traza.Dato("PeriodoObjetivo", $"{periodoSelGenerar.Nombre ?? periodoSelGenerar.Id.ToString()} ({periodoSelGenerar.PeriodoInicio ?? ""}→{periodoSelGenerar.PeriodoFin ?? ""})");
                }
                traza.Dato("Subunidad", sub);
                traza.Dato("TotalCombinaciones", lista.Count);
                // Fase 476: Trazabilidad de combinaciones en GenerarAulas
                traza.Dato("Combinaciones", string.Join(" | ", lista.Select(x => x.AulaRef)));

                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    using (var trans = db.BeginTransaction())
                    {
                        var existentes = new HashSet<string>(
                            await db.QueryAsync<string>("SELECT aula_ref FROM ag_aula WHERE aula_ref IS NOT NULL AND aula_ref <> ''", transaction: trans),
                            StringComparer.OrdinalIgnoreCase);

                        string sqlInsert = @"
INSERT INTO ag_aula (
    idsede, idperiodo, idprograma, codigo, nombre, capacidad, tipo_aula,
    es_virtual, estado, caracteristicas_json, semestre, grupo, modalidadVinculacion,
    fecha_inicio, fecha_fin, semestre_inicio, semestre_fin, ventana_origen, aula_ref, subunidad,
    jornada, horas_dia, dias_ventana
) VALUES (
    @IdSede, @IdPeriodo, @IdPrograma, @Codigo, @Nombre, @Capacidad, @TipoAula,
    @EsVirtual, @Estado, @CaracteristicasJson, @Semestre, @Grupo, @ModalidadVinculacion,
    @FechaInicio, @FechaFin, @SemestreInicio, @SemestreFin, @VentanaOrigen, @AulaRef, @Subunidad,
    @Jornada, @HorasDia, @DiasVentana
);";

                        bool primeraTrazaSql = true;
                        foreach (var item in lista)
                        {
                            int aseguradasItem = 0;
                            var (semRom, _) = AgFormatoAsignatura.ParsearSemestreGrupo(item.Semestre);
                            string semArab = AgFormatoAsignatura.SemestreArabigo(semRom);

                            // Fase 447 / Fase 467: Verificar si existen asignaturas en el pensum ANTES de insertar en ag_aula (evita aulas huérfanas)
                            const string sqlAsigCount = @"
SELECT COUNT(1) FROM asignaturas a
WHERE a.idprograma = @IdPrograma
  AND (@Semestre = '' OR a.semestre = @Semestre OR (@SemestreArabigo <> '' AND a.semestre = @SemestreArabigo))
  AND (a.subunidad = @Sub OR a.subunidad = '' OR a.subunidad IS NULL);";
                            int asigsCount = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                                sqlAsigCount, new { IdPrograma = item.IdPrograma, Semestre = semRom, SemestreArabigo = semArab, Sub = sub }, transaction: trans));

                            if (asigsCount == 0)
                            {
                                sinAsignaturas++;
                                omitidas++;
                                traza.Dato("AulaOmitidaSinAsignaturas", item.AulaRef);
                                porCombinacion[item.AulaRef] = 0;
                                continue; // ⛔ sin INSERT en ag_aula
                            }

                            long idAula = 0;
                            if (existentes.Contains(item.AulaRef))
                            {
                                omitidas++;
                                idAula = await db.ExecuteScalarAsync<long>(
                                    "SELECT id FROM ag_aula WHERE aula_ref = @AulaRef LIMIT 1;",
                                    new { item.AulaRef }, transaction: trans);

                                // Ronda 206 / Fase 1052 (I1): Sincronizar dimensiones de decisión ante reconfiguración del árbol
                                await db.ExecuteAsync(
                                    @"UPDATE ag_aula 
                                      SET jornada = @Jornada, horas_dia = @HorasDia, dias_ventana = @DiasVentana 
                                      WHERE aula_ref = @AulaRef;",
                                    new { item.Jornada, item.HorasDia, item.DiasVentana, item.AulaRef },
                                    transaction: trans);
                            }
                            else
                            {
                                var aula = AgAulaFactory.CrearDesdeDto(item, _ventanaResuelta, sub);
                                await db.ExecuteAsync(sqlInsert, aula, transaction: trans);
                                idAula = await db.ExecuteScalarAsync<long>("SELECT last_insert_rowid();", transaction: trans);
                                existentes.Add(item.AulaRef);
                                insertadas++;
                            }

                            // Fase 213, 218, 262, 264, 266 & 474: Asegurar que ag_aula_asignatura esté poblada tanto en aulas nuevas como preexistentes
                            if (idAula > 0)
                            {
                                // Fase 1016 & Ronda 206: ORDER BY COALESCE(aa.orden_curricular, 0), a.id para orden determinista
                                string sqlAsigs = @"
SELECT DISTINCT a.id AS Id, a.int_horaria AS IntHoraria,
       COALESCE(aa.dias_patron, '') AS DiasPatron,
       COALESCE(aa.orden_curricular, 0) AS OrdenCurricular
FROM asignaturas a
LEFT JOIN ag_aula_asignatura aa ON aa.id_aula = @IdAula AND aa.id_asignatura = a.id
WHERE a.idprograma = @IdPrograma 
  AND (@Semestre = '' OR a.semestre = @Semestre OR (@SemestreArabigo <> '' AND a.semestre = @SemestreArabigo))
  AND (a.subunidad = @Sub OR a.subunidad = '' OR a.subunidad IS NULL)
ORDER BY COALESCE(aa.orden_curricular, 0), a.id;";

                                // Fase 441: Mapeo nativo POCO con AsigSeedRow en net48 y soporte int_horaria anulable
                                var asigsData = (await db.QueryAsync<AsigSeedRow>(
                                    sqlAsigs,
                                    new { IdAula = idAula, IdPrograma = item.IdPrograma, Semestre = semRom, SemestreArabigo = semArab, Sub = sub },
                                    transaction: trans)).ToList();

                                if (primeraTrazaSql)
                                {
                                    traza.Sql("asignaturas", sqlAsigs, new { IdPrograma = item.IdPrograma, Semestre = semRom, SemestreArabigo = semArab, Sub = sub });
                                    traza.Conteo("idsAsigPrimeraMuestra", asigsData.Count);
                                }

                                if (asigsData.Count > 0)
                                {
                                    string sqlInsertAsig = @"
INSERT INTO ag_aula_asignatura (
    id_aula, id_asignatura, exclusivo, prioridad, horas_objetivo, orden_curricular, requerida, dias_patron, subunidad, origen_patron
) VALUES (
    @IdAula, @IdAsig, 0, 0, @HorasObjetivo, 1, 1, @DiasPatron, @Sub, 'AUTOMATICO'
)
ON CONFLICT(id_aula, id_asignatura) DO UPDATE SET
    dias_patron = CASE
        WHEN UPPER(COALESCE(ag_aula_asignatura.origen_patron,'AUTOMATICO'))='AUTOMATICO'
         AND excluded.dias_patron <> '' THEN excluded.dias_patron
        ELSE ag_aula_asignatura.dias_patron END,
    horas_objetivo = CASE WHEN excluded.horas_objetivo > 0 THEN excluded.horas_objetivo ELSE ag_aula_asignatura.horas_objetivo END;";

                                    // Fase 1013: sembrar el carril (2 días) rotado por asignatura, no la ventana completa
                                    int idxAulaAsig = 0;
                                    foreach (var aItem in asigsData)
                                    {
                                        // Fase 442 / Fase 444: Omitir asignaturas sin horas definidas en catálogo (evita inflar demanda con 64h ficticias)
                                        if (!aItem.IntHoraria.HasValue || aItem.IntHoraria.Value <= 0)
                                        {
                                            traza.Dato("AulaAsignaturaSinHoras", $"asig {aItem.Id} sin int_horaria; se omite (use HorasObjetivo explícito)");
                                            sinHoras++;
                                            continue;
                                        }

                                        double horasObj = (double)aItem.IntHoraria.Value;
                                        int horas = (int)horasObj;
                                        // Fase 1073 / Fase 1240: Sembrar subconjunto estricto por banda (<= 2 días), nunca la ventana completa en ventanas multi-día
                                        string diasPatron;
                                        var ventanaIso = AgDiasVentana.Dias(item.DiasVentana);
                                        if (ventanaIso.Count > 0)
                                        {
                                            var ventanaDias = ventanaIso;
                                            int nDias = ventanaDias.Count;
                                            int bandasSiembra = nDias == 2 ? 2 : Math.Max(1, (int)Math.Ceiling(nDias / 2.0));
                                            int bandaIdx = idxAulaAsig % bandasSiembra;
                                            int tamBanda = (int)Math.Ceiling(nDias / (double)bandasSiembra);
                                            int desde = bandaIdx * tamBanda;
                                            var diasBanda = ventanaDias.Skip(desde).Take(Math.Min(tamBanda, nDias - desde)).ToList();
                                            // Fase 1232: habilitado tras Fase 1226
                                            const bool USAR_BANDAS_DIAS = true;
                                            diasPatron = USAR_BANDAS_DIAS
                                                ? string.Join(",", diasBanda)
                                                : string.Join(",", AgCarriles.DelCarril(ventanaDias, idxAulaAsig));
                                            if (!string.IsNullOrWhiteSpace(aItem.DiasPatron) &&
                                                !string.Equals(aItem.DiasPatron, diasPatron, StringComparison.Ordinal))
                                            {
                                                traza.Dato("DiasPatronRecalculado", $"aula {item.AulaRef} asig {aItem.Id}: [{aItem.DiasPatron}] → [{diasPatron}]");
                                            }
                                        }
                                        else
                                        {
                                            diasPatron = string.Empty;
                                            traza.Dato("AulaAsignaturaSinVentana", $"aula {item.AulaRef} asig {aItem.Id}: sin dias_ventana");
                                        }
                                        idxAulaAsig++;

                                        // Fase 477: Contar solo inserciones reales para evitar sobreconteo tras INSERT OR IGNORE
                                        int filas = await db.ExecuteAsync(
                                            sqlInsertAsig,
                                            new { IdAula = idAula, IdAsig = aItem.Id, HorasObjetivo = horasObj, DiasPatron = diasPatron, Sub = sub },
                                            transaction: trans);

                                        if (filas > 0)
                                        {
                                            aseguradas++;
                                            aseguradasItem++;
                                        }
                                    }
                                }
                                primeraTrazaSql = false;
                            }

                            porCombinacion[item.AulaRef] = aseguradasItem;
                        }

                        trans.Commit();
                        try { ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty(); } catch { } // Fase 268: Persistir ag_aula y ag_aula_asignatura al vault
                    }
                }

                // Fase 479: Aseguradas acumulado real y desglose por combinación
                traza.Conteo("Aseguradas", aseguradas);
                traza.Dato("AseguradasPorCombinacion", string.Join(" | ", porCombinacion.Select(kv => $"{kv.Key}={kv.Value}")));
                traza.Dato("Insertadas", insertadas);
                traza.Dato("Omitidas", omitidas);
                // Fase 478: Se elimina traza.Dato("Aseguradas", ...) redundante
                traza.Dato("SinHoras", sinHoras);
                traza.Dato("SinAsignaturas", sinAsignaturas);
                traza.Dato("FaltantesEnCatalogo", $"{sinAsignaturas}");
                if (sinAsignaturas > 0)
                {
                    traza.Causa(AgCausaFallo.SinAulaAsignaturas, $"{sinAsignaturas} aulas quedaron sin asignaturas vinculadas");
                }
                traza.Fin(sinAsignaturas == 0, $"insertadas={insertadas}, omitidas={omitidas}, aseguradas={aseguradas}, sinHoras={sinHoras}, sinAsignaturas={sinAsignaturas}");

                // Fase 275 / Fase 444 / Fase 474: Reporte accionable con métricas separadas de aulas y asignaturas aseguradas
                string mensajeResumen = $"Proceso completado:\n- Aulas creadas: {insertadas}\n- Aulas preexistentes: {omitidas}\n- Aulas con materias aseguradas: {aseguradas}";
                if (sinHoras > 0)
                {
                    mensajeResumen += $"\n- Asignaturas sin int_horaria omitidas: {sinHoras}";
                }
                if (sinAsignaturas > 0)
                {
                    mensajeResumen += $"\n- Aulas sin asignaturas vinculadas: {sinAsignaturas} (verifique catálogo 'asignaturas')";
                }

                MessageBox.Show(
                    mensajeResumen,
                    "Aulas Generadas",
                    MessageBoxButtons.OK,
                    sinAsignaturas > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

                txtLogAG.AppendText($"[{DateTime.Now:HH:mm:ss}] Generación de encabezados: {insertadas} creadas, {omitidas} duplicadas omitidas, {sinHoras} sin int_horaria omitidas.\r\n");
                if (sinAsignaturas > 0)
                {
                    txtLogAG.AppendText($"[{DateTime.Now:HH:mm:ss}] [AVISO] {sinAsignaturas} aula(s) quedaron sin ag_aula_asignatura (programa sin asignaturas en catálogo 'asignaturas').\r\n");
                }
            }
            catch (Exception ex)
            {
                traza.Fallo("P04_GenerarEncabezadosAulas", ex);
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error generando aulas", ex);
                MessageBox.Show($"Error al generar aulas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Gestión de Filtros de Selección Persistentes (P6)

        // Fase 1442: Retorna bool indicando si los filtros fueron leídos correctamente sin silenciar fallos de BD
        private async Task<bool> RestaurarFiltrosAsync()
        {
            try
            {
                string sub = (SyncContext.ActiveSubunidad ?? string.Empty).Trim();
                var res = await _filtroSeleccionRepo.ObtenerFiltrosAsync(sub, CancellationToken.None);
                if (res.IsFailure || res.Value == null)
                {
                    LoggerHelper.Warning("[AgFormGeneradorHorarios] Filtros no restaurados; se inhibe el guardado para no borrar la persistencia.");
                    return false;
                }

                var f = res.Value;
                // Fase 1449: Reordenar restauración: Tipo -> lista filtrada -> marcas
                if (!string.IsNullOrWhiteSpace(f.TipoPrograma))
                {
                    int idx = cmbTipoPrograma.FindStringExact(f.TipoPrograma);
                    if (idx >= 0) cmbTipoPrograma.SelectedIndex = idx;
                }

                string tipoActual = cmbTipoPrograma.SelectedItem?.ToString() ?? AgTipoProgramaCatalog.Todos;
                CargarProgramasSegunTipo(tipoActual, f.ProgramasIds);

                // Fase 1424: Si no hay selección guardada, dejar sin marcar (vacío = todos). No forzar el primer período.
                if (f.PeriodosIds.Count > 0) cmbPeriodos.Marcar(f.PeriodosIds);
                if (f.SedesIds.Count > 0) cmbSedes.Marcar(f.SedesIds);

                if (f.TercerosIds.Count > 0)
                {
                    _tercerosSeleccionadosIds = new HashSet<int>(f.TercerosIds);
                }

                ActualizarResumenFiltros();
                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error restaurando filtros de selección", ex);
                return false;
            }
        }

        private async Task GuardarFiltrosAsync(bool forzar = false)
        {
            // Fase 1438 & 1446: bloquear solo si la lectura inicial falló; las acciones explícitas del usuario rehabilitan el guardado
            if (!_catalogosCargados) return;
            if (!_filtrosRestaurados && !forzar) return;

            try
            {
                string sub = (SyncContext.ActiveSubunidad ?? string.Empty).Trim();
                // Fase 1422 & 1458: Usar SeleccionadosIds() / ProgramasSeleccionados() para obtener los Ids reales
                var f = new AgFiltrosSeleccion
                {
                    PeriodosIds = cmbPeriodos.SeleccionadosIds(),
                    ProgramasIds = ProgramasSeleccionados(),
                    SedesIds = cmbSedes.SeleccionadosIds(),
                    TipoPrograma = cmbTipoPrograma.SelectedItem?.ToString() ?? AgTipoProgramaCatalog.Todos,
                    TercerosIds = _tercerosSeleccionadosIds != null ? _tercerosSeleccionadosIds.ToList() : new List<int>()
                };

                var res = await _filtroSeleccionRepo.GuardarFiltrosAsync(f, sub, CancellationToken.None);
                if (res.IsSuccess) _filtrosRestaurados = true;
                else LoggerHelper.Warning("[AgFormGeneradorHorarios] Guardado de filtros falló: " + res.Error);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error guardando filtros de selección", ex);
            }
        }

        // Fase 1425, 1438 & 1446: Persistencia síncrona directa con verificación de _catalogosCargados y _filtrosRestaurados
        private void GuardarFiltrosSincrono()
        {
            if (!_catalogosCargados || !_filtrosRestaurados) return;

            try
            {
                string sub = (SyncContext.ActiveSubunidad ?? string.Empty).Trim();
                var r1 = _filtroSeleccionRepo.Reemplazar(sub, AgFiltroTipo.Periodo, cmbPeriodos.SeleccionadosIds());
                if (r1.IsFailure) LoggerHelper.Warning("[AgFormGeneradorHorarios] No se persistió Periodo: " + r1.Error);
                var r2 = _filtroSeleccionRepo.Reemplazar(sub, AgFiltroTipo.Sede, cmbSedes.SeleccionadosIds());
                if (r2.IsFailure) LoggerHelper.Warning("[AgFormGeneradorHorarios] No se persistió Sede: " + r2.Error);
                var r3 = _filtroSeleccionRepo.Reemplazar(sub, AgFiltroTipo.Programa, ProgramasSeleccionados());
                if (r3.IsFailure) LoggerHelper.Warning("[AgFormGeneradorHorarios] No se persistió Programa: " + r3.Error);
                var r4 = _filtroSeleccionRepo.Reemplazar(sub, AgFiltroTipo.Tercero, _tercerosSeleccionadosIds ?? new HashSet<int>());
                if (r4.IsFailure) LoggerHelper.Warning("[AgFormGeneradorHorarios] No se persistió Tercero: " + r4.Error);
                string tipoSel = cmbTipoPrograma.SelectedItem?.ToString();
                string tipoCmp = !string.IsNullOrWhiteSpace(tipoSel) ? tipoSel.Trim() : AgTipoProgramaCatalog.Todos;
                int tipoRef = string.Equals(tipoCmp, AgTipoProgramaCatalog.Pregrado, StringComparison.OrdinalIgnoreCase) ? 1
                            : string.Equals(tipoCmp, AgTipoProgramaCatalog.Posgrado, StringComparison.OrdinalIgnoreCase) ? 2 : 0;
                var r5 = _filtroSeleccionRepo.Reemplazar(sub, AgFiltroTipo.TipoPrograma, new[] { tipoRef });
                if (r5.IsFailure) LoggerHelper.Warning("[AgFormGeneradorHorarios] No se persistió TipoPrograma: " + r5.Error);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFormGeneradorHorarios] Error guardando filtros al cierre", ex);
            }
        }

        private async Task EjecutarFiltroArbolAsync()
        {
            // Fase 1729: Inhibir acción si el diagnóstico pre-vuelo está en curso
            if (_diagnosticoEnCurso)
            {
                return;
            }

            var expandidos = ObtenerRutasExpandidas();
            PoblarArbolContextoAcademico();
            RestaurarRutasExpandidas(expandidos);
            // Fase 1443: Resetear indicador de árbol desactualizado
            _filtroArbolDesactualizado = false;
            ActualizarResumenFiltros();
            await GuardarFiltrosAsync(forzar: true);
        }

        private void AbrirSeleccionTerceros()
        {
            using (var form = new AgFormTerceros(_connectionFactory, _tercerosSeleccionadosIds))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    _tercerosSeleccionadosIds = new HashSet<int>(form.TercerosSeleccionadosIds);
                    ActualizarResumenFiltros();
                    _ = GuardarFiltrosAsync(forzar: true);
                }
            }
        }

        /// <summary>
        /// Carga en cmbProgramas únicamente los programas que coinciden con el tipo/modalidad seleccionada.
        /// Preserva las marcas previas solo para los programas que siguen visibles y retiene en memoria los programas ocultos.
        /// Fases 1448, 1449 y 1458.
        /// </summary>
        private void CargarProgramasSegunTipo(string tipo, IEnumerable<int> idsPreferidos = null)
        {
            var progsFiltrados = AgTipoProgramaCatalog.Filtrar(_programas, tipo);
            var idsVisibles = new HashSet<int>(progsFiltrados.Select(p => p.Id));
            var marcadosPrevios = idsPreferidos != null
                ? new HashSet<int>(idsPreferidos)
                : new HashSet<int>(cmbProgramas.SeleccionadosIds().Concat(_programasOcultos));

            _programasOcultos.Clear();
            _programasOcultos.UnionWith(marcadosPrevios.Where(id => !idsVisibles.Contains(id)));

            var aMarcar = marcadosPrevios.Where(idsVisibles.Contains).ToList();
            cmbProgramas.Cargar(progsFiltrados, p => p.Id, p => AgTipoProgramaCatalog.EtiquetaPrograma(p));
            if (aMarcar.Count > 0)
            {
                cmbProgramas.Marcar(aMarcar);
            }
            else
            {
                cmbProgramas.Limpiar();
            }
        }

        // Fase 1458: unión visible + ocultos para persistencia y resumen
        private List<int> ProgramasSeleccionados()
        {
            var visibles = cmbProgramas.SeleccionadosIds();
            if (cmbProgramas.TodosMarcadoPorUsuario)
            {
                _programasOcultos.Clear();
                return new List<int>();
            }
            return visibles.Concat(_programasOcultos).Distinct().ToList();
        }

        // Fase 1422, 1443, 1450, 1457, 1458, 1461, 1464 & 1469: Conteo usando SeleccionadosIds() / ProgramasSeleccionados() y resumen integrado
        private void ActualizarResumenFiltros()
        {
            int pCount = cmbPeriodos.SeleccionadosIds().Count;
            int sCount = cmbSedes.SeleccionadosIds().Count;
            var progsSel = ProgramasSeleccionados();
            // Fase 1472 / 1474: visibles = marcas vigentes en el combo (restaurado desde backup_ronda267).
            int progVisibles = cmbProgramas.SeleccionadosIds().Count;
            string progTxt = AgResumenFiltros.Programas(progsSel.Count, progVisibles);
            int tCount = _tercerosSeleccionadosIds?.Count ?? 0;
            string tipo = cmbTipoPrograma.SelectedItem?.ToString() ?? AgTipoProgramaCatalog.Todos;

            string hint = _filtroArbolDesactualizado ? "   ·   ⚠ «Filtrar árbol»" : string.Empty;
            grpSeleccion.Text = $"{TituloGrpSeleccion}   ·   Filtros: {tipo} | Períodos: {(pCount == 0 ? "Todos" : pCount.ToString())} | Sedes: {(sCount == 0 ? "Todas" : sCount.ToString())} | Programas: {progTxt} | Terceros: {(tCount == 0 ? "Todos" : tCount.ToString())}{hint}";
        }

        private List<string> ObtenerRutasExpandidas()
        {
            var res = new List<string>();
            void Visitar(TreeNode n, string ruta)
            {
                string tagStr = n.Tag is AgNodoArbolKey k ? $"{k.Nivel}:{k.Id}:{k.Codigo}" : n.Text;
                string actual = string.IsNullOrEmpty(ruta) ? tagStr : ruta + "/" + tagStr;
                if (n.IsExpanded)
                {
                    res.Add(actual);
                }
                foreach (TreeNode hijo in n.Nodes)
                {
                    Visitar(hijo, actual);
                }
            }
            foreach (TreeNode raiz in treeContextoAcademico.Nodes)
            {
                Visitar(raiz, "");
            }
            return res;
        }

        private void RestaurarRutasExpandidas(List<string> rutas)
        {
            if (rutas == null || rutas.Count == 0) return;
            var set = new HashSet<string>(rutas);
            void Visitar(TreeNode n, string ruta)
            {
                string tagStr = n.Tag is AgNodoArbolKey k ? $"{k.Nivel}:{k.Id}:{k.Codigo}" : n.Text;
                string actual = string.IsNullOrEmpty(ruta) ? tagStr : ruta + "/" + tagStr;
                if (set.Contains(actual))
                {
                    n.Expand();
                }
                foreach (TreeNode hijo in n.Nodes)
                {
                    Visitar(hijo, actual);
                }
            }
            foreach (TreeNode raiz in treeContextoAcademico.Nodes)
            {
                Visitar(raiz, "");
            }
        }

        #endregion

        #region Clases Auxiliares Internas (Fase 441)

        /// <summary>
        /// Fila auxiliar para la siembra de asignaturas en ag_aula_asignatura (Fase 441).
        /// Mapeable de forma nativa por Dapper en .NET Framework 4.8 con soporte de int_horaria anulable.
        /// </summary>
        private sealed class AsigSeedRow
        {
            public int Id { get; set; }
            public int? IntHoraria { get; set; }
            public string DiasPatron { get; set; }
            public int OrdenCurricular { get; set; }
        }

        #endregion
    }
}
