using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    public partial class FormUsuarios : Form
    {
        private readonly AcademicDbSession _session;
        private readonly Func<IDbConnection> _connectionFactory;
        private BindingSource _bs = new BindingSource();
        private Usuario _currentUsuario;

        // Fase 146 y Fase 147: Constantes y campos para gestión de perfil por defecto y preservación de perfiles huérfanos
        private const string PerfilPorDefecto = "Admin";
        private int _indicePerfilOriginal = -1;
        private string _perfilHuerfano = null;

        // Fase 152: Flag de estado de validación de perfiles
        private bool _perfilesValidados = false;
        // Fase 164: Detalle del motivo por el cual los perfiles no fueron validados
        private string _motivoPerfilesInvalidos = null;
        // Fase 166: Detalle del motivo por el cual la lista de subunidades está vacía
        private string _motivoSubunidadesVacias = null;
        // Fase 168: Conjunto de subunidades oficiales cargadas desde la tabla territoriales
        private readonly HashSet<string> _subunidadesOficiales = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private class ActivoOption
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }

        /// <summary>
        /// Constructor del formulario de usuarios.
        /// </summary>
        /// <param name="session">Sesión de base de datos compartida.</param>
        public FormUsuarios(AcademicDbSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _connectionFactory = () => session.Connection;
            InitializeComponent();
            
            cmbActivo.Items.Add(new ActivoOption { Text = "Inactivo", Value = 0 });
            cmbActivo.Items.Add(new ActivoOption { Text = "Activo", Value = 1 });
            cmbActivo.DisplayMember = "Text";
            cmbActivo.ValueMember = "Value";
            cmbActivo.SelectedIndex = 1;

            // Se remueve la carga de perfiles hardcoded y se reemplaza por carga dinámica desde base de datos.
            CargarPerfiles();

            try
            {
                var db = _connectionFactory();
                var subunidades = Dapper.SqlMapper.Query<string>(db, "SELECT DISTINCT subunidad FROM territoriales WHERE subunidad IS NOT NULL AND subunidad != '' ORDER BY subunidad").ToList();
                if (subunidades.Count > 0)
                {
                    foreach (var s in subunidades)
                    {
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            _subunidadesOficiales.Add(s.Trim());
                        }
                    }
                    cmbSubunidad.Items.AddRange(subunidades.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToArray());
                }
                else
                {
                    // Fase 166: Registrar causa si la tabla está vacía
                    _motivoSubunidadesVacias = "tabla 'territoriales' no contiene subunidades";
                }
            }
            catch (Exception ex)
            {
                // Fase 166: Registrar diagnóstico si la consulta falla
                LoggerHelper.Warning($"[FormUsuarios] Error cargando subunidades: {ex.Message}");
                _motivoSubunidadesVacias = $"error al consultar subunidades ({ex.Message})";
            }

            SetupGrid();
            CargarUsuarios();
        }

        /// <summary>
        /// Carga los perfiles disponibles de forma dinámica desde la tabla de permisos en la base de datos.
        /// (Fase 146, Fase 149, Fase 150, Fase 152, Fase 156, Fase 160)
        /// </summary>
        private void CargarPerfiles()
        {
            try
            {
                cmbPerfil.Items.Clear();
                cmbPerfil.Enabled = true;
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetTodosLosPermisos();
                
                if (result.IsSuccess)
                {
                    // Fase 150: Filtrar nulos/vacíos, normalizar y deduplicar por OrdinalIgnoreCase
                    var perfiles = result.Value
                        .Select(p => p?.Perfil)
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Select(p => p.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();

                    if (perfiles.Length > 0)
                    {
                        cmbPerfil.Items.AddRange(perfiles);
                        _perfilesValidados = true;
                        _motivoPerfilesInvalidos = null;
                        SeleccionarPerfilPorDefecto();
                        return;
                    }
                }

                // Fase 160 y Fase 164: Si no hay perfiles válidos o falló la consulta, deshabilitar sin agregar placeholder falso y registrar motivo
                _perfilesValidados = false;
                cmbPerfil.Items.Clear();
                cmbPerfil.Enabled = false;
                if (!result.IsSuccess)
                {
                    _motivoPerfilesInvalidos = result.Error;
                    LoggerHelper.Warning($"[FormUsuarios] GetTodosLosPermisos falló: {result.Error}");
                }
                else
                {
                    _motivoPerfilesInvalidos = "tabla 'permisos' vacía";
                    LoggerHelper.Warning("[FormUsuarios] La tabla 'permisos' no contiene perfiles.");
                }
            }
            catch (Exception ex)
            {
                // Fase 156, Fase 160 y Fase 164: Normalización en bloque catch sin placeholder y trazabilidad del motivo
                LoggerHelper.Error("[FormUsuarios] Excepción cargando perfiles", ex);
                _perfilesValidados = false;
                _motivoPerfilesInvalidos = $"excepción: {ex.Message}";
                cmbPerfil.Items.Clear();
                cmbPerfil.Enabled = false;
            }
        }

        /// <summary>
        /// Busca el índice de un perfil en cmbPerfil realizando una comparación insensible a mayúsculas/minúsculas y eliminando espacios adicionales.
        /// (Fase 148)
        /// </summary>
        /// <param name="perfil">Nombre del perfil a buscar.</param>
        /// <returns>Índice base 0 si se encuentra; de lo contrario -1.</returns>
        private int IndicePerfil(string perfil)
        {
            if (string.IsNullOrWhiteSpace(perfil)) return -1;
            for (int i = 0; i < cmbPerfil.Items.Count; i++)
            {
                var item = cmbPerfil.Items[i] as string;
                if (item != null && string.Equals(item.Trim(), perfil.Trim(), StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Selecciona el perfil por defecto ("Admin") o el primer perfil disponible si "Admin" no existe.
        /// (Fase 146)
        /// </summary>
        private void SeleccionarPerfilPorDefecto()
        {
            int idx = IndicePerfil(PerfilPorDefecto);
            cmbPerfil.SelectedIndex = idx >= 0 ? idx : (cmbPerfil.Items.Count > 0 ? 0 : -1);
        }

        private void SetupGrid()
        {
            dgvUsuarios.AutoGenerateColumns = false;
            dgvUsuarios.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "NombreUsuario", HeaderText = "Usuario", Width = 150 });
            dgvUsuarios.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Nombre", HeaderText = "Nombre", Width = 200 });
            dgvUsuarios.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Perfil", HeaderText = "Perfil", Width = 100 });
            dgvUsuarios.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Activo", HeaderText = "Activo", Width = 60 });
            dgvUsuarios.DataSource = _bs;
            
            dgvUsuarios.SelectionChanged += DgvUsuarios_SelectionChanged;
        }

        private void CargarUsuarios()
        {
            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                var result = repo.GetUsuarios();
                if (result.IsSuccess)
                {
                    _bs.DataSource = result.Value.ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar usuarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvUsuarios_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvUsuarios.CurrentRow?.DataBoundItem is Usuario u)
            {
                _currentUsuario = u;
                txtNombre.Text = u.Nombre;
                txtNombreUsuario.Text = u.NombreUsuario;
                
                // Fase 147, Fase 148, Fase 154 y Fase 158: Manejo de perfil huérfano, nulo y selección desmarcada
                _perfilHuerfano = u.Perfil;
                int idxPerfil = IndicePerfil(u.Perfil);
                if (idxPerfil >= 0)
                {
                    cmbPerfil.SelectedIndex = idxPerfil;
                }
                else
                {
                    // Fase 154 y Fase 158: Desmarcar (-1) y registrar diagnóstico diferenciado si es null/vacío
                    cmbPerfil.SelectedIndex = -1;
                    if (string.IsNullOrWhiteSpace(u.Perfil))
                    {
                        LoggerHelper.Warning($"[FormUsuarios] Usuario '{u.NombreUsuario}' sin perfil asignado.");
                    }
                    else
                    {
                        LoggerHelper.Warning($"[FormUsuarios] Perfil '{u.Perfil}' del usuario '{u.NombreUsuario}' no existe en la tabla permisos.");
                    }
                }
                _indicePerfilOriginal = cmbPerfil.SelectedIndex;
                
                // Fase 165 y Fase 167: Búsqueda case-insensitive y preservación de subunidades no listadas
                int idxSub = -1;
                for (int i = 0; i < cmbSubunidad.Items.Count; i++)
                {
                    var item = cmbSubunidad.Items[i] as string;
                    if (item != null && string.Equals(item.Trim(), (u.Subunidad ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        idxSub = i;
                        break;
                    }
                }

                if (idxSub >= 0)
                {
                    cmbSubunidad.SelectedIndex = idxSub;
                }
                else if (!string.IsNullOrWhiteSpace(u.Subunidad))
                {
                    // Fase 165: Subunidad no listada en 'territoriales' → agregar y seleccionar para preservar en el guardado
                    cmbSubunidad.Items.Add(u.Subunidad.Trim());
                    cmbSubunidad.SelectedItem = u.Subunidad.Trim();
                }
                else
                {
                    // Fase 167: Usuario sin subunidad → desmarcar (-1) para exigir elección explícita y evitar asignar la primera subunidad
                    cmbSubunidad.SelectedIndex = -1;
                }
                
                int idx = 1;
                for (int i = 0; i < cmbActivo.Items.Count; i++)
                {
                    if (((ActivoOption)cmbActivo.Items[i]).Value == u.Activo)
                    {
                        idx = i;
                        break;
                    }
                }
                cmbActivo.SelectedIndex = idx;
                
                txtFirma.Text = u.Firma;
                MostrarFirma(u.Firma);
                
                btnEliminar.Enabled = true;
            }
            else
            {
                LimpiarCampos();
            }
        }

        private void LimpiarCampos()
        {
            _currentUsuario = null;
            txtNombre.Clear();
            txtNombreUsuario.Clear();

            // Fase 146 y Fase 147: Resetear perfil huérfano y restablecer perfil por defecto ("Admin")
            _perfilHuerfano = null;
            SeleccionarPerfilPorDefecto();
            _indicePerfilOriginal = cmbPerfil.SelectedIndex;

            // Fase 166: Desmarcar selección para exigir elección explícita de subunidad
            cmbSubunidad.SelectedIndex = -1;
            cmbActivo.SelectedIndex = 1;
            txtFirma.Clear();
            picFirma.Image?.Dispose();
            picFirma.Image = null;
            btnEliminar.Enabled = false;
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
            txtNombreUsuario.Focus();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNombreUsuario.Text))
            {
                MessageBox.Show("El nombre de usuario no puede estar vacío.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validar tamaño de firma (max ~150KB de base64 ≈ 110KB imagen)
            if (!string.IsNullOrWhiteSpace(txtFirma.Text) && txtFirma.Text.Length > 150000)
            {
                MessageBox.Show("La imagen de firma es demasiado grande. Seleccione una imagen más pequeña.", 
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("El nombre es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 152, Fase 160 y Fase 164: Validar que los perfiles se hayan cargado correctamente desde BD e informar motivo preciso
            if (!_perfilesValidados)
            {
                string causa = _motivoPerfilesInvalidos ?? "error desconocido";
                MessageBox.Show($"No se pudo cargar la lista de perfiles ({causa}). No es posible guardar ni actualizar usuarios.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fase 158: Mensaje descriptivo si el usuario no tiene perfil asignado
            if (cmbPerfil.SelectedItem == null && string.IsNullOrWhiteSpace(_perfilHuerfano))
            {
                string motivo = _currentUsuario != null && string.IsNullOrWhiteSpace(_currentUsuario.Perfil)
                    ? "El usuario no tiene perfil asignado. Seleccione un perfil válido."
                    : "Seleccione un perfil.";
                MessageBox.Show(motivo, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var db = _connectionFactory();
                var repo = new AcademicRepository(db);
                
                var nombreUsuarioTrimmed = txtNombreUsuario.Text.Trim();
                var nombreTrimmed = txtNombre.Text.Trim();
                
                // Fase 5: Solo validar existencia si el nombre de usuario fue modificado o es un registro nuevo.
                // Esto previene bloqueos por colisiones de Ids nulos o duplicados latentes al actualizar.
                bool nombreModificado = _currentUsuario == null || 
                                        !string.Equals(_currentUsuario.NombreUsuario, nombreUsuarioTrimmed, StringComparison.OrdinalIgnoreCase);

                if (nombreModificado)
                {
                    var existente = repo.GetUsuarioByNombreUsuario(nombreUsuarioTrimmed);
                    if (existente.IsSuccess && existente.Value != null)
                    {
                        MessageBox.Show("Nombre de usuario ya existe.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                string subunidadAnterior = _currentUsuario?.Subunidad;
                string nombreAnterior = _currentUsuario?.NombreUsuario;

                // Fase 147 y Fase 154: Determinar perfil a asignar sin modificar la entidad 'u' aún (Fase 157)
                string perfilSeleccionado = cmbPerfil.SelectedItem?.ToString() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(_perfilHuerfano) && cmbPerfil.SelectedIndex == _indicePerfilOriginal)
                {
                    perfilSeleccionado = _perfilHuerfano;
                }

                // Fase 153 y Fase 159: Validar existencia del perfil en la tabla permisos con modal único consolidado
                var catalogo = repo.GetTodosLosPermisos();
                if (!catalogo.IsSuccess)
                {
                    MessageBox.Show($"No se pudieron validar los perfiles: {catalogo.Error}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                bool perfilExiste = catalogo.Value.Any(p => string.Equals(p.Perfil, perfilSeleccionado, StringComparison.OrdinalIgnoreCase));
                if (!perfilExiste)
                {
                    bool esHuerfanoPreservado = !string.IsNullOrWhiteSpace(_perfilHuerfano)
                        && string.Equals(perfilSeleccionado, _perfilHuerfano, StringComparison.OrdinalIgnoreCase);
                    if (!esHuerfanoPreservado)
                    {
                        MessageBox.Show($"El perfil '{perfilSeleccionado}' no existe en la tabla de permisos. Asigne un perfil válido.",
                            "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    MessageBox.Show($"Se conservará el perfil '{perfilSeleccionado}' porque no se seleccionó otro.\n\n" +
                                    "Advertencia: no existe en permisos; el usuario quedará sin botones visibles.",
                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // Fase 163, Fase 166 y Fase 169: Validar subunidad no vacía usando el catálogo de subunidades oficiales como criterio diagnóstico
                string subunidadSeleccionada = cmbSubunidad.SelectedItem?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(subunidadSeleccionada))
                {
                    string msg = _subunidadesOficiales.Count == 0
                        ? $"No hay subunidades oficiales disponibles ({_motivoSubunidadesVacias ?? "tabla 'territoriales' vacía"}). No es posible asignar subunidad."
                        : "Debe seleccionar una subunidad válida para el usuario.";
                    MessageBox.Show(msg, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 168: Validar que la subunidad sea oficial o coincida con la subunidad legada del usuario actual
                bool subunidadValida = _subunidadesOficiales.Contains(subunidadSeleccionada.Trim())
                    || string.Equals(subunidadSeleccionada.Trim(), _currentUsuario?.Subunidad?.Trim(), StringComparison.OrdinalIgnoreCase);
                if (!subunidadValida)
                {
                    MessageBox.Show($"La subunidad '{subunidadSeleccionada}' no existe en la tabla 'territoriales'. Seleccione una subunidad válida.",
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Fase 157: Mutación atómica de la entidad 'u' solo tras pasar exitosamente todas las validaciones
                var u = _currentUsuario ?? new Usuario();
                u.Nombre = nombreTrimmed;
                u.NombreUsuario = nombreUsuarioTrimmed;
                u.Perfil = perfilSeleccionado;
                u.Subunidad = subunidadSeleccionada;
                u.Firma = txtFirma.Text.Trim();
                u.Activo = cmbActivo.SelectedItem is ActivoOption opt ? opt.Value : 1;

                if (string.IsNullOrEmpty(u.Id))
                {
                    var result = repo.InsertUsuario(u);
                    if (result.IsFailure) { MessageBox.Show($"Error al insertar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    u.Id = result.Value;
                    LoggerHelper.Info($"[FormUsuarios] InsertUsuario guardó el ID: {u.Id}");
                }
                else
                {
                    var result = repo.UpdateUsuario(u);
                    if (result.IsFailure) { MessageBox.Show($"Error al actualizar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    LoggerHelper.Info($"[FormUsuarios] UpdateUsuario retornó SUCCESS para ID: {u.Id}");
                }
                
                // Fase 6 y Fase 4: Actualizar config y requerir reinicio si el usuario actual cambió su subunidad
                bool subunidadCambiada = !string.Equals(subunidadAnterior, u.Subunidad, StringComparison.OrdinalIgnoreCase);

                if (string.Equals(u.NombreUsuario, Environment.UserName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(nombreAnterior, Environment.UserName, StringComparison.OrdinalIgnoreCase))
                {
                    var configRepo = new ConfigRepository(db);
                    var configResult = configRepo.UpdateSubunidad(u.Subunidad);
                    if (configResult.IsFailure)
                    {
                        MessageBox.Show($"Error al actualizar la configuración: {configResult.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    
                    var persistResult = _session.Persist();
                    if (persistResult.IsFailure)
                    {
                        MessageBox.Show($"Error al persistir la base de datos: {persistResult.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                                             
                    if (nombreModificado || subunidadCambiada)
                    {
                        MessageBox.Show("Ha cambiado su subunidad principal o nombre de usuario.\n\n" +
                            "Para aplicar los cambios y cargar la base de datos correcta, todas las ventanas " +
                            "actuales se cerrarán. Debe reiniciar Excel por completo antes de continuar.", 
                            "Reinicio Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.DialogResult = DialogResult.Abort;
                        this.Close();
                        return;
                    }
                }
                else
                {
                    var persistResult = _session.Persist();
                    if (persistResult.IsFailure)
                    {
                        MessageBox.Show($"Error al persistir la base de datos: {persistResult.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                
                MessageBox.Show("Usuario guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Invalidar Ribbon para reflejar cambios de perfil en los botones del usuario actual.
                ConversorPDF.XLL.RibbonController.InvalidateRibbon();
                CargarUsuarios();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (_currentUsuario == null) return;

            if (MessageBox.Show($"¿Está seguro de eliminar el usuario {_currentUsuario.NombreUsuario}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var db = _connectionFactory();
                    var repo = new AcademicRepository(db);
                    var result = repo.DeleteUsuario(_currentUsuario.Id);
                    if (result.IsFailure) { MessageBox.Show($"Error al eliminar: {result.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                    var persistResult = _session.Persist();
                    if (persistResult.IsFailure)
                    {
                        MessageBox.Show($"Error al persistir la base de datos: {persistResult.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    CargarUsuarios();
                    LimpiarCampos();
                    // Invalidar Ribbon: si se eliminó el usuario actual, sus botones deben ocultarse.
                    ConversorPDF.XLL.RibbonController.InvalidateRibbon();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al eliminar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSeleccionarFirma_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Imágenes (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp";
                ofd.Title = "Seleccionar imagen de firma";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var img = Image.FromFile(ofd.FileName))
                        {
                            int maxWidth = 300;
                            Image resized;
                            if (img.Width > maxWidth)
                            {
                                int newHeight = (int)(img.Height * (maxWidth / (float)img.Width));
                                resized = new Bitmap(img, new Size(maxWidth, newHeight));
                            }
                            else
                            {
                                resized = new Bitmap(img);
                            }

                            picFirma.Image?.Dispose();
                            picFirma.Image = resized; // asignación directa, sin clon extra

                            // Convertir a base64
                            using (var ms = new MemoryStream())
                            {
                                resized.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                txtFirma.Text = Convert.ToBase64String(ms.ToArray());
                            }
                            // NO disponer resized — PictureBox es dueño ahora
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al cargar la imagen: {ex.Message}", 
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnLimpiarFirma_Click(object sender, EventArgs e)
        {
            txtFirma.Clear();
            picFirma.Image?.Dispose();
            picFirma.Image = null;
        }

        private void MostrarFirma(string base64)
        {
            picFirma.Image?.Dispose();
            picFirma.Image = null;

            if (!string.IsNullOrWhiteSpace(base64))
            {
                try
                {
                    byte[] bytes = Convert.FromBase64String(base64);
                    using (var ms = new MemoryStream(bytes))
                    {
                        // new Bitmap copia los datos y desacopla del stream
                        picFirma.Image = new Bitmap(Image.FromStream(ms));
                    }
                }
                catch
                {
                    picFirma.Image = null; // Base64 inválido → no mostrar
                }
            }
        }
    }
}
