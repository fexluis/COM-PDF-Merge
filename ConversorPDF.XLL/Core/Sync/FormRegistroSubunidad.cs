using System;
using System.Windows.Forms;
using System.Linq;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Diálogo bloqueante que se muestra cuando el usuario de Windows actual
    ///   NO está registrado en la tabla `usuarios` de la base de datos.
    ///
    /// COMPORTAMIENTO BLOQUEANTE:
    ///   El add-in NO termina de cargar hasta que el usuario:
    ///   (a) Se registre correctamente (DialogResult.OK), o
    ///   (b) Cancele/cierre el diálogo (DialogResult.Cancel → sync deshabilitado).
    ///
    /// DATOS QUE RECOPILA:
    ///   - Nombre completo del usuario (para visualización).
    ///   - Subunidad a la que pertenece (dropdown con lista predefinida).
    ///
    /// PERSISTENCIA:
    ///   Al confirmar, inserta el registro en la tabla `usuarios` y llama
    ///   ReEncryptAndPersist para guardar en el vault cifrado.
    ///
    /// DISEÑO:
    ///   Formulario simple sin dependencias de grid/DataGridView.
    ///   Compatible con el entorno Excel-DNA (STA thread, WinForms puro).
    /// </summary>
    internal sealed class FormRegistroSubunidad : Form
    {
        // ====================================================================
        // CONTROLES UI
        // ====================================================================

        private Label _labelInstrucciones;
        private Label _labelNombre;
        private TextBox _textBoxNombre;
        private Label _labelSubunidad;
        private ComboBox _comboBoxSubunidad;
        private Label _labelUsuarioWindows;
        private Button _buttonAceptar;
        private Button _buttonCancelar;
        private Panel _panelInfo;
        private Label _labelWarning;

        // ====================================================================
        // DATOS
        // ====================================================================

        /// <summary>
        /// Subunidad seleccionada por el usuario.
        /// Solo tiene valor si DialogResult = OK.
        /// </summary>
        internal string SubunidadSeleccionada { get; private set; }

        /// <summary>
        /// Nombre completo ingresado por el usuario.
        /// Solo tiene valor si DialogResult = OK.
        /// </summary>
        internal string NombreCompleto { get; private set; }

        /// <summary>
        /// Nombre de usuario de Windows (Environment.UserName).
        /// Se muestra de solo lectura para confirmación.
        /// </summary>
        private readonly string _windowsUserName;

        // ====================================================================
        // CONSTRUCTOR
        // ====================================================================

        internal FormRegistroSubunidad()
        {
            _windowsUserName = Environment.UserName;
            InitializeComponent();
            CenterToScreen();
        }

        // ====================================================================
        // INICIALIZACIÓN DE COMPONENTES (manual — sin Designer)
        // ====================================================================

        /// <summary>
        /// Construye la interfaz manualmente (sin archivo .Designer.cs).
        ///
        /// DECISIÓN: Se usa código manual en lugar de Designer para mantener
        /// el formulario en un solo archivo y facilitar la revisión de código.
        /// Al ser un formulario simple y puntual, el overhead es mínimo.
        /// </summary>
        private void InitializeComponent()
        {
            // ----------------------------------------------------------------
            // Configuración del formulario
            // ----------------------------------------------------------------
            this.Text = "ConversorPDF — Registro de Subunidad";
            this.Size = new System.Drawing.Size(480, 360);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new System.Drawing.Font("Segoe UI", 9f);
            this.AutoScroll = true;
            // TopMost garantiza visibilidad cuando el formulario se muestra desde un hilo STA
            // de background (ExcelAddIn.AutoOpen). Sin esto, Excel mantiene el foco y el
            // diálogo queda oculto detrás de la ventana de Excel, bloqueando la inicialización.
            this.TopMost = true;

            // ----------------------------------------------------------------
            // Panel de advertencia/instrucción (banda superior)
            // ----------------------------------------------------------------
            _panelInfo = new Panel
            {
                BackColor = System.Drawing.Color.FromArgb(0, 114, 188),  // azul ESAP
                Dock = DockStyle.Top,
                Height = 60
            };

            _labelWarning = new Label
            {
                Text = "⚙  Primer uso en este equipo",
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            _panelInfo.Controls.Add(_labelWarning);

            // ----------------------------------------------------------------
            // Instrucciones
            // ----------------------------------------------------------------
            _labelInstrucciones = new Label
            {
                Text = "El usuario de Windows actual no está registrado en el sistema.\n" +
                       "Complete los datos a continuación para habilitar la sincronización.",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(440, 42),
                ForeColor = System.Drawing.Color.FromArgb(60, 60, 60)
            };

            // ----------------------------------------------------------------
            // Campo: Usuario Windows (solo lectura)
            // ----------------------------------------------------------------
            var labelUsuarioWindowsTitle = new Label
            {
                Text = "Usuario de Windows:",
                Location = new System.Drawing.Point(20, 135),
                Size = new System.Drawing.Size(140, 22),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            _labelUsuarioWindows = new Label
            {
                Text = _windowsUserName,
                Location = new System.Drawing.Point(165, 135),
                Size = new System.Drawing.Size(290, 22),
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 114, 188),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            // ----------------------------------------------------------------
            // Campo: Nombre completo
            // ----------------------------------------------------------------
            _labelNombre = new Label
            {
                Text = "Nombre completo:",
                Location = new System.Drawing.Point(20, 170),
                Size = new System.Drawing.Size(140, 22),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            _textBoxNombre = new TextBox
            {
                Location = new System.Drawing.Point(165, 170),
                Size = new System.Drawing.Size(290, 24),
                MaxLength = 100
            };

            // ----------------------------------------------------------------
            // Campo: Subunidad (dropdown)
            // ----------------------------------------------------------------
            _labelSubunidad = new Label
            {
                Text = "Subunidad territorial:",
                Location = new System.Drawing.Point(20, 210),
                Size = new System.Drawing.Size(140, 22),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            _comboBoxSubunidad = new ComboBox
            {
                Location = new System.Drawing.Point(165, 210),
                Size = new System.Drawing.Size(200, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var subunidades = new System.Collections.Generic.List<string>();
            try
            {
                // Fase 64b FIX: Usar GetContextForPath() en lugar de GetContext()
                // GetContext() retorna un Singleton. Si se envuelve en 'using', se desecha 
                // para toda la aplicación causando ObjectDisposedException.
                using (var context = ConversorPDF.XLL.Core.Helpers.SqliteHelper.GetContextForPath())
                {
                    var db = context.Open();
                    var result = Dapper.SqlMapper.Query<string>(db, "SELECT DISTINCT subunidad FROM territoriales WHERE subunidad IS NOT NULL AND subunidad <> '' ORDER BY subunidad").ToList();
                    subunidades.AddRange(result);
                }
            }
            catch (Exception ex)
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Warning($"[FormRegistroSubunidad] Error cargando subunidades: {ex.Message}");
            }
            if (subunidades.Count == 0) subunidades.Add("ESAPCENTRAL"); // Fallback

            _comboBoxSubunidad.Items.AddRange(subunidades.ToArray());
            _comboBoxSubunidad.SelectedIndex = 0;  // Primer elemento por defecto

            // ----------------------------------------------------------------
            // Botones
            // ----------------------------------------------------------------
            _buttonCancelar = new Button
            {
                Text = "Cancelar (sin sync)",
                Location = new System.Drawing.Point(20, 280),
                Size = new System.Drawing.Size(160, 34),
                DialogResult = DialogResult.Cancel,
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
            };

            _buttonAceptar = new Button
            {
                Text = "Registrarme",
                Location = new System.Drawing.Point(300, 280),
                Size = new System.Drawing.Size(155, 34),
                DialogResult = DialogResult.None,  // Manejado manualmente en Click
                BackColor = System.Drawing.Color.FromArgb(0, 114, 188),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold)
            };

            this.AcceptButton = _buttonAceptar;
            this.CancelButton = _buttonCancelar;

            // ----------------------------------------------------------------
            // Eventos
            // ----------------------------------------------------------------
            _buttonAceptar.Click += OnAceptarClick;

            // ----------------------------------------------------------------
            // Agregar controles al formulario
            // ----------------------------------------------------------------
            this.Controls.Add(_panelInfo);
            this.Controls.Add(_labelInstrucciones);
            this.Controls.Add(labelUsuarioWindowsTitle);
            this.Controls.Add(_labelUsuarioWindows);
            this.Controls.Add(_labelNombre);
            this.Controls.Add(_textBoxNombre);
            this.Controls.Add(_labelSubunidad);
            this.Controls.Add(_comboBoxSubunidad);
            this.Controls.Add(_buttonCancelar);
            this.Controls.Add(_buttonAceptar);
        }

        // ====================================================================
        // MANEJADORES DE EVENTOS
        // ====================================================================

        /// <summary>
        /// Valida los campos y cierra el formulario si todo está correcto.
        /// </summary>
        private void OnAceptarClick(object sender, EventArgs e)
        {
            // Validar nombre
            string nombre = _textBoxNombre.Text.Trim();
            if (string.IsNullOrWhiteSpace(nombre))
            {
                MessageBox.Show(
                    "Por favor ingrese su nombre completo.",
                    "Campo requerido",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                _textBoxNombre.Focus();
                return;
            }

            // Validar subunidad seleccionada
            if (_comboBoxSubunidad.SelectedIndex < 0)
            {
                MessageBox.Show(
                    "Por favor seleccione la subunidad territorial a la que pertenece.",
                    "Campo requerido",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                _comboBoxSubunidad.Focus();
                return;
            }

            // Guardar los valores seleccionados
            NombreCompleto = nombre;
            SubunidadSeleccionada = _comboBoxSubunidad.SelectedItem?.ToString();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
