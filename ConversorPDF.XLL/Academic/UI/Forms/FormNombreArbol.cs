using System;
using System.Drawing;
using System.Windows.Forms;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Diálogo modal liviano (code-only) para solicitar el nombre de un perfil de árbol académico (Fase 240).
    /// </summary>
    public class FormNombreArbol : Form
    {
        private TextBox _txtNombre;
        private Button _btnAceptar;
        private Button _btnCancelar;
        private Label _lblTitulo;

        public string NombrePerfil => _txtNombre?.Text?.Trim() ?? string.Empty;

        public FormNombreArbol(string nombreSugerido = "")
        {
            InitializeComponent(nombreSugerido);
        }

        private void InitializeComponent(string nombreSugerido)
        {
            this.Text = "Guardar Perfil de Selección de Árbol";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(380, 140);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

            _lblTitulo = new Label
            {
                Text = "Ingrese el nombre para el perfil de selección:",
                Location = new Point(16, 16),
                AutoSize = true
            };

            _txtNombre = new TextBox
            {
                Location = new Point(18, 42),
                Size = new Size(344, 23),
                Text = nombreSugerido ?? string.Empty
            };

            _btnAceptar = new Button
            {
                Text = "Aceptar",
                DialogResult = DialogResult.OK,
                Location = new Point(190, 85),
                Size = new Size(82, 28),
                UseVisualStyleBackColor = true
            };
            _btnAceptar.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_txtNombre.Text))
                {
                    MessageBox.Show(this, "El nombre del perfil no puede estar vacío.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.DialogResult = DialogResult.None;
                }
            };

            _btnCancelar = new Button
            {
                Text = "Cancelar",
                DialogResult = DialogResult.Cancel,
                Location = new Point(280, 85),
                Size = new Size(82, 28),
                UseVisualStyleBackColor = true
            };

            this.Controls.Add(_lblTitulo);
            this.Controls.Add(_txtNombre);
            this.Controls.Add(_btnAceptar);
            this.Controls.Add(_btnCancelar);

            this.AcceptButton = _btnAceptar;
            this.CancelButton = _btnCancelar;
        }
    }
}
