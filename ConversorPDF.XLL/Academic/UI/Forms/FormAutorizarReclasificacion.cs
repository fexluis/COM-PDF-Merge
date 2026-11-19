using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Diálogo modal para la selección y autorización individual de reclasificaciones de patrones de aula (Fase 1380).
    /// Permite al usuario autorizar explícitamente qué materias pueden redistribuir días, manteniendo desmarcadas por defecto.
    /// </summary>
    public class FormAutorizarReclasificacion : Form
    {
        private readonly IReadOnlyList<AgReclasificacionPatron> _candidatos;
        private CheckedListBox _chkLista;
        private Button _btnAceptar;
        private Button _btnCancelar;
        private Button _btnTodos;
        private Button _btnNinguno;
        private Label _lblDescripcion;

        public FormAutorizarReclasificacion(IReadOnlyList<AgReclasificacionPatron> candidatos)
        {
            _candidatos = candidatos ?? new List<AgReclasificacionPatron>();
            InitializeComponent();
            PoblarLista();
        }

        private void InitializeComponent()
        {
            this.Text = "Autorización de Reclasificación de Patrones";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(680, 420);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

            _lblDescripcion = new Label
            {
                Text = "Seleccione individualmente las asignaturas cuya restricción de días desea reclasificar a AUTOMATICO para redistribución. Las asignaturas no marcadas conservarán su patrón intacto:",
                Location = new Point(16, 12),
                Size = new Size(648, 40),
                AutoEllipsis = true
            };

            _chkLista = new CheckedListBox
            {
                Location = new Point(16, 56),
                Size = new Size(648, 290),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };

            _btnTodos = new Button
            {
                Text = "Marcar Todos",
                Location = new Point(16, 356),
                Size = new Size(110, 28),
                UseVisualStyleBackColor = true
            };
            _btnTodos.Click += (s, e) => MarcarTodos(true);

            _btnNinguno = new Button
            {
                Text = "Desmarcar Todos",
                Location = new Point(132, 356),
                Size = new Size(120, 28),
                UseVisualStyleBackColor = true
            };
            _btnNinguno.Click += (s, e) => MarcarTodos(false);

            _btnAceptar = new Button
            {
                Text = "Aceptar",
                DialogResult = DialogResult.OK,
                Location = new Point(486, 356),
                Size = new Size(86, 28),
                UseVisualStyleBackColor = true
            };
            _btnAceptar.Click += (s, e) => ConfirmarSeleccion();

            _btnCancelar = new Button
            {
                Text = "Cancelar",
                DialogResult = DialogResult.Cancel,
                Location = new Point(578, 356),
                Size = new Size(86, 28),
                UseVisualStyleBackColor = true
            };

            this.Controls.Add(_lblDescripcion);
            this.Controls.Add(_chkLista);
            this.Controls.Add(_btnTodos);
            this.Controls.Add(_btnNinguno);
            this.Controls.Add(_btnAceptar);
            this.Controls.Add(_btnCancelar);

            this.AcceptButton = _btnAceptar;
            this.CancelButton = _btnCancelar;
        }

        private void PoblarLista()
        {
            _chkLista.Items.Clear();
            for (int i = 0; i < _candidatos.Count; i++)
            {
                var c = _candidatos[i];
                if (c == null)
                {
                    _chkLista.Items.Add("(Elemento nulo)", false);
                    continue;
                }

                string ubicacionBase;
                if (!string.IsNullOrWhiteSpace(c.Sede) && !string.IsNullOrWhiteSpace(c.Programa))
                {
                    ubicacionBase = $"{c.Sede} - {c.Programa}";
                }
                else if (!string.IsNullOrWhiteSpace(c.Sede))
                {
                    ubicacionBase = c.Sede;
                }
                else if (!string.IsNullOrWhiteSpace(c.Programa))
                {
                    ubicacionBase = c.Programa;
                }
                else
                {
                    ubicacionBase = c.AulaRef;
                }

                string infoUbicacion = !string.IsNullOrWhiteSpace(ubicacionBase)
                    ? $"[{ubicacionBase}] "
                    : string.Empty;

                string infoGrupo = !string.IsNullOrWhiteSpace(c.SemestreGrupo)
                    ? $" [{c.SemestreGrupo}]"
                    : string.Empty;

                // Fase 2122: Rotular 'En riesgo' en vez de 'Déficit' para cotas superiores sin plan factible
                string infoDeficit = c.SesionesPendientes > 0
                    ? $" | En riesgo: {c.SesionesPendientes} ses ({c.HorasPendientes:0.#}h)"
                    : string.Empty;

                string diasAnt = !string.IsNullOrWhiteSpace(c.DiasPatronAnterior)
                    ? c.DiasPatronAnterior
                    : "N/D";

                string itemTexto = $"{infoUbicacion}{c.NombreAsignatura}{infoGrupo} (Días: {diasAnt} -> {c.OrigenPropuesto}){infoDeficit}";
                // Inicialmente desmarcado salvo que ya estuviera autorizado explícitamente
                _chkLista.Items.Add(itemTexto, c.Autorizado);
            }
        }

        private void MarcarTodos(bool marcar)
        {
            for (int i = 0; i < _chkLista.Items.Count; i++)
            {
                _chkLista.SetItemChecked(i, marcar);
            }
        }

        private void ConfirmarSeleccion()
        {
            for (int i = 0; i < _candidatos.Count; i++)
            {
                if (_candidatos[i] != null && i < _chkLista.Items.Count)
                {
                    _candidatos[i].Autorizado = _chkLista.GetItemChecked(i);
                }
            }
        }
    }
}

