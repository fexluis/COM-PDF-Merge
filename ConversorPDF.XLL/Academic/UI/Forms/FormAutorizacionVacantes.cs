using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Diálogo modal para la selección y autorización explícita de vacantes no consentidas antes de la promoción a producción (Fase 2037).
    /// Evita la autoautorización indiscriminada obligando al usuario a seleccionar qué asignaturas vacantes aprueba promover.
    /// </summary>
    public class FormAutorizacionVacantes : Form
    {
        private readonly List<ItemVacanteDto> _items;
        private CheckedListBox _chkLista;
        private Button _btnAceptar;
        private Button _btnCancelar;
        private Button _btnTodos;
        private Button _btnNinguno;
        private Label _lblDescripcion;

        public ISet<int> Autorizadas { get; private set; } = new HashSet<int>();
        public ISet<string> AutorizadasClaves { get; private set; } = new HashSet<string>();

        public FormAutorizacionVacantes(AgIndividuo individuo, AgContexto contexto = null)
        {
            _items = ObtenerItemsVacantes(individuo, contexto);
            InitializeComponent();
            PoblarLista();
        }

        private static List<ItemVacanteDto> ObtenerItemsVacantes(AgIndividuo individuo, AgContexto contexto)
        {
            var res = new List<ItemVacanteDto>();
            if (individuo?.Genes == null) return res;

            var genesVacantes = individuo.Genes
                .Where(g => g.EsVacante == 1 || g.IdTercero <= 0)
                .Where(g => contexto == null || !contexto.EstaConsentida(g.IdAsignatura, g.IdSede))
                .GroupBy(g => (g.IdAsignatura, g.IdSede));

            foreach (var grp in genesVacantes)
            {
                int idAsig = grp.Key.IdAsignatura;
                int idSede = grp.Key.IdSede;
                string nombre = $"Asignatura #{idAsig}";
                if (contexto?.AsignaturasPorId != null && contexto.AsignaturasPorId.TryGetValue(idAsig, out var asig) && !string.IsNullOrWhiteSpace(asig.Nombre))
                {
                    nombre = $"{asig.Nombre} (ID {idAsig})";
                }
                string sedeDesc = idSede > 0 ? $" [Sede {idSede}]" : string.Empty;
                int count = grp.Count();
                res.Add(new ItemVacanteDto
                {
                    IdAsignatura = idAsig,
                    IdSede = idSede,
                    Texto = $"{nombre}{sedeDesc} — {count} sesión(es) vacante(s)"
                });
            }

            return res.OrderBy(i => i.IdAsignatura).ThenBy(i => i.IdSede).ToList();
        }

        private void InitializeComponent()
        {
            this.Text = "Autorización de Vacantes para Promoción";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(620, 380);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

            _lblDescripcion = new Label
            {
                Text = "Seleccione explícitamente las asignaturas con vacantes que autoriza promover a producción sin docente asignado. Las no seleccionadas bloquearán la promoción.",
                Location = new Point(16, 12),
                Size = new Size(588, 36),
                AutoEllipsis = true
            };

            _chkLista = new CheckedListBox
            {
                Location = new Point(16, 52),
                Size = new Size(588, 260),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };

            _btnTodos = new Button
            {
                Text = "Marcar Todas",
                Location = new Point(16, 330),
                Size = new Size(110, 30)
            };
            _btnTodos.Click += (s, e) => CambiarSeleccion(true);

            _btnNinguno = new Button
            {
                Text = "Desmarcar",
                Location = new Point(132, 330),
                Size = new Size(100, 30)
            };
            _btnNinguno.Click += (s, e) => CambiarSeleccion(false);

            _btnAceptar = new Button
            {
                Text = "Autorizar",
                DialogResult = DialogResult.OK,
                Location = new Point(410, 330),
                Size = new Size(94, 30)
            };
            _btnAceptar.Click += BtnAceptar_Click;

            _btnCancelar = new Button
            {
                Text = "Cancelar",
                DialogResult = DialogResult.Cancel,
                Location = new Point(510, 330),
                Size = new Size(94, 30)
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
            foreach (var item in _items)
            {
                _chkLista.Items.Add(item, false);
            }
        }

        private void CambiarSeleccion(bool estado)
        {
            for (int i = 0; i < _chkLista.Items.Count; i++)
            {
                _chkLista.SetItemChecked(i, estado);
            }
        }

        private void BtnAceptar_Click(object sender, EventArgs e)
        {
            Autorizadas = new HashSet<int>();
            AutorizadasClaves = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in _chkLista.CheckedItems)
            {
                if (item is ItemVacanteDto dto)
                {
                    if (dto.IdSede > 0)
                    {
                        AutorizadasClaves.Add(dto.Clave);
                    }
                    else
                    {
                        Autorizadas.Add(dto.IdAsignatura);
                    }
                }
            }
        }
    }
}
