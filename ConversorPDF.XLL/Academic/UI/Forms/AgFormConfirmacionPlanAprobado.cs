using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>Ratificación humana del origen reconciliado y, opcionalmente, de un único paso candidato.</summary>
    internal sealed class AgFormConfirmacionPlanAprobado : Form
    {
        private readonly TextBox _motivo = new TextBox { Multiline = true, Height = 58, Dock = DockStyle.Top, MaxLength = 1000 };
        private readonly CheckBox _ratifica = new CheckBox { AutoSize = true, Text = "He revisado y ratifico el horario histórico/manual reconciliado mostrado" };
        private readonly CheckBox _paso = new CheckBox { AutoSize = true, Text = "Autorizo únicamente este siguiente paso para el aula completa (sin modificar producción)" };
        private readonly List<CheckBox> _legado = new List<CheckBox>();
        internal bool Ratifica => _ratifica.Checked;
        internal bool AutorizaPaso => _paso.Checked;
        internal string Motivo => _motivo.Text.Trim();
        internal IReadOnlyCollection<int> LegadoConsentido => _legado.Where(c => c.Checked).Select(c => (int)c.Tag).ToArray();

        internal AgFormConfirmacionPlanAprobado(int idAula, AgPlanAprobadoSnapshot snapshot,
            AgPropuestaEscaladoAprobada propuesta, string detalle, AgCandidatoPlanAprobado archivada = null)
        {
            Text = "Ratificar origen y confirmar candidato — aula " + idAula;
            Width = 940; Height = 720; StartPosition = FormStartPosition.CenterParent;
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
                WrapContents = false, AutoScroll = true };
            var texto = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Width = 865, Height = 340, Text = detalle + Environment.NewLine +
                    "Origen reconciliado: " + snapshot.HuellaOrigen + Environment.NewLine +
                    string.Join(Environment.NewLine, snapshot.Relaciones.Where(r => r.ClavesCandidatas.Count == 1
                        && r.ClavesCandidatas[0].IdAula == idAula).Select(r =>
                        "Vinculación " + r.IdVinculacion + ": " + r.ClavesCandidatas[0].Valor +
                        " | sesiones " + string.Join(", ", r.Sesiones.Select(s =>
                            s.IdSesion + " " + s.Fecha + " " + s.HoraInicio + "–" + s.HoraFin + " " + s.HorasLectivas + "h")))) };
            panel.Controls.Add(texto);
            panel.Controls.Add(new Label { Text = "Motivo obligatorio de la ratificación y revisión:", AutoSize = true });
            _motivo.Width = 850;
            if (archivada != null) _motivo.Text = archivada.Motivo;
            panel.Controls.Add(_motivo);
            panel.Controls.Add(_ratifica);
            if (propuesta != null)
            {
                panel.Controls.Add(_paso);
                foreach (var r in propuesta.Relaciones.Where(r => r.RedistribuyeLegado && r.OrigenPatron == "LEGADO"))
                {
                    var opcion = new CheckBox { AutoSize = true, Tag = r.Clave.IdAsignatura,
                        Text = "Consiento redistribuir LEGADO asignatura " + r.Clave.IdAsignatura +
                            " (" + r.Clave.Valor + ")" };
                    if (archivada != null) opcion.Checked = archivada.Consentimiento.Split(',')
                        .Contains(r.Clave.IdAsignatura.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    _legado.Add(opcion); panel.Controls.Add(opcion);
                }
                panel.Controls.Add(new Label { AutoSize = true, Text = "FIJO nunca se convierte. Si el paso no cubre, I5 y siguiente oferta sólo en otra corrida." });
            }
            var aceptar = new Button { Text = "Confirmar ratificación", AutoSize = true, DialogResult = DialogResult.None };
            aceptar.Click += (s, e) =>
            {
                if (Motivo.Length == 0 || !_ratifica.Checked || (propuesta != null && !_paso.Checked))
                {
                    MessageBox.Show(this, "Motivo, ratificación y autorización explícita del paso son obligatorios.");
                    return;
                }
                DialogResult = DialogResult.OK; Close();
            };
            panel.Controls.Add(aceptar);
            var cancelar = new Button { Text = "Rechazar / cancelar", AutoSize = true, DialogResult = DialogResult.Cancel };
            panel.Controls.Add(cancelar);
            Controls.Add(panel);
            CancelButton = cancelar;
        }
    }
}
