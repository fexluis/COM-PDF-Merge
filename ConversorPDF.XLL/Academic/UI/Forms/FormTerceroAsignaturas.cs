using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Formulario modal para el alta y edición de asignaturas habilitadas para un docente en el pool.
    /// </summary>
    public partial class FormTerceroAsignaturas : Form
    {
        private readonly int _idTercero;
        private readonly TerceroAsignatura _originalEntity;
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly ITerceroAsignaturaRepository _repository;

        private List<Programa> _programas = new List<Programa>();
        private List<Asignatura> _asignaturas = new List<Asignatura>();

        public FormTerceroAsignaturas(int idTercero, TerceroAsignatura item, Func<IDbConnection> connectionFactory)
        {
            _idTercero = idTercero;
            _originalEntity = item;
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _repository = new TerceroAsignaturaRepository(connectionFactory);

            InitializeComponent();
            InicializarSemestres();
            CargarProgramas();
            CargarDatos();
        }

        private void InicializarSemestres()
        {
            cmbSemestre.Items.Clear();
            cmbSemestre.Items.Add("(todos)");
            string[] semestres = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" };
            foreach (var s in semestres)
            {
                cmbSemestre.Items.Add(s);
            }
            cmbSemestre.SelectedIndex = 0;
        }

        private void CargarProgramas()
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    // Fase 1470: catalogo canonico compartido con el AG (sin filtro activo, COALESCE defensivo).
                    _programas = db.Query<Programa>(AgProgramaCatalogoSql.SeleccionProgramas).ToList();

                    cmbPrograma.DisplayMember = "Nombre";
                    cmbPrograma.ValueMember = "Id";
                    cmbPrograma.Format -= CmbPrograma_Format;
                    cmbPrograma.Format += CmbPrograma_Format;
                    cmbPrograma.DataSource = _programas;

                    int pregrados = _programas.Count(p => AgTipoProgramaCatalog.Nivel(p.Modalidad) == AgNivelAcademico.Pregrado);
                    int posgrados = _programas.Count(p => AgTipoProgramaCatalog.Nivel(p.Modalidad) == AgNivelAcademico.Posgrado);
                    LoggerHelper.Info($"[FormTerceroAsignaturas] programas={_programas.Count} (pregrado={pregrados}, posgrado={posgrados})");

                    if (_programas.Count == 0)
                    {
                        LoggerHelper.Warning("[FormTerceroAsignaturas] Catalogo de programas vacio: no se podran habilitar asignaturas.");
                        MessageBox.Show("No se encontro el catalogo de programas academicos.", "Catalogo vacio",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormTerceroAsignaturas] Error al cargar programas", ex);
                MessageBox.Show($"Error al cargar programas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Fase 1470: muestra "ABREV [NIVEL]" para distinguir pregrado/posgrado en el combo.
        /// </summary>
        private void CmbPrograma_Format(object sender, ListControlConvertEventArgs e)
        {
            if (!(e.ListItem is Programa p)) return;

            string etiqueta = AgTipoProgramaCatalog.EtiquetaPrograma(p);
            string nivel = AgTipoProgramaCatalog.Etiqueta(AgTipoProgramaCatalog.Nivel(p.Modalidad));
            e.Value = $"{etiqueta} [{nivel}]";
        }

        private void CargarAsignaturas()
        {
            if (cmbPrograma.SelectedValue == null || !(cmbPrograma.SelectedValue is int idPrograma))
            {
                cmbAsignatura.DataSource = null;
                return;
            }

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string semestreFiltro = cmbSemestre.SelectedItem?.ToString() ?? string.Empty;
                    string sql = "SELECT id as Id, nombre as Nombre, semestre as Semestre, num_nivel as NumNivel, abreviatura as Abreviatura FROM asignaturas WHERE idprograma = @idPrograma";
                    if (!string.IsNullOrWhiteSpace(semestreFiltro) && semestreFiltro != "(todos)")
                    {
                        sql += " AND semestre = @semestreFiltro";
                    }
                    sql += " ORDER BY num_nivel, nombre";

                    _asignaturas = db.Query<Asignatura>(sql, new { idPrograma, semestreFiltro }).ToList();
                    cmbAsignatura.DisplayMember = "Nombre";
                    cmbAsignatura.ValueMember = "Id";
                    cmbAsignatura.DataSource = _asignaturas;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[FormTerceroAsignaturas] Error al cargar asignaturas", ex);
                MessageBox.Show($"Error al cargar asignaturas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarDatos()
        {
            if (_originalEntity != null)
            {
                this.Text = "Editar Asignatura Habilitada";
                if (_originalEntity.IdPrograma > 0)
                {
                    cmbPrograma.SelectedValue = _originalEntity.IdPrograma;

                    // Fase 1470: si el programa historico no existe en el catalogo, se agrega un item sintetico
                    // para no perder la edicion de la habilitacion existente.
                    if (!(_programas.Any(p => p.Id == _originalEntity.IdPrograma)))
                    {
                        LoggerHelper.Warning($"[FormTerceroAsignaturas] Programa {_originalEntity.IdPrograma} no esta en el catalogo; se agrega como referencia.");
                        _programas.Insert(0, new Programa
                        {
                            Id = _originalEntity.IdPrograma,
                            NombrePrograma = $"Programa #{_originalEntity.IdPrograma} (no catalogado)",
                            Modalidad = string.Empty
                        });
                        cmbPrograma.DataSource = null;
                        cmbPrograma.DataSource = _programas;
                        cmbPrograma.SelectedValue = _originalEntity.IdPrograma;
                    }
                }

                if (!string.IsNullOrWhiteSpace(_originalEntity.Semestre))
                {
                    int index = cmbSemestre.FindStringExact(_originalEntity.Semestre);
                    if (index >= 0) cmbSemestre.SelectedIndex = index;
                }

                CargarAsignaturas();
                if (_originalEntity.IdAsignatura > 0)
                {
                    cmbAsignatura.SelectedValue = _originalEntity.IdAsignatura;
                }

                chkActivo.Checked = _originalEntity.Activo == 1;
                numPrioridad.Value = Math.Min(numPrioridad.Maximum, Math.Max(numPrioridad.Minimum, _originalEntity.Prioridad));
                lblVecesAsignadaValor.Text = _originalEntity.VecesAsignada.ToString();
                lblUltimaAsignacionValor.Text = string.IsNullOrWhiteSpace(_originalEntity.FechaUltimaAsignacion) ? "N/A" : _originalEntity.FechaUltimaAsignacion;
                txtOrigen.Text = string.IsNullOrWhiteSpace(_originalEntity.Origen) ? "MANUAL" : _originalEntity.Origen;
                txtObservaciones.Text = _originalEntity.Observaciones;
            }
            else
            {
                this.Text = "Habilitar Nueva Asignatura";
                CargarAsignaturas();
            }
        }

        private void CmbPrograma_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarAsignaturas();
        }

        private void CmbSemestre_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarAsignaturas();
        }

        private async void BtnGuardar_Click(object sender, EventArgs e)
        {
            if (cmbPrograma.SelectedValue == null || !(cmbPrograma.SelectedValue is int idPrograma) || idPrograma <= 0)
            {
                MessageBox.Show("Debe seleccionar un programa académico válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbAsignatura.SelectedValue == null || !(cmbAsignatura.SelectedValue is int idAsignatura) || idAsignatura <= 0)
            {
                MessageBox.Show("Debe seleccionar una asignatura válida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string semestre = cmbSemestre.SelectedItem?.ToString() ?? string.Empty;
            if (semestre == "(todos)") semestre = string.Empty;

            long nuevoId = TerceroAsignaturaKey.Calcular(_idTercero, idPrograma, idAsignatura, semestre);

            var entity = new TerceroAsignatura
            {
                Id = nuevoId,
                IdTercero = _idTercero,
                IdPrograma = idPrograma,
                IdAsignatura = idAsignatura,
                Semestre = semestre,
                Activo = chkActivo.Checked ? 1 : 0,
                Prioridad = (int)numPrioridad.Value,
                VecesAsignada = _originalEntity?.VecesAsignada ?? 0,
                FechaUltimaAsignacion = _originalEntity?.FechaUltimaAsignacion ?? string.Empty,
                Origen = _originalEntity?.Origen ?? "MANUAL",
                Observaciones = txtObservaciones.Text.Trim(),
                Subunidad = _originalEntity?.Subunidad ?? string.Empty
            };

            // Si cambió la clave natural respecto al registro original, se elimina el ID anterior
            if (_originalEntity != null && _originalEntity.Id != 0 && _originalEntity.Id != nuevoId)
            {
                var deleteResult = await _repository.DeleteAsync(_originalEntity.Id, CancellationToken.None);
                if (deleteResult.IsFailure)
                {
                    LoggerHelper.Warning($"[FormTerceroAsignaturas] No se pudo eliminar la clave anterior {_originalEntity.Id}: {deleteResult.Error}");
                }
            }

            var upsertResult = await _repository.UpsertAsync(entity, CancellationToken.None);
            if (upsertResult.IsFailure)
            {
                MessageBox.Show($"Error al guardar habilitación docente: {upsertResult.Error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
