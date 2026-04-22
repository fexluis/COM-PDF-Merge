using System;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Data;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Services;
namespace ConversorPDF.XLL.FormProgramador
{
    public partial class FormTableEditor : Form
    {
        private readonly string _tableName;
        private readonly IViaticosService _viaticosService;
        private System.Data.DataTable _dataTable;

        public FormTableEditor(string tableName, string title, int wide, IViaticosService viaticosService)
        {
            _viaticosService = viaticosService;
            InitializeComponent();

            this.Width = wide;
            _tableName = tableName;
            this.Text = title;

            this.Load += FormTableEditor_Load;
            this.btnGuardar.Click += BtnGuardar_Click;
            this.btnCancelar.Click += BtnCancelar_Click;
        }

        private void FormTableEditor_Load(object sender, EventArgs e)
        {
            try
            {
                object[,] data = _viaticosService.GetTableData(_tableName);
                _dataTable = new System.Data.DataTable();


                if (data != null)
                {
                    int rows = data.GetLength(0);
                    int cols = data.GetLength(1);

                    // Crear columnas dinámicas basadas en el número de columnas devueltas
                    for (int c = 1; c <= cols; c++)
                    {
                        _dataTable.Columns.Add("Columna " + c);
                    }

                    for (int r = 1; r <= rows; r++)
                    {
                        System.Data.DataRow row = _dataTable.NewRow();
                        for (int c = 1; c <= cols; c++)
                        {
                            row[c - 1] = data[r, c];
                        }
                        _dataTable.Rows.Add(row);
                    }
                }
                else
                {
                    // Si no hay datos, al menos crear una columna para Categoría 1 o 3 para Destinos
                    int defaultCols = _tableName == "tblCategoria1" ? 1 : 3;
                    for (int c = 1; c <= defaultCols; c++)
                    {
                        _dataTable.Columns.Add("Columna " + c);
                    }
                }

                dgvData.DataSource = _dataTable;

                // Ajustar encabezados si es posible

                if (_tableName == "tblDestinos" && dgvData.Columns.Count >= 3)
                {
                    dgvData.Columns[0].HeaderText = "DEPARTAMENTO";
                    dgvData.Columns[1].HeaderText = "CIUDAD ERRADA";
                    dgvData.Columns[2].HeaderText = "CIUDAD CORRECTA";
                }
                else if (_tableName == "tblCategoria1" && dgvData.Columns.Count >= 1)
                {
                    dgvData.Columns[0].HeaderText = "MUNICIPIO CATEGORIA 1";
                }
                else if (_tableName == "tblDependencias" && dgvData.Columns.Count >= 9)
                {
                    dgvData.Columns[0].HeaderText = "DEPENDENCIA";
                    dgvData.Columns[1].HeaderText = "DEPENDENCIA DESCRIPCION";
                    dgvData.Columns[2].HeaderText = "CDP";
                    dgvData.Columns[3].HeaderText = "FECHA CDP";
                    dgvData.Columns[4].HeaderText = "RUBRO";
                    dgvData.Columns[5].HeaderText = "RESPONSABLE";
                    dgvData.Columns[6].HeaderText = "CARGO";
                    dgvData.Columns[7].HeaderText = "ELABORÓ";
                    dgvData.Columns[8].HeaderText = "REVISÓ";
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error cargando datos en FormTableEditor para {_tableName}", ex);
                MessageBox.Show("No se pudo cargar la tabla. Verifique que existe en la hoja Dependencias.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                dgvData.EndEdit();
                    int rows = _dataTable.Rows.Count;
                    int cols = _dataTable.Columns.Count;
                    object[,] data = new object[rows, cols];
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            data[r, c] = _dataTable.Rows[r][c];
                        }
                    }
                    _viaticosService.SaveTableData(_tableName, data);



                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"Error guardando datos en FormTableEditor para {_tableName}", ex);
                MessageBox.Show("Error al guardar los cambios.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}