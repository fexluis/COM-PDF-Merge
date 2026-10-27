using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ConversorPDF.XLL.Academic.UI.Controls
{
    public enum CheckedComboModo
    {
        Multiple,
        Unico
    }

    /// <summary>
    /// ComboBox con selección múltiple o única mediante checkboxes en un dropdown flotante (Fase 1312 / 1386).
    /// </summary>
    public sealed class CheckedComboBox : UserControl
    {
        private TextBox _txtDisplay;
        private Button _btnDropDown;
        private ToolStripDropDown _dropDown;
        private ToolStripControlHost _host;
        private CheckedListBox _chkList;

        private sealed class ItemFiltro
        {
            public int Id { get; set; }
            public string Texto { get; set; }
            public object Tag { get; set; }
            public bool Marcado { get; set; }
            public bool EsTodos { get; set; }

            public override string ToString() => Texto;
        }

        private readonly List<ItemFiltro> _items = new List<ItemFiltro>();
        private bool _suprimiendoEventos;
        private bool _sincronizandoLista;
        private string _firmaSeleccion;

        public CheckedComboModo Modo { get; set; } = CheckedComboModo.Multiple;
        public bool VacioEsTodos { get; set; } = true;

        /// <summary>Indica si se debe incluir el pseudo-ítem "Todos" en la lista.</summary>
        public bool IncluirItemTodos { get; set; } = false;

        /// <summary>Texto desplegado para el pseudo-ítem "Todos".</summary>
        public string TextoItemTodos { get; set; } = "-- Todos los Programas --";

        /// <summary>Texto desplegado en el cuadro de edición cuando no hay selección real o cuando "Todos" está marcado.</summary>
        public string TextoVacio { get; set; }

        private bool _todosMarcadoPorUsuario;

        /// <summary>Indica si el pseudo-ítem "Todos" fue marcado explícitamente por el usuario o sincronizado con el estado efectivo.</summary>
        public bool TodosMarcadoPorUsuario => _todosMarcadoPorUsuario;

        /// <summary>Texto actual renderizado en el cuadro de edición interno del control.</summary>
        public string TextoMostrado => _txtDisplay?.Text ?? string.Empty;

        public event EventHandler SeleccionCambiada;
        public event EventHandler SelectedIndexChanged;

        public int SelectedIndex
        {
            get => Primario();
            set
            {
                if (value <= 0)
                {
                    Marcar(Enumerable.Empty<int>());
                }
                else
                {
                    Marcar(new[] { value });
                }
            }
        }

        public object SelectedItem
        {
            get
            {
                var item = _items.FirstOrDefault(i => i.Marcado && !i.EsTodos);
                return item?.Tag ?? (item != null ? (object)item.Id : null);
            }
            set
            {
                if (value == null)
                {
                    Marcar(Enumerable.Empty<int>());
                }
                else if (value is int intVal)
                {
                    Marcar(new[] { intVal });
                }
                else
                {
                    var item = _items.FirstOrDefault(i => !i.EsTodos && Equals(i.Tag, value));
                    if (item != null)
                    {
                        Marcar(new[] { item.Id });
                    }
                }
            }
        }

        public CheckedComboBox()
        {
            InicializarComponentes();
        }

        private void InicializarComponentes()
        {
            this.Size = new Size(200, 23);
            this.Padding = new Padding(1);
            this.BackColor = Color.FromArgb(171, 171, 171); // Borde simulado

            _txtDisplay = new TextBox
            {
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = this.Font,
                Cursor = Cursors.Default
            };
            _txtDisplay.Click += (s, e) => AlternarDropDown();

            _btnDropDown = new Button
            {
                Width = 18,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Control,
                Text = "▼",
                Font = new Font("Segoe UI", 6.5f),
                Cursor = Cursors.Hand
            };
            _btnDropDown.FlatAppearance.BorderSize = 0;
            _btnDropDown.Click += (s, e) => AlternarDropDown();

            _chkList = new CheckedListBox
            {
                CheckOnClick = true,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                Font = this.Font
            };
            _chkList.ItemCheck += ChkList_ItemCheck;
            _chkList.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Escape)
                {
                    _dropDown.Close();
                    e.Handled = true;
                }
            };

            _host = new ToolStripControlHost(_chkList)
            {
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = false
            };

            _dropDown = new ToolStripDropDown
            {
                Padding = Padding.Empty,
                AutoClose = true
            };
            _dropDown.Items.Add(_host);
            _dropDown.Closed += (s, e) =>
            {
                RefrescarTexto();
                NotificarCambio();
            };

            this.Controls.Add(_txtDisplay);
            this.Controls.Add(_btnDropDown);
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            if (_txtDisplay != null) _txtDisplay.Font = this.Font;
            if (_chkList != null) _chkList.Font = this.Font;
        }

        private void AlternarDropDown()
        {
            if (_dropDown.Visible)
            {
                _dropDown.Close();
            }
            else
            {
                int ancho = Math.Max(this.Width, 220);
                int alto = Math.Min(Math.Max(_chkList.Items.Count * 22 + 4, 60), 250);
                _chkList.Size = new Size(ancho, alto);
                _host.Size = new Size(ancho, alto);
                _dropDown.Show(this, 0, this.Height);
            }
        }

        private void ChkList_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_suprimiendoEventos || _sincronizandoLista) return;

            if (Modo == CheckedComboModo.Unico && e.NewValue == CheckState.Checked)
            {
                _sincronizandoLista = true;
                for (int i = 0; i < _chkList.Items.Count; i++)
                {
                    if (i != e.Index && _chkList.GetItemChecked(i))
                    {
                        _chkList.SetItemChecked(i, false);
                        if (i < _items.Count) _items[i].Marcado = false;
                    }
                }
                _sincronizandoLista = false;
            }

            if (e.Index >= 0 && e.Index < _items.Count)
            {
                _items[e.Index].Marcado = (e.NewValue == CheckState.Checked);

                // Fase 1458: Registrar si "Todos" fue marcado explícitamente por el usuario
                if (IncluirItemTodos && _items[e.Index].EsTodos && e.NewValue == CheckState.Checked)
                {
                    _todosMarcadoPorUsuario = true;
                }
                if (!_items[e.Index].EsTodos && e.NewValue == CheckState.Checked)
                {
                    _todosMarcadoPorUsuario = false;
                }
            }

            // Fase 1447: Exclusión mutua entre pseudo-ítem Todos e ítems individuales
            if (IncluirItemTodos)
            {
                AplicarExclusionTodos(e.Index, e.NewValue == CheckState.Checked);
            }

            // Actualización de texto diferida para mostrar inmediatamente
            this.BeginInvoke((MethodInvoker)delegate
            {
                RefrescarTexto();
            });
        }

        /// <summary>
        /// Aplica la exclusión mutua entre el pseudo-ítem "Todos" y los ítems individuales reales.
        /// Fase 1447.
        /// </summary>
        internal void AplicarExclusionTodos(int index, bool marcado)
        {
            if (!IncluirItemTodos || _items.Count == 0 || index < 0 || index >= _items.Count) return;

            var item = _items[index];
            _sincronizandoLista = true;
            try
            {
                if (item.EsTodos)
                {
                    if (marcado)
                    {
                        for (int i = 0; i < _items.Count; i++)
                        {
                            if (!_items[i].EsTodos && _items[i].Marcado)
                            {
                                _items[i].Marcado = false;
                                _chkList.SetItemChecked(i, false);
                            }
                        }
                    }
                    else
                    {
                        // Si se intenta desmarcar "Todos" y ningún ítem real está marcado, mantener "Todos" marcado
                        if (!_items.Any(i => !i.EsTodos && i.Marcado))
                        {
                            item.Marcado = true;
                            _chkList.SetItemChecked(index, true);
                        }
                    }
                }
                else
                {
                    if (marcado)
                    {
                        for (int i = 0; i < _items.Count; i++)
                        {
                            if (_items[i].EsTodos && _items[i].Marcado)
                            {
                                _items[i].Marcado = false;
                                _chkList.SetItemChecked(i, false);
                            }
                        }
                    }
                    else
                    {
                        // Si se desmarcó el último ítem real, marcar "Todos" automáticamente
                        if (!_items.Any(i => !i.EsTodos && i.Marcado))
                        {
                            for (int i = 0; i < _items.Count; i++)
                            {
                                if (_items[i].EsTodos && !_items[i].Marcado)
                                {
                                    _items[i].Marcado = true;
                                    _chkList.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                }

                // Fase 1463: sincronizar intención con estado efectivo de selección
                _todosMarcadoPorUsuario = !_items.Any(i => !i.EsTodos && i.Marcado);
            }
            finally
            {
                _sincronizandoLista = false;
            }
        }

        private void NotificarCambio()
        {
            if (_suprimiendoEventos) return;
            string firma = string.Join(",", _items.Where(i => i.Marcado && !i.EsTodos).Select(i => i.Id));
            if (firma == _firmaSeleccion) return;
            _firmaSeleccion = firma;
            SeleccionCambiada?.Invoke(this, EventArgs.Empty);
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Cargar(IEnumerable<(int Id, string Texto)> items)
        {
            _suprimiendoEventos = true;
            _items.Clear();
            _chkList.Items.Clear();
            _firmaSeleccion = null; // Fase 1459: los ítems se reconstruyen; la firma previa ya no representa el estado
            _todosMarcadoPorUsuario = false; // Fase 1463

            if (IncluirItemTodos)
            {
                var itemTodos = new ItemFiltro { Id = 0, Texto = TextoItemTodos, Marcado = false, EsTodos = true };
                _items.Add(itemTodos);
                _chkList.Items.Add(itemTodos, false);
            }

            if (items != null)
            {
                foreach (var (id, texto) in items)
                {
                    var item = new ItemFiltro { Id = id, Texto = texto, Marcado = false };
                    _items.Add(item);
                    _chkList.Items.Add(item, false);
                }
            }

            if (IncluirItemTodos && _items.Count > 0)
            {
                _items[0].Marcado = true;
                _chkList.SetItemChecked(0, true);
            }

            _suprimiendoEventos = false;
            RefrescarTexto();
        }

        public void Cargar<T>(IEnumerable<T> items, Func<T, int> idSelector, Func<T, string> textoSelector)
        {
            Cargar(items, textoSelector, idSelector);
        }

        public void Cargar<T>(IEnumerable<T> items, Func<T, string> textoSelector, Func<T, int> idSelector)
        {
            if (items == null)
            {
                Cargar(Enumerable.Empty<(int, string)>());
                return;
            }

            _suprimiendoEventos = true;
            _items.Clear();
            _chkList.Items.Clear();
            _firmaSeleccion = null; // Fase 1459: los ítems se reconstruyen; la firma previa ya no representa el estado
            _todosMarcadoPorUsuario = false; // Fase 1463

            if (IncluirItemTodos)
            {
                var itemTodos = new ItemFiltro { Id = 0, Texto = TextoItemTodos, Tag = null, Marcado = false, EsTodos = true };
                _items.Add(itemTodos);
                _chkList.Items.Add(itemTodos, false);
            }

            foreach (var elem in items)
            {
                int id = idSelector != null ? idSelector(elem) : 0;
                string texto = textoSelector != null ? textoSelector(elem) : elem.ToString();
                var item = new ItemFiltro { Id = id, Texto = texto, Tag = elem, Marcado = false };
                _items.Add(item);
                _chkList.Items.Add(item, false);
            }

            if (IncluirItemTodos && _items.Count > 0)
            {
                _items[0].Marcado = true;
                _chkList.SetItemChecked(0, true);
            }

            _suprimiendoEventos = false;
            RefrescarTexto();
        }

        public void Marcar(IEnumerable<int> ids)
        {
            _suprimiendoEventos = true;
            _sincronizandoLista = true;
            _todosMarcadoPorUsuario = false; // Fase 1463: la selección programática anula la intención explícita de "Todos"
            var set = new HashSet<int>((ids ?? Enumerable.Empty<int>()).Where(id => id != 0));

            // Fase 1455: Si ningún id del set coincide con los ítems reales disponibles, marcar "Todos"
            bool marcarTodos = IncluirItemTodos &&
                (set.Count == 0 || !_items.Any(i => !i.EsTodos && set.Contains(i.Id)));

            for (int i = 0; i < _items.Count; i++)
            {
                bool marcar;
                if (_items[i].EsTodos)
                {
                    marcar = marcarTodos;
                }
                else
                {
                    marcar = !marcarTodos && set.Contains(_items[i].Id);
                    // Fase 1456: En modo único, solo marcar el primer ítem real encontrado
                    if (Modo == CheckedComboModo.Unico && marcar && set.Count > 1)
                    {
                        marcar = !_items.Take(i).Any(it => !it.EsTodos && it.Marcado);
                    }
                }

                _items[i].Marcado = marcar;
                _chkList.SetItemChecked(i, marcar);
            }

            _sincronizandoLista = false;
            _suprimiendoEventos = false;
            RefrescarTexto();
            NotificarCambio();
        }

        public void Marcar(object id)
        {
            if (id == null)
            {
                Limpiar();
                return;
            }
            try
            {
                int intId = Convert.ToInt32(id);
                Marcar(new[] { intId });
            }
            catch
            {
                // Si no es convertible a int, buscar por Tag
                var item = _items.FirstOrDefault(i => !i.EsTodos && Equals(i.Tag, id));
                if (item != null)
                {
                    Marcar(new[] { item.Id });
                }
            }
        }

        public void SeleccionarTodos(IEnumerable<int> ids)
        {
            Marcar(ids);
        }

        public void SeleccionarTodos()
        {
            if (Modo == CheckedComboModo.Unico) return;
            if (IncluirItemTodos)
            {
                Limpiar();
            }
            else
            {
                Marcar(_items.Where(i => !i.EsTodos).Select(i => i.Id));
            }
        }

        public void Limpiar()
        {
            Marcar(Enumerable.Empty<int>());
        }

        public List<int> Seleccionados()
        {
            if (_items.Count == 0) return new List<int>();

            var marcados = _items.Where(i => i.Marcado && !i.EsTodos).Select(i => i.Id).ToList();
            if (marcados.Count == 0 && VacioEsTodos)
            {
                return new List<int>(); // Vacio = Todos (sin filtro)
            }
            return marcados;
        }

        public List<T> Seleccionados<T>()
        {
            if (_items.Count == 0) return new List<T>();

            var marcados = _items.Where(i => i.Marcado && !i.EsTodos && i.Tag is T).Select(i => (T)i.Tag).ToList();
            return marcados;
        }

        public List<int> SeleccionadosIds()
        {
            if (_items.Count == 0) return new List<int>();
            return _items.Where(i => i.Marcado && !i.EsTodos).Select(i => i.Id).ToList();
        }

        public int Primario()
        {
            return _items.FirstOrDefault(i => i.Marcado && !i.EsTodos)?.Id ?? 0;
        }

        public T Primario<T>() where T : class
        {
            var item = _items.FirstOrDefault(i => i.Marcado && !i.EsTodos);
            return item?.Tag as T;
        }

        public void RefrescarTexto()
        {
            var itemsReales = _items.Where(i => !i.EsTodos).ToList();
            if (itemsReales.Count == 0)
            {
                // Fase 1460: Si no hay ítems reales pero "Todos" está marcado, mostrar TextoVacio en lugar de "—"
                bool todosVacio = _items.Any(i => i.EsTodos && i.Marcado);
                _txtDisplay.Text = (todosVacio && !string.IsNullOrWhiteSpace(TextoVacio)) ? TextoVacio : "—";
                return;
            }

            var marcadosReales = itemsReales.Where(i => i.Marcado).ToList();
            int cantMarcadosReales = marcadosReales.Count;
            bool todosMarcado = _items.Any(i => i.EsTodos && i.Marcado);

            if (cantMarcadosReales == 0 || todosMarcado)
            {
                if (!string.IsNullOrWhiteSpace(TextoVacio))
                {
                    _txtDisplay.Text = TextoVacio;
                }
                else
                {
                    _txtDisplay.Text = VacioEsTodos ? $"Todos ({itemsReales.Count})" : "Ninguno";
                }
            }
            else if (cantMarcadosReales == 1)
            {
                _txtDisplay.Text = marcadosReales[0].Texto;
            }
            else if (cantMarcadosReales == itemsReales.Count)
            {
                if (!string.IsNullOrWhiteSpace(TextoVacio))
                {
                    _txtDisplay.Text = TextoVacio;
                }
                else
                {
                    _txtDisplay.Text = $"Todos ({itemsReales.Count})";
                }
            }
            else
            {
                _txtDisplay.Text = $"{cantMarcadosReales} de {itemsReales.Count}";
            }
        }

        // Fase 1436: Liberación explícita de ToolStripDropDown y ToolStripControlHost para prevenir fugas de recursos GDI/USER
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_dropDown != null)
                {
                    _dropDown.Dispose();
                    _dropDown = null;
                }
                if (_host != null)
                {
                    _host.Dispose();
                    _host = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
