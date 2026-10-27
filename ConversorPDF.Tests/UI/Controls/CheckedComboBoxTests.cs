using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.UI.Controls;

namespace ConversorPDF.Tests.UI.Controls
{
    [TestClass]
    public class CheckedComboBoxTests
    {
        private static void RunInSTA(Action action)
        {
            Exception threadEx = null;
            var thread = new Thread(() =>
            {
                try { action(); }
                catch (Exception ex) { threadEx = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (threadEx != null) throw threadEx;
        }

        [TestMethod]
        public void Cargar_ConIncluirItemTodos_InsertaPseudoItemEnIndiceCeroMarcado()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.TextoItemTodos = "-- Todos los Programas --";
                    cmb.TextoVacio = "Todos los Programas";

                    var items = new[]
                    {
                        (10, "Programa 10"),
                        (20, "Programa 20")
                    };

                    cmb.Cargar(items);

                    // Id 0 no debe ser retornado en SeleccionadosIds
                    var ids = cmb.SeleccionadosIds();
                    Assert.AreEqual(0, ids.Count);
                    Assert.IsFalse(ids.Contains(0));
                }
            });
        }

        [TestMethod]
        public void AplicarExclusionTodos_AlMarcarItemReal_DesmarcaPseudoItemTodos()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    var items = new[]
                    {
                        (10, "Prog A"),
                        (20, "Prog B")
                    };
                    cmb.Cargar(items);

                    // Inicialmente "Todos" (índice 0) está marcado
                    Assert.AreEqual(0, cmb.SeleccionadosIds().Count);

                    // Marcar ítem real en índice 1 (Id 10)
                    cmb.Marcar(new[] { 10 });

                    var ids = cmb.SeleccionadosIds();
                    Assert.AreEqual(1, ids.Count);
                    Assert.AreEqual(10, ids[0]);
                    Assert.IsFalse(ids.Contains(0));
                }
            });
        }

        [TestMethod]
        public void AplicarExclusionTodos_AlMarcarTodos_DesmarcaItemsReales()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    var items = new[]
                    {
                        (10, "Prog A"),
                        (20, "Prog B")
                    };
                    cmb.Cargar(items);

                    // Marcar ítems reales
                    cmb.Marcar(new[] { 10, 20 });
                    Assert.AreEqual(2, cmb.SeleccionadosIds().Count);

                    // Simular marcado de "Todos" (índice 0)
                    cmb.AplicarExclusionTodos(0, true);

                    var ids = cmb.SeleccionadosIds();
                    Assert.AreEqual(0, ids.Count);
                    Assert.IsFalse(ids.Contains(0));
                }
            });
        }

        [TestMethod]
        public void SeleccionadosIds_Invariante_NuncaContieneIdCero()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });

                    // Con selección por defecto
                    Assert.IsFalse(cmb.SeleccionadosIds().Contains(0));

                    // Con selección explícita incluyendo 0 por error
                    cmb.Marcar(new[] { 0, 10 });
                    var ids = cmb.SeleccionadosIds();
                    Assert.IsFalse(ids.Contains(0));
                    Assert.AreEqual(1, ids.Count);
                    Assert.AreEqual(10, ids[0]);

                    // Limpiando
                    cmb.Limpiar();
                    Assert.IsFalse(cmb.SeleccionadosIds().Contains(0));
                    Assert.AreEqual(0, cmb.SeleccionadosIds().Count);
                }
            });
        }

        [TestMethod]
        public void Marcar_ConListaVacia_MarcaTodosYDesmarcaReales()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.TextoVacio = "Todos los Programas";
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });

                    cmb.Marcar(new[] { 10 });
                    Assert.AreEqual(1, cmb.SeleccionadosIds().Count);

                    cmb.Marcar(Enumerable.Empty<int>());
                    Assert.AreEqual(0, cmb.SeleccionadosIds().Count);
                }
            });
        }

        [TestMethod]
        public void Marcar_ConIdsInexistentes_MarcaTodosYNoDejaEstadoVacioIncoherente()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.TextoVacio = "-- Todos los Programas --";
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });

                    // Se pasa un ID que no existe en el catálogo disponible
                    cmb.Marcar(new[] { 999 });

                    // Debe marcar "Todos" y SeleccionadosIds retornar vacío (semántica de todos)
                    Assert.AreEqual(0, cmb.SeleccionadosIds().Count);
                    Assert.IsFalse(cmb.SeleccionadosIds().Contains(999));
                    Assert.IsFalse(cmb.SeleccionadosIds().Contains(0));
                }
            });
        }

        [TestMethod]
        public void Marcar_EnModoUnicoConMultiplesIds_SoloMarcaElPrimerItemReal()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.Modo = CheckedComboModo.Unico;
                    cmb.IncluirItemTodos = true;
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });

                    // Pasar múltiples IDs válidos en modo único
                    cmb.Marcar(new[] { 10, 20 });

                    var ids = cmb.SeleccionadosIds();
                    Assert.AreEqual(1, ids.Count);
                    Assert.AreEqual(10, ids[0]);
                }
            });
        }

        [TestMethod]
        public void Cargar_ReseteaFirmaSeleccion_YPermiteDispararSeleccionCambiada()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });
                    cmb.Marcar(new[] { 10 });

                    // Simular notificación de cambio para registrar firma "10"
                    int eventosDisparados = 0;
                    cmb.SeleccionCambiada += (s, e) => eventosDisparados++;

                    // Recargar ítems
                    cmb.Cargar(new[] { (10, "Prog A"), (20, "Prog B") });
                    // Volver a marcar "10": debe disparar el evento si la firma se reseteó a null
                    cmb.Marcar(new[] { 10 });

                    Assert.IsTrue(eventosDisparados > 0);
                }
            });
        }

        [TestMethod]
        public void RefrescarTexto_SinItemsRealesYTodosMarcado_MuestraTextoVacio()
        {
            RunInSTA(() =>
            {
                using (var cmb = new CheckedComboBox())
                {
                    cmb.IncluirItemTodos = true;
                    cmb.TextoVacio = "Todos los Programas";
                    // Cargar lista vacía de ítems reales (solo tendrá pseudo-ítem Todos)
                    cmb.Cargar(Enumerable.Empty<(int, string)>());

                    Assert.AreEqual("Todos los Programas", cmb.TextoMostrado);
                }
            });
        }
    }
}
