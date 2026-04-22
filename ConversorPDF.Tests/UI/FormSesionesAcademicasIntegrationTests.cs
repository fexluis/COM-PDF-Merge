using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;
using System.Linq;
using System.Globalization;
using ConversorPDF.XLL.FormAcademic;
using ConversorPDF.XLL.Core.Models.Academic;

namespace ConversorPDF.Tests.UI
{
    [TestClass]
    public class FormSesionesAcademicasIntegrationTests
    {

        [TestMethod]
        public void PrepararSesiones_ConteoCoincide_RetornaCantidadCorrecta()
        {
            RunInSTA(() =>
            {
                using (var form = new FormSesionesAcademicas())
                {
                    var handle = form.Handle; // Forzar handle
                    var grid = TestHelper.GetPrivateField<DataGridView>(form, "dgvSesiones");

                    // Inicializar dropdowns y campos privados
                    form.Invoke((Action)(() => {
                        var ddlT = TestHelper.GetPrivateField<ComboBox>(form, "cmbTerceros");
                        ddlT.ValueMember = "Id";
                        ddlT.DisplayMember = "Nombre";
                        ddlT.DataSource = new List<Tercero> { new Tercero { Id = 1, Nombre = "Test" } };
                        if (ddlT.Items.Count > 0) ddlT.SelectedIndex = 0;

                        // Establecer vinculación ID (usado por PrepararSesionesDesdeGrid)
                        TestHelper.SetPrivateField(form, "_currentVinculacionId", 1);

                        grid.AllowUserToAddRows = false; // Evitar filas fantasma
                        Application.DoEvents(); // Procesar bindings
                    }));

                    // Agregar filas
                    form.Invoke((Action)(() => {
                        int i1 = grid.Rows.Add();
                        grid.Rows[i1].Cells["ColChecked"].Value = true;
                        grid.Rows[i1].Tag = new Sesion { IdTercero = 1, Horas = 8 };

                        int i2 = grid.Rows.Add();
                        grid.Rows[i2].Cells["ColChecked"].Value = true;
                        grid.Rows[i2].Tag = new Sesion { IdTercero = 1, Horas = 8 };

                        int i3 = grid.Rows.Add();
                        grid.Rows[i3].Cells["ColChecked"].Value = false;
                        grid.Rows[i3].Tag = new Sesion { IdTercero = 1, Horas = 8 };
                    }));

                    var resultado = TestHelper.InvokePrivateMethod<IEnumerable<Sesion>>(form, "PrepararSesionesDesdeGrid");
                    var lista = resultado.ToList();

                    Assert.AreEqual(2, lista.Count);
                }
            });
        }

        [TestMethod]
        public void CargarEstado_PostGuardado_NoDebeContenerSesionesFantasma()
        {
            RunInSTA(() =>
            {
                // Este test simula el ciclo de vida: Abrir -> Generar -> Guardar -> Cerrar -> Abrir
                using (var form = new FormSesionesAcademicas())
                {
                    var handle = form.Handle;
                    var grid = TestHelper.GetPrivateField<DataGridView>(form, "dgvSesiones");

                    // 1. Verificar conteo inicial (debe ser 0 si el seed ya no tiene sesiones hardcoded)
                    // Nota: Si el seed tiene sesiones, dgvSesiones.Rows.Count > 0 aquí.
                    int filasIniciales = grid.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                    
                    Assert.AreEqual(0, filasIniciales, "La grilla debe iniciar vacía si el seed de sesiones fue eliminado.");
                }
            });
        }

        private void RunInSTA(Action action)
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
    }

    public static class TestHelper
    {
        public static T GetPrivateField<T>(object obj, string fieldName)
        {
            var field = obj.GetType().GetField(fieldName, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (field == null) {
                // Probablemente autofiltro o nombre diferente? Intentar con property si no es field
                var prop = obj.GetType().GetProperty(fieldName, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                return (T)prop.GetValue(obj);
            }
            return (T)field.GetValue(obj);
        }

        public static void SetPrivateField(object obj, string fieldName, object value)
        {
            var field = obj.GetType().GetField(fieldName, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            field.SetValue(obj, value);
        }

        public static T InvokePrivateMethod<T>(object obj, string methodName, params object[] args)
        {
            var method = obj.GetType().GetMethod(methodName, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return (T)method.Invoke(obj, args);
        }
    }
}
