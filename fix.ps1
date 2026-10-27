$filePath = "ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs"
$lines = Get-Content $filePath -Encoding UTF8
$newLines = $lines[0..614]
$newLines | Set-Content $filePath -Encoding UTF8

$append = @"
                MessageBox.Show(`"Exportación a la hoja '{nombreHoja}' completada con éxito ({terceros.Count} registros).`", `"Éxito`", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(`"Error al exportar a Excel: `{ex.Message}`", `"Error de Exportación`", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        }

        private void EscribirCelda(dynamic ws, int fila, int col, object valor)
        {
            dynamic cell = null;
            try
            {
                cell = ws.Cells[fila, col];
                cell.Value = valor;
            }
            finally
            {
                if (cell != null) Marshal.ReleaseComObject(cell);
            }
        }

        private void chkListDepartamentos_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string currentSubunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
            string item = chkListDepartamentos.Items[e.Index].ToString();

            if (!string.Equals(item, currentSubunidad, StringComparison.OrdinalIgnoreCase))
            {
                e.NewValue = e.CurrentValue;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            // Fase 15 FIX: Refrescar FormVinculaciones inmediatamente si está abierto
            foreach (Form form in Application.OpenForms)
            {
                if (form is FormVinculaciones fv)
                {
                    fv.RecargarCatalogosUI();
                }
            }
        }
    }
}
"@

Add-Content -Path $filePath -Value $append -Encoding UTF8
