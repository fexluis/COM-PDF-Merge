using System;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace ConversorPDF
{
    public static class FirmasGeneradorHelper
    {
        public static string GenerarStringBase64DesdeImagen()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Selecciona la imagen de la firma";
                openFileDialog.Filter = "Archivos PNG (*.png)|*.png|Todos los archivos (*.*)|*.*";
                
                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    if (!File.Exists(filePath)) return "";

                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string safeName = LimpiarNombrePropiedad(fileName);
                    
                    byte[] imageBytes = File.ReadAllBytes(filePath);
                    string base64String = Convert.ToBase64String(imageBytes);

                    // Formatear el resultado para que sea fácil de copiar a FirmasHelper.cs
                    string resultado = 
$@"// ============================================================================
// PROPIEDAD C#: {safeName}
// Descripción: Retorna imagen {fileName}.png como String Base64
// Archivo origen: {fileName}.png
// Fecha generación: {DateTime.Now:dd/MM/yyyy HH:mm:ss}
// Longitud Base64: {base64String.Length} caracteres
// ============================================================================
public static readonly string Firma{safeName} = 
    ""{base64String}"";
";
                    // Guardar en un TXT en la misma ruta que la imagen
                    string directory = Path.GetDirectoryName(filePath);
                    string outputTxt = Path.Combine(directory, $"Firma{safeName}.txt");
                    
                    File.WriteAllText(outputTxt, resultado);
                    
                    // Abrir el TXT para el usuario
                    System.Diagnostics.Process.Start("notepad.exe", outputTxt);
                    
                    return outputTxt;
                }
                return "";
            }
            catch (Exception ex)
            {
                return "ERROR: " + ex.Message;
            }
        }

        private static string LimpiarNombrePropiedad(string nombre)
        {
            if (string.IsNullOrWhiteSpace(nombre)) return "Firma";

            // Reemplazar caracteres no alfanuméricos por string vacío
            string clean = Regex.Replace(nombre, "[^a-zA-Z0-9]", "");
            
            // Si el primer caracter es un número, añadir un prefijo
            if (clean.Length > 0 && char.IsDigit(clean[0]))
            {
                clean = "_" + clean;
            }

            // Asegurarnos que la primera letra es mayúscula (PascalCase)
            if (clean.Length > 0)
            {
                clean = char.ToUpper(clean[0]) + clean.Substring(1);
            }

            return string.IsNullOrEmpty(clean) ? "Firma" : clean;
        }
    }
}
