using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.IO;
namespace ConversorPDF.XLL.Core.Helpers
{
    public static class ViaticosHelper
    {
        public static DateTime ToDateTime(object val)
        {
            if (val == null) return DateTime.MinValue;
            if (val is DateTime dt) return dt;
            if (val is double d) return DateTime.FromOADate(d);
            if (double.TryParse(val?.ToString(), out double d2)) return DateTime.FromOADate(d2);
            if (DateTime.TryParse(val?.ToString(), out DateTime dt2)) return dt2;
            return DateTime.MinValue;
        }

        public static string SinTildes(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return texto;
            const string conTilde = "ÁÉÍÓÚáéíóú";
            const string sinTilde = "AEIOUaeiou";
            for (int i = 0; i < conTilde.Length; i++)
                texto = texto.Replace(conTilde[i], sinTilde[i]);
            return texto;
        }

        public static string LimpiarTexto(string t) => t?.Replace("\"", "").Replace(":", "").Replace("/", "-");

        public static string CrearCarpetaDestino()
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Viaticos-" + DateTime.Now.ToString("yyyyMMdd"));
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }
    }
}