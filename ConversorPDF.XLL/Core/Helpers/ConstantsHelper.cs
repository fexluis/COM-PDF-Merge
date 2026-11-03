using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Core.Helpers
{
    public static class ConstantsHelper
    {
        public const string FORMATO_PASSWORD = "qapst01";
        public const string DB_PASSWORD = "Sqlite@ESAP.2026";

        // Columnas Hoja Programación
        public const int PROG_TIPOFORMATO = 1;
        public const int PROG_CEDULA = 2;
        public const int PROG_NOMBRE = 3;
        public const int PROG_ITINERARIO = 6;
        public const int PROG_FECHAVIAJE = 8;
        public const int PROG_FECHAFIN = 9;
        public const int PROG_DURACION = 10;
        public const int PROG_TERMINAL = 16;
        public const int PROG_TRANSPORTE = 17;
        public const int PROG_TRANSPORTEAEREO = 18;
        public const int PROG_OBJETO = 20;
        public const int PROG_DEPENDENCIA = 21;
        public const int PROG_COL_MAX = 21;

        // Columnas Hoja CONSULTAS
        public const int CONSULTAS_FECHA_COMISION = 1;
        public const int CONSULTAS_NUMERO_COMISION = 2;
        public const int CONSULTAS_ESTADO = 3;
        public const int CONSULTAS_VALOR_TOTAL = 4;
        public const int CONSULTAS_INFORME = 8;
        public const int CONSULTAS_CEDULA = 10;
        public const int CONSULTAS_NOMBRE_COMPLETO = 11;
        public const int CONSULTAS_ITINERARIO = 12;
        public const int CONSULTAS_FECHA_INICIO = 13;
        public const int CONSULTAS_FECHA_FIN = 14;
        public const int CONSULTAS_CDP = 18;
        public const int CONSULTAS_RP = 19;
        public const int CONSULTAS_OB = 21;
        public const int CONSULTAS_OP = 22;
        public const int CONSULTAS_ROL = 27;
        public const int CONSULTAS_OBJETO_RAW = 32;
        public const int CONSULTAS_ES_MASIVO = 33;
        public const int CONSULTAS_COL_MAX = 33;

        // Columnas Hoja DATOS
        public const int DATOS_NOMBRE = 1;
        public const int DATOS_CEDULA = 2;
        public const int DATOS_SALARIO = 4;
        public const int DATOS_ROL = 5;
        public const int DATOS_CATEGORIA = 6;
        public const int DATOS_ENTIDAD = 7;
        public const int DATOS_CUENTA_NUM = 8;
        public const int DATOS_CUENTA_TIPO = 9;
        public const int DATOS_TEL = 10;
        public const int DATOS_CORREO = 11;
        public const int DATOS_NACIMIENTO = 12;
        public const int DATOS_CONTRATO = 13;
        public const int DATOS_DEPENDENCIA = 14;

        // Columnas Tabla tblDependencias
        public const int DEP_CORTA = 1;
        public const int DEP_FULL = 2;
        public const int DEP_CDP = 3;
        public const int DEP_FCDP = 4;
        public const int DEP_RUBRO = 5;
        public const int DEP_ELABORO = 8;
        public const int DEP_REVISO = 9;

        // Columnas Tabla tblDestinos
        public const int DEST_DEP = 1;
        public const int DEST_CIUDAD = 2;
        public const int DEST_NORMALIZADA = 3;

        // Hoja Config — rango A1:B16, columna de valores = 2 (columna B)
        public const int CONFIG_COL_VALOR        = 2;
        public const int CONFIG_FECHA_INICIO     = 1;
        public const int CONFIG_FECHA_FIN        = 2;
        public const int CONFIG_CDP              = 3;
        public const int CONFIG_SALARIO_DECRETO  = 4;
        public const int CONFIG_VIATICO_A        = 5;
        public const int CONFIG_VIATICO_B        = 6;
        public const int CONFIG_PORC_A           = 7;
        public const int CONFIG_PORC_B           = 8;
        public const int CONFIG_NOMBRE_DT        = 9;
        public const int CONFIG_CARGO_DT         = 10;
        public const int CONFIG_CEDULA_GERENTE   = 11;
        public const int CONFIG_GERENTE_SUSTITUTO = 12;
        public const int CONFIG_CARGO_SUSTITUTO  = 13;
        public const int CONFIG_DEPTO_BASE       = 14;
        public const int CONFIG_CIUDAD_BASE      = 15;
        public const int CONFIG_AEROPUERTO_BASE  = 16;
        public const int CONFIG_ROW_MAX          = 16;

        public const string SHEET_PROG = "Programación";
        public const string SHEET_CONSULTAS = "CONSULTAS";
        public const string SHEET_DATOS = "Datos";
        public const string SHEET_CONFIG = "Config";
        public const string SHEET_DEPENDENCIAS = "Dependencias";

        /// <summary>
        /// Carpetas (dentro de \Informes) donde se validan y unifican parejas INF + CUM.
        /// </summary>
        public static readonly string[] INFORMES_CARPETAS_ESPECIALES =
        {
            "Academica",
            "Capacitacion",
            "Direccion",
            "Fortalecimiento"
        };

        public const string THEME_PATH = @"C:\Program Files\Microsoft Office\root\Document Themes 16\Theme Colors\Office 2013 - 2022.xml";

        public static string GetTemplate(TipoFormato type)
        {
            switch (type)
            {
                case TipoFormato.Contratista: return "Contratistas";
                case TipoFormato.Servidor: return "Servidores";
                case TipoFormato.DocenteCatedra: return "Docentes Catedra";
                default: return "Contratistas";
            }
        }
    }
}