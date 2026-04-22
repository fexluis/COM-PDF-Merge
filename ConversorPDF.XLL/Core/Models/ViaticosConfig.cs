using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Modelo de configuración global para el procesamiento de viáticos, mapeado desde la hoja 'Config'.
    /// </summary>
    public class ViaticosConfig
    {
        /// <summary> Fecha de inicio del periodo configurado. </summary>
        public DateTime FechaInicio { get; set; }

        /// <summary> Fecha de fin del periodo configurado. </summary>
        public DateTime FechaFin { get; set; }

        /// <summary> Número de CDP global para la vigencia actual. </summary>
        public string Cdp { get; set; }

        /// <summary> Valor del salario decretado para el cálculo de topes. </summary>
        public object SalarioDecreto { get; set; }

        /// <summary> Tarifa de viáticos para el nivel A. </summary>
        public object ViaticoA { get; set; }

        /// <summary> Tarifa de viáticos para el nivel B. </summary>
        public object ViaticoB { get; set; }

        /// <summary> Porcentaje de liquidación para el nivel A (ej: 0.8). </summary>
        public object PorcA { get; set; }

        /// <summary> Porcentaje de liquidación para el nivel B. </summary>
        public object PorcB { get; set; }

        /// <summary> Nombre del Director Técnico o autoridad competente. </summary>
        public string NombreDT { get; set; }

        /// <summary> Cargo de la autoridad competente. </summary>
        public string CargoDT { get; set; }

        /// <summary> Nombre del Gerente o responsable de la entidad. </summary>
        public string Gerente { get; set; }

        /// <summary> Cargo principal del responsable. </summary>
        public string Cargo { get; set; }

        /// <summary> Documento de identidad del gerente. </summary>
        public string CedulaGerente { get; set; }

        /// <summary> Nombre de quien sustituye en caso de ausencia. </summary>
        public string GerenteSustituto { get; set; }

        /// <summary> Cargo de la persona sustituta. </summary>
        public string CargoSustituto { get; set; }

        /// <summary> Departamento base para el cálculo de origen (ej: Cundinamarca). </summary>
        public string DeptoBase { get; set; }

        /// <summary> Ciudad base (ej: Bogotá). </summary>
        public string CiudadBase { get; set; }

        /// <summary> Sigla de aeropuerto principal (ej: BOG). </summary>
        public string AeropuertoBase { get; set; }

        /// <summary> Sigla o nombre corto de la dependencia. </summary>
        public string CortaDependencia { get; set; }

        /// <summary> Nombre completo de la dependencia administradora. </summary>
        public string FullDependencia { get; set; }

        /// <summary> Fuente de financiación del CDP. </summary>
        public string Fcdp { get; set; }

        /// <summary> Código presupuestal (Rubro). </summary>
        public string Rubro { get; set; }

        /// <summary> Nombre de quien elaboró el documento. </summary>
        public string Elaboro { get; set; }

        /// <summary> Nombre de quien revisó o auditó el trámite. </summary>
        public string Reviso { get; set; }

        /// <summary> Ruta física donde se guardarán los PDFs generados. </summary>
        public string CarpetaDestino { get; set; }
    }
}