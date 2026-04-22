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
    /// Representa la información de un empleado o contratista extraída de la base de datos de Excel.
    /// </summary>
    public class Empleado
    {
        /// <summary> Documento de identidad del empleado. </summary>
        public string Cedula { get; set; }

        /// <summary> Nombre completo del trabajador. </summary>
        public string NombreCompleto { get; set; }

        /// <summary> Salario base para el cálculo de viáticos (puede ser decimal o double). </summary>
        public object Salario { get; set; }

        /// <summary> Cargo o rol desempeñado (ej: Profesional Universitario). </summary>
        public string Rol { get; set; }

        /// <summary> Nivel o categoría salarial. </summary>
        public string Categoria { get; set; }

        /// <summary> Banco donde el empleado recibe sus pagos. </summary>
        public string EntidadBancaria { get; set; }

        /// <summary> Número de cuenta bancaria. </summary>
        public string NumeroCuenta { get; set; }

        /// <summary> Tipo de cuenta (Ahorros/Corriente). </summary>
        public string TipoCuenta { get; set; }

        /// <summary> Número telefónico de contacto. </summary>
        public string Telefono { get; set; }

        /// <summary> Dirección de correo electrónico institucional. </summary>
        public string Correo { get; set; }

        /// <summary> Fecha de nacimiento del empleado. </summary>
        public object FechaNacimiento { get; set; }

        /// <summary> Identificador del contrato vigente. </summary>
        public string Contrato { get; set; }

        /// <summary> Dependencia o unidad administrativa a la que pertenece. </summary>
        public string Dependencia { get; set; }
    }
}