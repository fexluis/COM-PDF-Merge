using System;

namespace ConversorPDF.XLL.Core.Models.Academic
{
    public class Tercero
    {
        public int Id { get; set; }
        public string Departamento { get; set; }
        public string Nombre { get; set; }
        public string Cedula { get; set; }
        
        /// <summary>
        /// Tipos: "Docentes catedra", "Contratistas", "Servidores"
        /// </summary>
        public string Tipo { get; set; }
        
        public decimal Salario { get; set; }
        
        /// <summary>
        /// Roles: "Docente catedrático - Capacitadores", "Ocasionales", "Docente de Planta", "Contratista", "Funcionario"
        /// </summary>
        public string Rol { get; set; }
        
        public string EntidadBancaria { get; set; }
        
        /// <summary>
        /// Tipos de Cuenta: "Ahorro", "Corriente"
        /// </summary>
        public string CuentaBancaria { get; set; }
        
        public string Telefono { get; set; }
        public string Correo { get; set; }
        public DateTime FechaDeNacimiento { get; set; }
        public string Contrato { get; set; }
    }
}
