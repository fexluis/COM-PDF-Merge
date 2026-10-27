using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Tercero
    {
        public int Id { get; set; }
        public string Departamento { get; set; }
        public string Subunidad { get; set; }
        public string Nombre { get; set; }
        public string Cedula { get; set; }
        
        /// <summary>
        /// Tipos: "CARRERA", "OCASIONAL", "CATEDRA", "ESPECIAL/VISITANTE", "CONTRATISTA"
        /// </summary>
        public string Tipo { get; set; }
        
        public decimal Salario { get; set; }
        
        /// <summary>
        /// Roles: "Docente catedrático - Capacitadores", "Ocasionales", "Docente de Planta", "Contratista", "Funcionario"
        /// </summary>
        public string Categoria { get; set; }
        
        public int Pensionado { get; set; }
        
        public string EntidadBancaria { get; set; }
        
        /// <summary>
        /// Tipos de Cuenta: "Ahorro", "Corriente"
        /// </summary>
        public string CuentaBancaria { get; set; }
        public string TipoCuentaBancaria { get; set; }
        
        public string Telefono { get; set; }
        public string Correo { get; set; }
        public string CorreoPersonal { get; set; }
        public string FechaDeNacimiento { get; set; }
        public string Contrato { get; set; }
        public string Residencia { get; set; }
        
        /// <summary>
        /// Identificador de la residencia cruzado con la tabla divipola
        /// </summary>
        public int? IdResidenciaDivipola { get; set; }
        
        /// <summary>
        /// Mapea a la columna 'activo' (INTEGER 1 o 0 en SQLite)
        /// </summary>
        public int Activo { get; set; } = 1;

        public decimal PuntajeOcasionales { get; set; }
    }
}
