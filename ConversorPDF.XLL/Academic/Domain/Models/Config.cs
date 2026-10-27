using System;
using System.ComponentModel;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa la configuración general de la aplicación.
    /// Entidad de un único registro permanente.
    /// </summary>
    public class Config
    {
        [Browsable(false)]
        public int Id { get; set; }
        
        [Category("General")]
        [DisplayName("Subunidad")]
        public string Subunidad { get; set; }
        
        [Category("General")]
        [DisplayName("ID Territorial")]
        public int IdTerritorial { get; set; }
        
        [Category("General")]
        [DisplayName("Periodo")]
        public int IdPeriodo { get; set; }
        
        [Category("General")]
        [DisplayName("Fecha Inicio")]
        public string FechaInicio { get; set; }
        
        [Category("General")]
        [DisplayName("Fecha Fin")]
        public string FechaFin { get; set; }
        
        [Category("Financiero")]
        [DisplayName("CDP")]
        public string Cdp { get; set; }
        
        [Category("Financiero")]
        [DisplayName("Salario Decreto 613 (2025)")]
        public double SalarioDecreto6132025 { get; set; }
        
        [Category("Financiero")]
        [DisplayName("Viático Diario Tarifa A")]
        public double ViaticoDiarioTarifaA { get; set; }
        
        [Category("Financiero")]
        [DisplayName("Viático Diario Tarifa B")]
        public double ViaticoDiarioTarifaB { get; set; }
        
        [Category("Financiero")]
        [DisplayName("Porcentaje Decreto A")]
        public double PorcentajeDecretoA { get; set; }
        
        [Category("Financiero")]
        [DisplayName("Porcentaje Decreto B")]
        public double PorcentajeDecretoB { get; set; }
        
        [Category("Directivos")]
        [DisplayName("Director Territorial Nombre")]
        public string DtNombre { get; set; }
        
        [Category("Directivos")]
        [DisplayName("Director Territorial Cargo")]
        public string DtCargo { get; set; }
        
        [Category("Directivos")]
        [DisplayName("Corporativo Cédula")]
        public string CorporativoCedula { get; set; }
        
        [Category("Directivos")]
        [DisplayName("Corporativo Nombre")]
        public string CorporativoNombre { get; set; }
        
        [Category("Directivos")]
        [DisplayName("Corporativo Cargo")]
        public string CorporativoCargo { get; set; }
        
        [Category("Base")]
        [DisplayName("Departamento Base")]
        public string DepartamentoBase { get; set; }
        
        [Category("Base")]
        [DisplayName("Ciudad Base")]
        public string CiudadBase { get; set; }
        
        [Category("Base")]
        [DisplayName("Aeropuerto Base")]
        public string AeropuertoBase { get; set; }
        
        [Browsable(false)]
        public string CreatedAt { get; set; }
        
        [Browsable(false)]
        public string UpdatedAt { get; set; }
    }
}
