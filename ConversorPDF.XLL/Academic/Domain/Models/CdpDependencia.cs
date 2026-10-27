namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class CdpDependencia
    {
        public int Id { get; set; }
        public int IdPeriodo { get; set; }
        public string CdpNum { get; set; }
        public string FechaCdp { get; set; }
        public string Rubro { get; set; }
        public string Responsable { get; set; }
        public string Cargo { get; set; }
        public string Elaboro { get; set; }
        public string Reviso { get; set; }
        public string Dependencia { get; set; }
        public string DependenciaFull { get; set; }
        public int Activo { get; set; }
        public string Subunidad { get; set; }
    }
}
