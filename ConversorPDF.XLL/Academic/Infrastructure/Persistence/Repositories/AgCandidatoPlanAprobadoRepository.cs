using System;
using System.Data;
using System.Globalization;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>Inserta efecto y ratificación en la transacción del confirmador; nunca actualiza producción.</summary>
    public sealed class AgCandidatoPlanAprobadoRepository : IAgCandidatoPlanAprobadoRepository
    {
        /// <inheritdoc />
        public bool Guardar(IDbConnection db, IDbTransaction tx, AgCandidatoPlanAprobado c)
        {
            if (db == null || tx == null || c == null) throw new ArgumentNullException(nameof(c));
            var datos = new {
                Operacion = c.Operacion.ToString("N"), Intento = c.Intento.ToString("N"), c.IdAula,
                c.Actor, c.Motivo, c.Alcance, c.HuellaManifiesto, c.HuellaOrigen, c.HuellaPrevia,
                c.HuellaEfecto, c.Desde, c.Hacia, c.Consentimiento,
                Inicio = c.Inicio.ToString("O", CultureInfo.InvariantCulture),
                Fin = c.Fin.ToString("O", CultureInfo.InvariantCulture),
                Creado = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
            };
            if (db.Execute(@"INSERT INTO ag_candidato_plan_aprobado
(operacion,intento,id_aula,actor,motivo,alcance,huella_manifiesto,huella_origen,huella_previa,huella_efecto,desde,hacia,consentimiento,inicio,fin,creado_utc)
VALUES (@Operacion,@Intento,@IdAula,@Actor,@Motivo,@Alcance,@HuellaManifiesto,@HuellaOrigen,@HuellaPrevia,@HuellaEfecto,@Desde,@Hacia,@Consentimiento,@Inicio,@Fin,@Creado)", datos, tx) != 1) return false;
            return db.Execute(@"INSERT INTO ag_auditoria_ratificacion
(operacion,actor,huella_origen,motivo,alcance,creado_utc)
VALUES (@Operacion,@Actor,@HuellaOrigen,@Motivo,@Alcance,@Creado)", datos, tx) == 1;
        }
    }
}
