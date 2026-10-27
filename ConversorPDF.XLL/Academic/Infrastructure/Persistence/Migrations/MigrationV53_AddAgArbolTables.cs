using System;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V53: Tablas para perfiles de selección de árbol académico (ag_arbol y ag_arbol_nodo).
    /// Permite persistir y restaurar configuraciones del árbol de 6 niveles, ventana temporal y parámetros evolutivos (Fase 236).
    /// </summary>
    public class MigrationV53_AddAgArbolTables : IMigration, IMigrationTolerante
    {
        public int Version => 53;
        public string Name => "AddAgArbolTables";
        public bool Aditiva => false; // Fase 259: No aditiva, fallo de esquema aborta el arranque general

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V53] Creando tablas ag_arbol y ag_arbol_nodo...");

            const string sqlArbol = @"
CREATE TABLE IF NOT EXISTS ag_arbol (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL DEFAULT '',
    descripcion TEXT NOT NULL DEFAULT '',
    idperiodo INTEGER NOT NULL DEFAULT 0,
    fecha_inicio TEXT NOT NULL DEFAULT '',
    fecha_fin TEXT NOT NULL DEFAULT '',
    ventana_origen TEXT NOT NULL DEFAULT '',
    manual_ventana INTEGER NOT NULL DEFAULT 0,
    poblacion INTEGER NOT NULL DEFAULT 50,
    generaciones INTEGER NOT NULL DEFAULT 200,
    tasa_cruce REAL NOT NULL DEFAULT 0.85,
    tasa_mutacion REAL NOT NULL DEFAULT 0.10,
    elitismo INTEGER NOT NULL DEFAULT 2,
    estancamiento_max INTEGER NOT NULL DEFAULT 20,
    semilla INTEGER NOT NULL DEFAULT 0,
    usuario TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT '',
    creado_en TEXT NOT NULL DEFAULT '',
    modificado_en TEXT NOT NULL DEFAULT '',
    activo INTEGER NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_ag_arbol_nombre ON ag_arbol(nombre, subunidad);
";

            const string sqlArbolNodo = @"
CREATE TABLE IF NOT EXISTS ag_arbol_nodo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idarbol INTEGER NOT NULL,
    idperiodo INTEGER NOT NULL DEFAULT 0,
    idsede INTEGER NOT NULL DEFAULT 0,
    idprograma INTEGER NOT NULL DEFAULT 0,
    semestre TEXT NOT NULL DEFAULT '',
    grupo TEXT NOT NULL DEFAULT '',
    modalidad TEXT NOT NULL DEFAULT '',
    aula_ref TEXT NOT NULL DEFAULT '',
    capacidad INTEGER NOT NULL DEFAULT 0,
    tipo_aula TEXT NOT NULL DEFAULT '',
    caracteristicas_json TEXT NOT NULL DEFAULT '{}',
    codigo TEXT NOT NULL DEFAULT '',
    nombre TEXT NOT NULL DEFAULT '',
    es_virtual INTEGER NOT NULL DEFAULT 0,
    habilitado INTEGER NOT NULL DEFAULT 1,
    subunidad TEXT NOT NULL DEFAULT ''
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref);
CREATE INDEX IF NOT EXISTS idx_ag_arbol_nodo_arbol ON ag_arbol_nodo(idarbol);
";

            connection.Execute(sqlArbol, transaction: transaction);
            connection.Execute(sqlArbolNodo, transaction: transaction);

            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V53] Tablas ag_arbol y ag_arbol_nodo creadas exitosamente.");
        }
    }
}
