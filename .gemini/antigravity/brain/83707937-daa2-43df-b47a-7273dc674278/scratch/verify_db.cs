using System;
using System.Data;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models.Academic;
using System.Linq;

namespace Verification
{
    public class DbVerifier
    {
        public static void Run()
        {
            try
            {
                Console.WriteLine("Iniciando verificación de base de datos...");
                
                using (var db = SqliteHelper.GetConnection())
                {
                    Console.WriteLine("Conexión establecida.");

                    // Verificar tablas
                    var tables = db.Query<string>("SELECT name FROM sqlite_master WHERE type='table'").ToList();
                    string[] expectedTables = { "terceros", "sedes", "vinculaciones", "sesiones" };
                    
                    foreach (var table in expectedTables)
                    {
                        if (tables.Contains(table))
                        {
                            Console.WriteLine($"Tabla '{table}' encontrada.");
                        }
                        else
                        {
                            Console.WriteLine($"ERROR: Tabla '{table}' NO encontrada.");
                        }
                    }

                    // Prueba de inserción en Terceros
                    var testTercero = new Tercero 
                    { 
                        Nombre = "Test User", 
                        Cedula = "12345", 
                        Tipo = "Servidores",
                        Salario = 5000,
                        Rol = "Funcionario",
                        EntidadBancaria = "Banco Test",
                        CuentaBancaria = "Ahorro",
                        FechaDeNacimiento = DateTime.Now.AddYears(-30)
                    };
                    
                    string sql = @"INSERT INTO terceros (nombre, cedula, tipo, salario, rol, entidad_bancaria, cuenta_bancaria, fecha_nacimiento) 
                                   VALUES (@Nombre, @Cedula, @Tipo, @Salario, @Rol, @EntidadBancaria, @CuentaBancaria, @FechaDeNacimiento); 
                                   SELECT last_insert_rowid();";
                    
                    long id = db.QuerySingle<long>(sql, testTercero);
                    Console.WriteLine($"Inserción en terceros exitosa. ID: {id}");

                    // Limpieza
                    db.Execute("DELETE FROM terceros WHERE id = @id", new { id });
                    Console.WriteLine("Limpieza completada.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR DURANTE LA VERIFICACIÓN: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
