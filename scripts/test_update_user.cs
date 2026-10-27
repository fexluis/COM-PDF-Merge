using System;
using System.Data.SQLite;
using System.IO;
using Dapper;

class Program
{
    static void Main()
    {
        var tempDir = new DirectoryInfo(@"C:\Users\USUARIO\AppData\Local\Temp");
        var files = tempDir.GetFiles("esap_acad_tmp_AcademicData_BOLIVAR*.db");
        if (files.Length == 0) return;
        Array.Sort(files, (a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));
        var targetDb = files[0].FullName;
        
        using (var conn = new SQLiteConnection(string.Format("Data Source={0};Version=3;", targetDb)))
        {
            conn.Open();
            conn.EnableExtensions(true);
            conn.LoadExtension(@"C:\Users\USUARIO\AppData\Local\Temp\crsqlite_esap.dll", "sqlite3_crsqlite_init");

            try {
                // Obtenemos un usuario cualquiera
                var id = conn.QueryFirstOrDefault<string>("SELECT id FROM usuarios LIMIT 1");
                if (id == null) {
                    Console.WriteLine("No hay usuarios");
                    return;
                }
                
                Console.WriteLine("Updating user " + id);
                
                string sql = @"UPDATE usuarios SET 
                    nombre = 'Test', nombreusuario = 'Test', subunidad = 'Test', activo = 1, perfil = 'Admin', firma = 'TestFirma' 
                    WHERE id = @Id";
                
                int filas = conn.Execute(sql, new { Id = id });
                Console.WriteLine("OK, updated " + filas);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }
    }
}
