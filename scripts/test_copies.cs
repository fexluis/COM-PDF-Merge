using System;
using System.Data.SQLite;
using System.IO;

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
            
            TestTable(conn, "config_copy", "CREATE TABLE \"config_copy\" ( id INTEGER PRIMARY KEY NOT NULL, subunidad TEXT , id_territorial INTEGER , id_periodo INTEGER , fecha_inicio TEXT, fecha_fin TEXT, cdp TEXT, salario_decreto_613_2025 REAL, viatico_diario_tarifa_a REAL, viatico_diario_tarifa_b REAL, porcentaje_decreto_a REAL, porcentaje_decreto_b REAL, dt_nombre TEXT, dt_cargo TEXT, corporativo_cedula TEXT, corporativo_nombre TEXT, corporativo_cargo TEXT, departamento_base TEXT, ciudad_base TEXT, aeropuerto_base TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')) )");
            
            TestTable(conn, "salvedades_copy", "CREATE TABLE \"salvedades_copy\" ( id INTEGER PRIMARY KEY NOT NULL, idsesion_a INTEGER , idsesion_b INTEGER , idtercero INTEGER , idperiodo INTEGER, justificacion TEXT, fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP, subunidad TEXT DEFAULT '' )");
            
        }
    }
    
    static void TestTable(SQLiteConnection conn, string name, string sql) {
        conn.Execute("DROP TABLE IF EXISTS " + name);
        conn.Execute(sql);
        try {
            conn.Execute(string.Format("SELECT crsql_as_crr('{0}');", name));
            using (var cmd = new SQLiteCommand(string.Format("SELECT * FROM crsql_changes WHERE \"table\" = '{0}' LIMIT 1;", name), conn))
            using (var reader = cmd.ExecuteReader())
            {
                Console.WriteLine("OK: " + name);
            }
        } catch (Exception ex) {
            Console.WriteLine("FAIL: " + name + " -> " + ex.Message);
        }
    }
}
public static class SqliteExt {
    public static void Execute(this SQLiteConnection conn, string sql) {
        using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
    }
}
