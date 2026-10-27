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

            conn.Execute("DROP TABLE IF EXISTS config_copy;");
            conn.Execute("CREATE TABLE \"config_copy\" ( id INTEGER PRIMARY KEY NOT NULL, subunidad TEXT , id_territorial INTEGER , id_periodo INTEGER , fecha_inicio TEXT, fecha_fin TEXT, cdp TEXT, salario_decreto_613_2025 REAL, viatico_diario_tarifa_a REAL, viatico_diario_tarifa_b REAL, porcentaje_decreto_a REAL, porcentaje_decreto_b REAL, dt_nombre TEXT, dt_cargo TEXT, corporativo_cedula TEXT, corporativo_nombre TEXT, corporativo_cargo TEXT, departamento_base TEXT, ciudad_base TEXT, aeropuerto_base TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')) )");
            conn.Execute("SELECT crsql_as_crr('config_copy');");

            Compare(conn, "config");
            Compare(conn, "config_copy");
        }
    }

    static void Compare(SQLiteConnection conn, string name) {
        Console.WriteLine("\n--- " + name + " triggers ---");
        using(var cmd = new SQLiteCommand(string.Format("SELECT name, sql FROM sqlite_master WHERE type='trigger' AND tbl_name='{0}'", name), conn))
        using(var reader = cmd.ExecuteReader()) {
            while(reader.Read()) {
                Console.WriteLine(string.Format("{0}: {1}", reader["name"], reader["sql"]));
            }
        }
    }
}

public static class SqliteExt {
    public static void Execute(this SQLiteConnection conn, string sql) {
        using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
    }
}
