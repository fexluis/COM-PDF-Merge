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
            
            TestTable(conn, "test_ts1", "CREATE TABLE test_ts1 ( id INTEGER PRIMARY KEY NOT NULL, fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP );");
            TestTable(conn, "test_ts2", "CREATE TABLE test_ts2 ( id INTEGER PRIMARY KEY NOT NULL, fecha_creacion TEXT DEFAULT (datetime('now')) );");
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
