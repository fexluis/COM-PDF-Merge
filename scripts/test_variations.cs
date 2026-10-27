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
            
            TestTable(conn, "test11", "CREATE TABLE test11 ( id INTEGER PRIMARY KEY, val TEXT );");
            TestTable(conn, "test12", "CREATE TABLE test12 ( id INTEGER PRIMARY KEY NOT NULL, val TEXT );");
            TestTable(conn, "test13", "CREATE TABLE test13 ( perfil TEXT PRIMARY KEY, val TEXT );");
            TestTable(conn, "test14", "CREATE TABLE test14 ( perfil TEXT PRIMARY KEY NOT NULL, val TEXT );");
            TestTable(conn, "test15", "CREATE TABLE test15 ( id INTEGER PRIMARY KEY AUTOINCREMENT, val TEXT );");
            TestTable(conn, "test16", "CREATE TABLE test16 ( id TEXT PRIMARY KEY NOT NULL DEFAULT 'a', val TEXT );");
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
