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
            
            conn.Execute("DROP TABLE IF EXISTS test_data;");
            conn.Execute("CREATE TABLE test_data ( id INTEGER PRIMARY KEY NOT NULL, val TEXT );");
            conn.Execute("INSERT INTO test_data (id, val) VALUES (1, 'hello');");
            conn.Execute("SELECT crsql_as_crr('test_data');");
            
            try {
                using (var cmd = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'test_data' LIMIT 1;", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    Console.WriteLine("OK: test_data");
                }
            } catch (Exception ex) {
                Console.WriteLine("FAIL: test_data -> " + ex.Message);
            }
        }
    }
}
public static class SqliteExt {
    public static void Execute(this SQLiteConnection conn, string sql) {
        using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
    }
}
