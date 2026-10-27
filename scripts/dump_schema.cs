using System;
using System.Data.SQLite;
using System.IO;
using System.Text;

class Program
{
    static void Main()
    {
        string dbPath = @"C:\Users\USUARIO\AppData\Local\ConversorPDF\Data\AcademicData.db.enc";
        string backupPath = @"C:\Users\USUARIO\AppData\Local\Temp\dump_db_test.db";
        
        var tempDir = new DirectoryInfo(@"C:\Users\USUARIO\AppData\Local\Temp");
        var files = tempDir.GetFiles("esap_acad_tmp_AcademicData_BOLIVAR*.db");
        
        if (files.Length == 0) {
            Console.WriteLine("No decrypted db found in Temp.");
            return;
        }
        
        // Let's sort by LastWriteTime descending to get the most recent one
        Array.Sort(files, (a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));
        var targetDb = files[0].FullName;
        Console.WriteLine("Using DB: " + targetDb);
        
        using (var conn = new SQLiteConnection(string.Format("Data Source={0};Version=3;", targetDb)))
        {
            conn.Open();
            
            try {
                conn.EnableExtensions(true);
                conn.LoadExtension(@"C:\Users\USUARIO\AppData\Local\Temp\crsqlite_esap.dll", "sqlite3_crsqlite_init");
                Console.WriteLine("crsqlite loaded.");
            } catch(Exception ex) {
                Console.WriteLine("Could not load crsqlite: " + ex.Message);
            }
            
            StringBuilder sb = new StringBuilder();
            using (var cmd = new SQLiteCommand("SELECT type, name, sql FROM sqlite_master WHERE name LIKE '%usuarios%';", conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    sb.AppendLine(string.Format("--- {0} : {1} ---", reader["type"], reader["name"]));
                    if (reader["sql"] != DBNull.Value) {
                        sb.AppendLine(reader["sql"].ToString());
                    }
                    sb.AppendLine();
                }
            }
            
            Console.WriteLine(sb.ToString());
            
            try {
                using (var cmd2 = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'usuarios' LIMIT 1;", conn))
                using (var reader2 = cmd2.ExecuteReader())
                {
                    Console.WriteLine("Query to crsql_changes successful");
                }
            } catch (Exception ex) {
                Console.WriteLine("Error querying crsql_changes: " + ex.Message);
            }
        }
    }
}
