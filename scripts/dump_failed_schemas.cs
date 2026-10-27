using System;
using System.Data.SQLite;
using System.IO;
using System.Text;

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
            StringBuilder sb = new StringBuilder();
            
            string[] tables = new[] { "salvedades_conflictos", "territoriales", "divipola", "permisos" };
            foreach(var tbl in tables) {
                using (var cmd = new SQLiteCommand(string.Format("SELECT type, name, sql FROM sqlite_master WHERE name LIKE '%{0}%';", tbl), conn))
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
            }
            Console.WriteLine(sb.ToString());
        }
    }
}
