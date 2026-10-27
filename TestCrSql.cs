using System;
using System.Data.SQLite;

class Program
{
    static void Main()
    {
        string dbPath = ""test_crsql.db"";
        if (System.IO.File.Exists(dbPath)) System.IO.File.Delete(dbPath);

        using (var conn = new SQLiteConnection($""Data Source={dbPath}""))
        {
            conn.Open();
            conn.EnableExtensions(true);
            try
            {
                conn.LoadExtension(""crsqlite.dll"");
            }
            catch(Exception e)
            {
                Console.WriteLine(""Error load: "" + e.Message);
                return;
            }

            // Create table
            var cmd = conn.CreateCommand();
            cmd.CommandText = ""CREATE TABLE test (id INTEGER PRIMARY KEY, a TEXT);"";
            cmd.ExecuteNonQuery();

            // Make CRR
            cmd.CommandText = ""SELECT crsql_as_crr('test');"";
            cmd.ExecuteNonQuery();

            // Alter inside transaction
            using(var tx = conn.BeginTransaction())
            {
                cmd.CommandText = ""ALTER TABLE test ADD COLUMN b TEXT;"";
                cmd.ExecuteNonQuery();
                tx.Commit();
            }

            Console.WriteLine(""Altered successfully."");

            // Insert should fail because triggers out of sync
            try
            {
                cmd.CommandText = ""INSERT INTO test (a, b) VALUES ('val_a', 'val_b');"";
                cmd.ExecuteNonQuery();
                Console.WriteLine(""Insert successful (unexpected if triggers are out of sync)."");
            }
            catch(Exception e)
            {
                Console.WriteLine(""Insert failed as expected: "" + e.Message);
            }

            // Try to fix it by running crsql_as_crr again?
            try
            {
                cmd.CommandText = ""SELECT crsql_as_crr('test');"";
                cmd.ExecuteNonQuery();
                Console.WriteLine(""crsql_as_crr ran again."");
            }
            catch(Exception e)
            {
                Console.WriteLine(""crsql_as_crr failed: "" + e.Message);
            }

            // Try insert again
            try
            {
                cmd.CommandText = ""INSERT INTO test (a, b) VALUES ('val_a2', 'val_b2');"";
                cmd.ExecuteNonQuery();
                Console.WriteLine(""Insert 2 successful after crsql_as_crr."");
            }
            catch(Exception e)
            {
                Console.WriteLine(""Insert 2 failed: "" + e.Message);
            }

            // Fix by crsql_alter_begin ?
            try
            {
                cmd.CommandText = ""SELECT crsql_alter_begin('test'); SELECT crsql_alter_commit('test');"";
                cmd.ExecuteNonQuery();
                Console.WriteLine(""crsql_alter begin/commit successful."");
            }
            catch(Exception e)
            {
                Console.WriteLine(""crsql_alter failed: "" + e.Message);
            }
            
            // Try insert 3
            try
            {
                cmd.CommandText = ""INSERT INTO test (a, b) VALUES ('val_a3', 'val_b3');"";
                cmd.ExecuteNonQuery();
                Console.WriteLine(""Insert 3 successful after alter begin/commit."");
            }
            catch(Exception e)
            {
                Console.WriteLine(""Insert 3 failed: "" + e.Message);
            }
        }
    }
}
