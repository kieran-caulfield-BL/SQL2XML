using System;

namespace SQL2XML
{
    class Program
    {
        static void Main(string[] args)
        {
            string connString = @"dsn=2600;UID=sysprogress;PWD=sysprogress";

            string sql = @"Select
    PUB.MATDB.""MT-CODE"",
    PUB.MATDB.""CL-CODE"",
    PUB.MATDB.DESCRIPTION As ""MAT-DESCRIPTION"",
    PUB.MATDB.IMPLEMENTATION,
    PUB.MATDB.""DATE-OPENED"",
    PUB.MATDB.""MT-TYPE"",
    PUB.HISTORY.DESCRIPTION As ""HST-DESCRIPTION"",
    PUB.HISTORY.""LEVEL-FEE-EARNER"",
    PUB.HISTORY.""DOCUMENT-NAME"",
    PUB.HISTORY.""DOCUMENT-TYPE"",
    PUB.HISTORY.""HISTORY-NO"",
    PUB.HISTORY.""DATE-INSERTED"",
    PUB.""DOC-CONTROL"".""DOC-GROUP"",
    PUB.""DOC-CONTROL"".""ST-LOCATION"",
    PUB.""DOC-CONTROL"".""SUB-PATH"",
    PUB.""DOC-CONTROL"".""TYPE"" As ""DOC-TYPE"",
    PUB.""DOC-CONTROL"".EXTENSION,
    PUB.""DOC-CONTROL"".VERSION,
    PUB.""FILE-LOCATION"".DESCRIPTION As ""LOC-DESCRIPTION"",
    PUB.""FILE-LOCATION"".""LOC-OPTIONS"",
    PUB.""FILE-LOCATION"".""TYPE"" As ""LOC-TYPE""
From
    PUB.MATDB Inner Join
    PUB.HISTORY On PUB.HISTORY.""MT-CODE"" = PUB.MATDB.""MT-CODE""
            And PUB.HISTORY.IMPLEMENTATION = PUB.MATDB.IMPLEMENTATION
            And PUB.HISTORY.""DATE-INSERTED"" >= PUB.MATDB.""DATE-OPENED"" Inner Join
    PUB.""DOC-CONTROL"" On PUB.""DOC-CONTROL"".""DOC-ID"" = PUB.HISTORY.""DOCUMENT-NAME"" Inner Join
    PUB.""FILE-LOCATION"" On PUB.""FILE-LOCATION"".""LOC-NAME"" = PUB.""DOC-CONTROL"".""ST-LOCATION""
Where
    PUB.MATDB.""MT-CODE"" = '081389-000002' And
    PUB.HISTORY.""DOCUMENT-TYPE"" Is Not Null And
    PUB.HISTORY.""DOCUMENT-TYPE"" <> ''
";
            System.Data.Odbc.OdbcConnection conn = null;
            System.Data.Odbc.OdbcDataReader reader = null;

            try
            {
                // open connection
                conn = new System.Data.Odbc.OdbcConnection(connString);
                conn.Open();
                // execute the SQL
                System.Data.Odbc.OdbcCommand cmd = new System.Data.Odbc.OdbcCommand(sql, conn);
                reader = cmd.ExecuteReader();

                Console.WriteLine("Database = {0} \nDriver = {1} \nQuery {2}\nConnection String = {3}\nServer Version = {4}\nDataSource = {5}",
                conn.Database, conn.Driver, cmd.CommandText, conn.ConnectionString, conn.ServerVersion, conn.DataSource);

                while (reader.Read())
                {
                    Console.WriteLine("{0}|{1}|{2}",
                        reader["MT-CODE"],
                        reader["HST-DESCRIPTION"],
                        reader["ST-LOCATION"]);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e);
            }
            finally
            {
                reader.Close();
                conn.Close();
                Console.WriteLine("End.");
            }
            
        }
    }
}
