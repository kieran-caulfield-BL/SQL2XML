using System;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Data.Odbc;
using System.Text.RegularExpressions;


namespace SQL2XML
{
    class Program
    {
        static void Main(string[] args)
        {
            // we can have 2 arguments, the client code and the matter code
            // if the matter code is blank then we search all matters (this is for 2nd dev)

            Regex rgx = new Regex(@"[0-9]{6}-[0-9]{6}");

            // Test if input arguments were supplied.
            if (args.Length == 0)
            {
                Console.WriteLine("Please enter a Matter Code XXXXXX-0000XX.");
                Environment.Exit(1);
            }

            if (!rgx.IsMatch(args[0].ToString()))
            {
                Console.WriteLine("Please enter a Matter Code in format 6 numbers with '-' followed by 6 numbers.");
                Environment.Exit(1);
            }

            String inputMatterCode = args[0].ToString();

            string connString = @"dsn=2600;UID=sysprogress;PWD=sysprogress";

            string sql = @"Select
    PUB.MATDB.""MT-CODE"",
    PUB.MATDB.""CL-CODE"",
    PUB.MATDB.DESCRIPTION As ""MAT-DESCRIPTION"",
    PUB.MATDB.IMPLEMENTATION,
    PUB.MATDB.""DATE-OPENED"",
    PUB.MATDB.""MT-TYPE"",
    PUB.CLIDB.CONTACT1,
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
    PUB.""FILE-LOCATION"" On PUB.""FILE-LOCATION"".""LOC-NAME"" = PUB.""DOC-CONTROL"".""ST-LOCATION"" Inner Join
    PUB.CLIDB On PUB.CLIDB.""CL-CODE"" = PUB.MATDB.""CL-CODE""
Where
    PUB.MATDB.""MT-CODE"" = ? And
    PUB.HISTORY.""DOCUMENT-TYPE"" Is Not Null And
    PUB.HISTORY.""DOCUMENT-TYPE"" <> ''
order by 
    PUB.MATDB.""CL-CODE"",
    PUB.MATDB.""MT-CODE"",
    PUB.HISTORY.""DATE-INSERTED"" ASC
";
            OdbcConnection conn = null;
            OdbcDataReader reader = null;

            try
            {
                // open connection
                conn = new OdbcConnection(connString);
                conn.Open();
                // execute the SQL
                OdbcCommand cmd = new OdbcCommand(sql, conn);
                cmd.Parameters.Add("MatterIdentifier", OdbcType.VarChar).Value = inputMatterCode; // "example 081389-000002"
                reader = cmd.ExecuteReader();

                Console.WriteLine("Database = {0} \nDriver = {1} \nQuery {2}\nConnection String = {3}\nServer Version = {4}\nDataSource = {5}",
                conn.Database, conn.Driver, cmd.CommandText, conn.ConnectionString, conn.ServerVersion, conn.DataSource);

                XmlDocument xmlDoc = new XmlDocument();

                XmlNode rootNode = xmlDoc.CreateElement("SolcaseDocs");
                xmlDoc.AppendChild(rootNode);

                XmlNode clientNode = xmlDoc.CreateElement("Client");
                    XmlAttribute clientCode = xmlDoc.CreateAttribute("CL-CODE");
                    XmlAttribute clientName = xmlDoc.CreateAttribute("CL-NAME");

                XmlNode matterNode = xmlDoc.CreateElement("Matter");
                    XmlAttribute matterCode = xmlDoc.CreateAttribute("MT-CODE");
                    XmlAttribute matterDesc = xmlDoc.CreateAttribute("MAT-DESCRIPTION");


                Boolean clientNodeNotCreated = true;
                String currentMatterCode= ""; // on change of matter we need to create another matter node

                while (reader.Read())
                {
                    if(clientNodeNotCreated)
                    {
                        // only once used, turn off once processed                        
                            clientCode.Value = reader["CL-CODE"].ToString();
                            clientNode.Attributes.Append(clientCode);
                        
                            clientName.Value = reader["CONTACT1"].ToString();
                            clientNode.Attributes.Append(clientName);
                        rootNode.AppendChild(clientNode);
                        clientNodeNotCreated = false;
                    }


                    if(currentMatterCode != reader["MT-CODE"].ToString())
                    {
                        // the matter code has changed, create a new Matter Node                       
                            matterCode.Value = reader["MT-CODE"].ToString();
                            matterNode.Attributes.Append(matterCode);
                        
                            matterDesc.Value = reader["MAT-DESCRIPTION"].ToString();
                            matterNode.Attributes.Append(matterDesc);
                        clientNode.AppendChild(matterNode);
                    }

                    XmlNode soldocNode = xmlDoc.CreateElement("SolDoc");
                        XmlAttribute histDesc = xmlDoc.CreateAttribute("HST-DESCRIPTION");
                        XmlAttribute histNo = xmlDoc.CreateAttribute("HISTORY-NO");
                        XmlAttribute docName = xmlDoc.CreateAttribute("DOCUMENT-NAME");
                        XmlAttribute solcaseDocType = xmlDoc.CreateAttribute("DOCUMENT-TYPE");
                        XmlAttribute dateInserted = xmlDoc.CreateAttribute("DATE-INSERTED");
                        XmlAttribute histFE = xmlDoc.CreateAttribute("LEVEL-FEE-EARNER");
                        XmlAttribute docGroup = xmlDoc.CreateAttribute("DOC-GROUP");
                        XmlAttribute stLocation = xmlDoc.CreateAttribute("ST-LOCATION");
                        XmlAttribute subPath = xmlDoc.CreateAttribute("SUB-PATH");
                        XmlAttribute actualDocType = xmlDoc.CreateAttribute("DOC-TYPE");
                        XmlAttribute docExt = xmlDoc.CreateAttribute("EXTENSION");
                        XmlAttribute docVersion = xmlDoc.CreateAttribute("VERSION");
                        XmlAttribute fileLoc = xmlDoc.CreateAttribute("LOC-DESCRIPTION");
                        XmlAttribute fileLocType = xmlDoc.CreateAttribute("LOC-TYPE");

                        // assign values from result set reader
                        histNo.Value = reader["HISTORY-NO"].ToString();
                        histDesc.Value = reader["HST-DESCRIPTION"].ToString();
                        docName.Value = reader["DOCUMENT-NAME"].ToString();
                        solcaseDocType.Value = reader["DOCUMENT-TYPE"].ToString();
                        dateInserted.Value = reader["DATE-INSERTED"].ToString();
                        histFE.Value = reader["LEVEL-FEE-EARNER"].ToString();
                        docGroup.Value = reader["DOC-GROUP"].ToString();
                        stLocation.Value = reader["ST-LOCATION"].ToString();
                        subPath.Value = reader["SUB-PATH"].ToString();
                        actualDocType.Value = reader["DOC-TYPE"].ToString();
                        docExt.Value = reader["EXTENSION"].ToString();
                        docVersion.Value = reader["VERSION"].ToString();
                        fileLoc.Value = reader["LOC-DESCRIPTION"].ToString();
                        fileLocType.Value = reader["LOC-TYPE"].ToString();

                        // assign attributes to soldocNode
                        soldocNode.Attributes.Append(histNo);
                        soldocNode.Attributes.Append(histDesc);
                        soldocNode.Attributes.Append(docName);
                        soldocNode.Attributes.Append(solcaseDocType);
                        soldocNode.Attributes.Append(dateInserted);
                        soldocNode.Attributes.Append(histFE);
                        soldocNode.Attributes.Append(docGroup);
                        soldocNode.Attributes.Append(stLocation);
                        soldocNode.Attributes.Append(subPath);
                        soldocNode.Attributes.Append(actualDocType);
                        soldocNode.Attributes.Append(docExt);
                        soldocNode.Attributes.Append(docVersion);
                        soldocNode.Attributes.Append(fileLoc);
                        soldocNode.Attributes.Append(fileLocType);

                    // append node to matter
                    matterNode.AppendChild(soldocNode);

                    // flush xml doc
                    

                    Console.WriteLine("{0}|{1}|{2}",
                        reader["MT-CODE"],
                        reader["HST-DESCRIPTION"],
                        reader["ST-LOCATION"]);
                }

                xmlDoc.Save("test-doc.xml");
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
