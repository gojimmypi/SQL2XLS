// add assembly reference System.Configuration
using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data; // see https://github.com/gojimmypi/PatternsAndPractices
using System.Configuration;
using System.IO;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Text.RegularExpressions;

// see also https://github.com/JanKallman/EPPlus

namespace SQL2XLS
{
    class Program
    {
        private static string ServerName = "localhost";
        private static string DatabaseName = "common";
        private static bool blnShowHelp = false;
        private static bool blnDebugMode = false;
        private static bool errorFound = false;
        private static bool verboseMode = false;

        private static DataSet ds;
        private static string FileID = "";
        private static string fieldFileID = "batch_id"; // we can optionall number the files

        #region config
        //***********************************************************************************************************************************
        private static string FILE_FORMAT()
        //***********************************************************************************************************************************
        {
            string res = System.Configuration.ConfigurationManager.AppSettings["FILE_FORMAT"];
            if (String.IsNullOrEmpty(res))
            {
                res = "XLSX";
            }
            switch (res)
            {
                case "CSV":
                case "XLSX":
                    break;

               default:
                    throw new Exception("Unsupported FILE_FORMAT.");
            }

            return res;
        }
        
        //***********************************************************************************************************************************
        private static bool OVERWRITE_EXISTING()
        //***********************************************************************************************************************************
        {
            bool res = false;
            string str = System.Configuration.ConfigurationManager.AppSettings["OVERWRITE_EXISTING"];
            if (String.IsNullOrEmpty(str))
            {
                res = false;
            }
            switch (str.ToUpper())
            {
                case "TRUE":
                    res = true;
                    break;
                case "FALSE":
                    res = false;
                    break;

                default:
                    throw new Exception("Unsupported OVERWRITE_EXISTING value");
            }

            return res;
        }

        //***********************************************************************************************************************************
        private static bool FIELD_HEADER()
        //***********************************************************************************************************************************
        {
            bool res = false;
            string str = System.Configuration.ConfigurationManager.AppSettings["FIELD_HEADER"];
            if (String.IsNullOrEmpty(str))
            {
                res = false;
            }
            switch (str.ToUpper())
            {
                case "TRUE":
                    res = true;
                    break;
                case "FALSE":
                    res = false;
                    break;

                default:
                    throw new Exception("Unsupported FIELD_HEADER value");
            }

            return res;
        }
        //***********************************************************************************************************************************
        private static string FILE_PATH()
        //***********************************************************************************************************************************
        {
            string res = System.Configuration.ConfigurationManager.AppSettings["FILE_PATH"];
            if (String.IsNullOrEmpty(res))
            {
                res = "c:\\temp\\";
            }
            if (!res.EndsWith("\\"))
            {
                res += "\\"; // ensure path ends with slash
            }
            return res;
        }

        //***********************************************************************************************************************************
        private static string FILE_ROOT(int n)
        //***********************************************************************************************************************************
        {
            string res = System.Configuration.ConfigurationManager.AppSettings["FILE_ROOT_" + n.ToString() ];
            if (String.IsNullOrEmpty(res))
            {
                res = "TABLE_" + n.ToString(); // if there's no preferred file name, use TABLE_[n]
            }
            return res;
        }

        //***********************************************************************************************************************************
        private static string FILE_SUFFIX(int n)
        //***********************************************************************************************************************************
        {
            string res = System.Configuration.ConfigurationManager.AppSettings["FILE_ROOT_" + n.ToString()];
            if (String.IsNullOrEmpty(res))
            {
                res = "TABLE_" + n.ToString(); // if there's no preferred file name, use TABLE_[n]
            }
            return res;
        }


        //***********************************************************************************************************************************
        private static string SQL_COMMAND()
        //***********************************************************************************************************************************
        {
            string res = System.Configuration.ConfigurationManager.AppSettings["SQL_COMMAND"];
            if (String.IsNullOrEmpty(res))
            {
                throw new Exception("ERROR: SQL_COMMAND config setting not found.");
            }
            return res;
        }
        #endregion


        #region SQL
        //***********************************************************************************************************************************
        static string ConnectionString(string strServerName, string strDatabaseName)
        //***********************************************************************************************************************************
        {
            // see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemDataSqlClientSqlConnectionClassConnectionStringTopic.asp
            //
            // there is some debate as to whether the Oledb provider is indeed faster than the native client!
            //  
            return "Workstation ID=SQL2XLS;" +
                   "packet size=8192;" +
                   "Persist Security Info=false;" +
                   "Server=" + strServerName + ";" +
                   "Database=" + strDatabaseName + ";" +
                   "Trusted_Connection=true; " +
                   "Network Library=dbmssocn;" +
                   "Pooling=True; " +
                   "Enlist=True; " +
                   "Connection Lifetime=14400; " +
                   "Max Pool Size=20; Min Pool Size=0";
        }

        //***********************************************************************************************************************************
        private static void getDatasetFromSQL()
        // reminder that a Dataset can have an arbirary number of DataTables; this depends on the SQL. 
        // typically only stored procs will return more than one table result (or a multipe selects separated with semicolons)
        //***********************************************************************************************************************************
        {
            string strSQL = SQL_COMMAND(); // the SQL command is in our config file
            FileID = ""; 
            ds = SqlHelper.ExecuteDataset(ConnectionString(ServerName, DatabaseName), CommandType.Text, strSQL);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                if (ds.Tables[i].Rows.Count < 1)
                {
                    Console.WriteLine("No data found in table[" + i.ToString() + "] " + FileID);
                }
                else
                {
                    if (FileID == "")
                    {
                        // check to see if we have an available unique identifier from the data to add as file suffix (e.g. Batch ID)
                        if (ds.Tables[i].Columns.IndexOf(fieldFileID) >= 0)
                        {
                            FileID = ds.Tables[i].Rows[0][fieldFileID].ToString();
                        }
                    }
                }
            }
            if (FileID == "")
            {
                Console.WriteLine("Field: [" + fieldFileID + "] not found in dataset. No file suffix will be added.");
                FileID = "";
            }
            if ((ds is null) || (ds.Tables.Count == 0)) 
            {
                throw new Exception("SQL Statement does did not return any data.");
            }
        }

        /// <summary>
        ///   write the contents of [dataTable] to a CSV file named [FileName]
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="FileName"></param>
        /// <param name="overwriteExisting"></param>
        private static void writeDataTable_TO_CsvFile(DataTable dataTable, string FileName, Boolean overwriteExisting = false, bool IncludeHeaderNames = false)
        {
            string columns = "";
            string thisColumnName = "";
            string thisValue = "";
            string values = "";
            int rowCount = 0;
            if (File.Exists(FileName))
            {
                if (overwriteExisting)
                {
                    Console.WriteLine("Deleting existing " + FileName);
                    File.Delete(FileName);
                }
                else
                {
                    throw new Exception("File " + FileName + " already exists!");
                }
            }

            using (StreamWriter writetext = new StreamWriter(FileName))
            {

                // get the list of column names:  "[col1],[col2]..."
                if (IncludeHeaderNames)
                {
                    for (int index = 0; index < dataTable.Columns.Count; index++)
                    {
                        thisColumnName = dataTable.Columns[index].ToString();
                        columns += (index == 0) ? thisColumnName
                                              : "," + thisColumnName; // build list of comma-separated column names ([name1], [name2]...)
                    }
                    writetext.WriteLine(columns);
                }


                // get the data 
                foreach (DataRow dr in dataTable.Rows)
                {
                    values = ""; // need to clear the values for each new row

                    for (int index = 0; index < dataTable.Columns.Count; index++)
                    {
                        thisValue = dr.ItemArray[index].ToString();
                        // build a string of explicit values to insert (this works but is vulnerable to problems with single quotes)
                        // values += (index == 0) ? "'" + Regex.Replace(thisValue, @"\t|\n|\r", "\"") + "'"
                        //                     : ", '" + Regex.Replace(thisValue, @"\t|\n|\r", "\"") + "'"; // build values:  "('val1','val2'...)"

                        values += (index == 0) ? "" : ","; // build values:  "(?,?...)"
                        values += thisValue;
                    }
                    writetext.WriteLine(values);
                    rowCount++;
                }

            }
            Console.WriteLine("Wrote " + rowCount.ToString() + " rows to " + FileName);
        }

        //***********************************************************************************************************************************
        private static void writeDataTable_TO_ExcelFile(DataTable dataTable, string FileName, Boolean overwriteExisting = false)
        //***********************************************************************************************************************************
        //
        // see: https://stackoverflow.com/questions/34922325/how-to-insert-a-new-row-in-excel-using-oledb-c-net
        // see: https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbdataadapter.insertcommand(v=vs.110).aspx
        //
        {
            OleDbConnection connection = null;
            OleDbCommand command = null;
            string connectionString = "";
            string columns = "";
            string columnCreate = "";
            string values = "";
            string thisColumnName = "";
            string thisValue = "";
            int rowCount = 0;

            if (File.Exists(FileName))
            {
                if (overwriteExisting)
                {
                    Console.WriteLine("Deleting existing " + FileName);
                    File.Delete(FileName);
                }
                else
                {
                    throw new Exception("File " + FileName + " already exists!");
                }
            }

            try
            {
                // get the list of column names:  "[col1],[col2]..."
                for (int index = 0; index < dataTable.Columns.Count; index++)
                {
                    thisColumnName = dataTable.Columns[index].ToString();
                    columns += (index == 0) ? "[" + thisColumnName + "]" 
                                          : ", [" + thisColumnName + "]"; // build list of comma-separated column names ([name1], [name2]...)

                    columnCreate += (index == 0) ? "[" + thisColumnName + "] TEXT(255)"
                                               : ", [" + thisColumnName + "] TEXT(255)";  // build list of comma-separated column definitions ([name1] TEXT(255), [name2] TEXT(255)...)
                }

                // connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0 Xml;READONLY=FALSE;ImportMixedTypes=Text;HDR=YES\";";
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0 Xml; IMEX=0; READONLY=FALSE; ImportMixedTypes=Text; HDR=YES\";";
                using (connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message == "ERROR [IM002] [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified")
                        {
                            Console.WriteLine("Driver needed: AccessDatabaseEngine_X64.exe");
                            Console.WriteLine(" see https://www.microsoft.com/en-us/download/details.aspx?id=13255");
                        }
                        throw new Exception(ex.Message);
                    }

                    using (command = connection.CreateCommand())
                    {
                        // here we do a brute force, manual command builder as the built-in one does not seem to generate the proper code :|
                        command.CommandText = "CREATE TABLE [Sheet1] (" + columnCreate + ")"; // note contrary to other places, "Sheet1" does NOT need the "$" suffic here!
                        command.ExecuteNonQuery();
                        rowCount = dataTable.Rows.Count;

                        foreach (DataRow dr in dataTable.Rows)
                        {
                            values = ""; // need to clear the values for each new row
                            
                            for (int index = 0; index < dataTable.Columns.Count; index++)
                            {
                                thisValue = dr.ItemArray[index].ToString();
                                // build a string of explicit values to insert (this works but is vulnerable to problems with single quotes)
                                // values += (index == 0) ? "'" + Regex.Replace(thisValue, @"\t|\n|\r", "\"") + "'"
                                //                     : ", '" + Regex.Replace(thisValue, @"\t|\n|\r", "\"") + "'"; // build values:  "('val1','val2'...)"

                                // build parameterized values with wildcard names (this does not work)
                                // values += (index == 0) ? "?"
                                //                     : ", ?"; // build values:  "(?,?...)"
                                // command.Parameters.Add(thisValue);

                                // build parameterized values with explicit value names (this method seems most robust)
                                // see https://stackoverflow.com/questions/7501354/escaping-special-characters-for-ms-access-query
                                values += (index == 0) ? "?"
                                                     : ", ?"; // build values:  "(?,?...)"
                                command.Parameters.AddWithValue(dataTable.Columns[index].ToString(), thisValue);
                            }
                            command.CommandText = string.Format("Insert into [Sheet1$] ({0}) values({1})", columns, values);
                            command.ExecuteNonQuery();
                            // clear all the parameters for the next loop
                            while (command.Parameters.Count > 0) {
                                command.Parameters.RemoveAt(0);
                            }
                        }
                    }
                    connection.Close();

                }
                Console.WriteLine("Completed writing " + rowCount.ToString() +" rows to file: " + FileName);
            }
            catch (Exception ex)
            {
                errorFound = true;
                Console.WriteLine("SQLtoXLS Error: " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        #endregion


        #region NonWorkingCode
        // this section has sample code that was expected to work, but did not. For reference only
        private static void NonWorking_writeDataOdbc(string SaveAsFileName)
        {
            // sample code for writing ODBC, but fails with unexpected error (we end up manually assembling statements that should be automatically generated: see writeDataTable_TO_ExcelFile
            // for reference only: System.Data.Odbc.OdbcException: 'ERROR [42000] [Microsoft][ODBC Excel Driver] Syntax error in INSERT INTO statement
            DataSet dataSet = new DataSet();

            string strConn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + (SaveAsFileName) + ";";
            using (OdbcConnection myConnection = new OdbcConnection(strConn))
            {
                string strExcelTableName = "[Sheet1$]";// must use the $ after the object you reference in the spreadsheet
                OdbcDataAdapter adapter = new OdbcDataAdapter();
                //adapter.SelectCommand = new OdbcCommand("SELECT [name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role] FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                adapter.SelectCommand = new OdbcCommand("SELECT * FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                // adapter.SelectCommand = new OdbcCommand("SELECT [name] FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                OdbcCommandBuilder builder = new OdbcCommandBuilder(adapter);

                try
                {
                    myConnection.Open();
                }
                catch (Exception ex)
                {
                    if (ex.Message == "ERROR [IM002] [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified")
                    {
                        Console.WriteLine("Driver needed: AccessDatabaseEngine_X64.exe");
                        Console.WriteLine(" see https://www.microsoft.com/en-us/download/details.aspx?id=13255");
                    }
                    throw new Exception(ex.Message);
                }



                adapter.Fill(dataSet, strExcelTableName);

                //while (ds.Tables[0].Columns.Count > 1)
                //{
                //    string name = ds.Tables[0].Columns[1].ColumnName;
                //    ds.Tables[0].Columns.Remove(name);
                //}
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    dataSet.Tables[0].Rows.Add(dr.ItemArray);
                }

                // adapter.InsertCommand = new OdbcCommand("INSERT INTO [Sheet1$] ([name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", myConnection);
                adapter.InsertCommand = builder.GetInsertCommand();
                // adapter.UpdateCommand = builder.GetUpdateCommand();

                adapter.Update(dataSet, strExcelTableName);
            }

            // myData.Update(ds);
            // myConnection.GetSchema("tables");
        }


        private static void NonWorking_writeDataOleDb(string SaveAsFileName)
        // https://www.connectionstrings.com/ace-oledb-12-0/
        {
            DataSet dataSet = new DataSet();

            // string strConn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + (SaveAsFileName) + ";";
            // string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + "; Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ";
            // string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + "; Extended Properties = 'Excel 12.0 Xml;HDR=YES;IMEX=1;READONLY=FALSE;ImportMixedTypes=Text'; ";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + ";Extended Properties=\"Excel 12.0 Xml;READONLY=FALSE;ImportMixedTypes=Text;HDR=YES\";";
            using (OleDbConnection myConnection = new OleDbConnection(strConn))
            {
                string strExcelTableName = "[Sheet1$]";// must use the $ after the object you reference in the spreadsheet
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                ////adapter.SelectCommand = new OdbcCommand("SELECT [name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role] FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                ////adapter.SelectCommand = new OdbcCommand("SELECT * FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                adapter.SelectCommand = new OleDbCommand("SELECT * FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);

                try
                {
                    myConnection.Open();
                }
                catch (Exception ex)
                {
                    if (ex.Message == "ERROR [IM002] [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified")
                    {
                        Console.WriteLine("Driver needed: AccessDatabaseEngine_X64.exe");
                        Console.WriteLine(" see https://www.microsoft.com/en-us/download/details.aspx?id=13255");
                    }
                    throw new Exception(ex.Message);
                }

                adapter.Fill(dataSet, strExcelTableName);

                adapter.InsertCommand = builder.GetInsertCommand(true);
                string insertCommandTemplate = adapter.InsertCommand.CommandText;
                insertCommandTemplate = insertCommandTemplate.Replace("Sheet1$", "[Sheet1$]");
                string insertCommand = "";
                var regex = new Regex(Regex.Escape("?"));

                foreach (DataRow dr in ds.Tables[0].Rows)
                {

                    insertCommand = insertCommandTemplate;
                    foreach (object item in dr.ItemArray)
                    {
                        insertCommand = regex.Replace(insertCommand, "'" + item.ToString() + "'", 1);
                    }
                    OleDbCommand command = new OleDbCommand(insertCommand, myConnection);
                    command.CommandText = insertCommand;
                    command.ExecuteNonQuery();
                }

                // adapter.InsertCommand = new OdbcCommand("INSERT INTO [Sheet1$] ([name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", myConnection);
                // adapter.InsertCommand = builder.GetInsertCommand();
                // adapter.UpdateCommand = builder.GetUpdateCommand();
                myConnection.Close();
            }

            //myData.Update(ds);
            //myConnection.GetSchema("tables");
        }

        private static void NonWorking_writeDataOleDbFull(string SaveAsFileName)
        // for reference only; this one causes: Syntax error in INSERT INTO statement.
        // https://www.connectionstrings.com/ace-oledb-12-0/
        {
            DataSet dataSet = new DataSet();

            // string strConn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + (SaveAsFileName) + ";";
            // string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + "; Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ";
            //string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + "; Extended Properties = 'Excel 12.0 Xml;HDR=YES;IMEX=1;READONLY=FALSE;ImportMixedTypes=Text'; ";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SaveAsFileName + ";Extended Properties=\"Excel 12.0 Xml;READONLY=FALSE;ImportMixedTypes=Text;HDR=YES\";";
            using (OleDbConnection myConnection = new OleDbConnection(strConn))
            {
                string strExcelTableName = "[Sheet1$]";// must use the $ after the object you reference in the spreadsheet
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                ////adapter.SelectCommand = new OdbcCommand("SELECT [name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role] FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                ////adapter.SelectCommand = new OdbcCommand("SELECT * FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                adapter.SelectCommand = new OleDbCommand("SELECT * FROM " + strExcelTableName + " WHERE 1=0", myConnection);
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);

                try
                {
                    myConnection.Open();
                }
                catch (Exception ex)
                {
                    if (ex.Message == "ERROR [IM002] [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified")
                    {
                        Console.WriteLine("Driver needed: AccessDatabaseEngine_X64.exe");
                        Console.WriteLine(" see https://www.microsoft.com/en-us/download/details.aspx?id=13255");
                    }
                    throw new Exception(ex.Message);
                }


                adapter.Fill(dataSet, strExcelTableName);

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    dataSet.Tables[0].Rows.Add(dr.ItemArray);
                }

                // adapter.InsertCommand = new OdbcCommand("INSERT INTO [Sheet1$] ([name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", myConnection);
                adapter.InsertCommand = builder.GetInsertCommand();
                // adapter.UpdateCommand = builder.GetUpdateCommand();

                adapter.Update(dataSet, strExcelTableName);


                // adapter.InsertCommand = new OdbcCommand("INSERT INTO [Sheet1$] ([name],[principal_id],[sid],[type],[type_desc],[is_disabled],[create_date],[modify_date],[default_database_name],[default_language_name],[credential_id],[owning_principal_id],[is_fixed_role]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", myConnection);
                //adapter.InsertCommand = builder.GetInsertCommand();
                // adapter.UpdateCommand = builder.GetUpdateCommand();
                myConnection.Close();
            }


        }

        #endregion

        #region 
        //'***********************************************************************************************************************************
        //'
        //'***********************************************************************************************************************************
        //Private Sub SetSchemaINI(SaveAsFileName As String)
        //    Dim strThisPath As String = System.IO.Path.GetDirectoryName(SaveAsFileName)
        //    If strThisPath = "" Then strThisPath = "."
        //    Dim myINIHelper As New INIHelper(strThisPath & "\schema.ini")
        //    Dim strSection As String = "[" & System.IO.Path.GetFileName(SaveAsFileName) & "]"
        //    myINIHelper.Value(strSection, "MaxScanRows") = "0" ' ensure all rows are scanned for proper data type check
        //    myINIHelper.WriteIniFile()
        //End Sub

        #endregion

        #region Main
        //***********************************************************************************************************************************
        static void initMain()
        //***********************************************************************************************************************************
        {
            string[] args = Environment.GetCommandLineArgs();
            string[] argumentOptions;
            foreach (string arg in args)
            {
                argumentOptions = arg.Split(new Char[] { ':' }, 2);
                switch (argumentOptions[0].ToUpper().Trim())
                {
                    case "/SERVER":
                        if (argumentOptions.GetUpperBound(0) > 0)
                        {
                            ServerName = argumentOptions[1];
                        }
                        else
                        {
                            errorFound = true;
                            throw new Exception("ERROR Server name not specified.");
                        }
                        break;

                    case "/DATABASE":
                        if (argumentOptions.GetUpperBound(0) > 0)
                        {
                            DatabaseName = argumentOptions[1];
                        }
                        else
                        {
                            errorFound = true;
                            throw new Exception("ERROR Database name not specified.");
                        }
                        break;
                    case "/VERBOSE":
                        verboseMode = true;
                        break;

                    case "/DEBUG":
                        blnDebugMode = true;
                        break;

                    case "FIELDFILEID":
                        if (argumentOptions.GetUpperBound(0) > 0)
                        {
                            fieldFileID = argumentOptions[1];
                        }
                        else
                        {
                            errorFound = true;
                            throw new Exception("ERROR fieldFileID not specified.");
                        }
                        break;
                        
                    case "/?":
                        blnShowHelp = true;
                        Console.WriteLine("SQL2XLS [/SERVER:localhost] [/DATABASE:common] ");
                        Console.WriteLine("");
                        Console.WriteLine("  /SERVER      SQL Server name (default is localhost)");
                        Console.WriteLine("");
                        Console.WriteLine("  /DATABASE    SQL Database name (default is common)");
                        Console.WriteLine("");
                        Console.WriteLine("  /VERBOSE     Verbose mode.");
                        Console.WriteLine("");
                        Console.WriteLine("  /FIELDFILEID Field name to be used to search for File ID Suffix (e.g. batch_id.");
                        Console.WriteLine("");
                        Console.WriteLine("example: sql2xls /SERVER:myTargetServer");
                        Console.WriteLine("");
                        Console.WriteLine("see sql2xls.exe.config for SQL statemnts");
                        Console.WriteLine("");
                        Console.WriteLine("  /?           Show this help screen.");
                        break;

                    default:
                        // throw new Exception("Parameter unknown: " + argumentOptions[0]);
                        break;
                }
            }
            if (blnDebugMode)
            {
                Console.WriteLine("Debug mode. Connect debugger and edit myBreak to continue...");
                int myBreak = 0;
                while (myBreak == 0)
                {
                    Console.WriteLine("Waiting for debug myBreak value to change...");
                    System.Threading.Thread.Sleep(5000);
                }
            }
        }

        static void showConfig()
        {
            Console.WriteLine(" SQL2XLS"); 
            Console.WriteLine("");
            Console.WriteLine(" Server:    " + ServerName);
            Console.WriteLine(" Database:  " + DatabaseName);
            Console.WriteLine(" Format:    " + FILE_FORMAT().ToString());
            Console.WriteLine(" Overwrite: " + OVERWRITE_EXISTING().ToString());
            Console.WriteLine("");
        }

        //***********************************************************************************************************************************
        //***********************************************************************************************************************************
        static void Main(string[] args)
        //***********************************************************************************************************************************
        //***********************************************************************************************************************************
        {
            initMain();
            showConfig();
            if (!blnShowHelp)
            {
                getDatasetFromSQL();
                if (!Directory.Exists(FILE_PATH()))
                {
                    Console.WriteLine("Creating directory: " + FILE_PATH());
                    Directory.CreateDirectory(FILE_PATH());
                }
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    if (FILE_FORMAT() == "CSV")
                    {
                        writeDataTable_TO_CsvFile(ds.Tables[i], 
                                                  FILE_PATH() + FILE_ROOT(i) + FileID + ".csv", 
                                                  OVERWRITE_EXISTING(),
                                                  FIELD_HEADER()
                                                  );
                    }
                    else
                    {
                        writeDataTable_TO_ExcelFile(ds.Tables[i], 
                                                    FILE_PATH() + FILE_ROOT(i) + FileID + ".xlsx", 
                                                    OVERWRITE_EXISTING()
                                                   );
                    }
                }
            }
        }
        #endregion
    }
}
