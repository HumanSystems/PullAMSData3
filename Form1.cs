using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using ClosedXML.Excel;
using System.IO;


namespace PullAMSData3
{
    public partial class Form1 : Form
    {

        private BindingSource bindingSource1 = new BindingSource();

        public Form1()
        {
            InitializeComponent();
        }

        //must be .NET Framework application to use 
        private void btnGetAMSData_Click(object sender, EventArgs e)
        {
            //OleDbConnection connection = new OleDbConnection("Provider=VFPOLEDB.1;Data Source=F:\\Gutters\\Data\\database.dbc;"


            //Free Table Directory no work for table
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data");

            //Single dbf file
            // OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\lots.dbf");


            //****************************************************************************************************************************************
            //* BEGIN Good section get list of table.columns
            //****************************************************************************************************************************************
            //Database Container: works for OleDbMetaDataCollectionNames but not table:
            //   odbcCon = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");
            //--
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");

            //try
            //{
            //    connection.Open();
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            //}

            //try
            //{
            //    DataTable tables = connection.GetSchema(
            //    System.Data.OleDb.OleDbMetaDataCollectionNames.Tables);

            //    foreach (System.Data.DataRow rowTables in tables.Rows)
            //    {
            //        Console.Out.WriteLine(rowTables["table_name"].ToString());
            //        DataTable columns = connection.GetSchema(
            //            System.Data.OleDb.OleDbMetaDataCollectionNames.Columns,
            //            new String[] { null, null, rowTables["table_name"].ToString(), null }
            //        );
            //        foreach (System.Data.DataRow rowColumns in columns.Rows)
            //        {
            //            Console.Out.WriteLine(
            //                rowTables["table_name"].ToString() + "." +
            //                rowColumns["column_name"].ToString() + " = " +
            //                rowColumns["data_type"].ToString()
            //            );
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to get AMS schema. Error is: " + ex.Message);
            //}
            //--
            //****************************************************************************************************************************************
            //*  END Good section get list of table.columns
            //****************************************************************************************************************************************


            //****************************************************************************************************************************************
            //* BEGIN Good section get list of tables
            //****************************************************************************************************************************************
            ////Database Container: works for OleDbMetaDataCollectionNames but not table:
            ////   odbcCon = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");

            //try
            //{
            //    connection.Open();
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            //}

            //try
            //{
            //    DataTable tables = connection.GetSchema(
            //    System.Data.OleDb.OleDbMetaDataCollectionNames.Tables);

            //    foreach (System.Data.DataRow rowTables in tables.Rows)
            //    {
            //        Console.Out.WriteLine(rowTables["table_name"].ToString());
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to get tables. Error is: " + ex.Message);
            //}
            //****************************************************************************************************************************************
            //*  END Good section get list of tables
            //***********************************************************************************************************************


            if (dataGridViewAMSTables.SelectedRows.Count == 0)
            {
                MessageBox.Show("you did not select any tables to extract");
                return;
            }


            //****************************************************************************************************************************************
            //* try another way
            //****************************************************************************************************************************************
            ////Database Container: works for OleDbMetaDataCollectionNames but not table:
            ////   odbcCon = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");
           //DEVELOPMENT 
            OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");
            
            //PRODUCTION
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\aw$\\Data\\au.dbc");

            try
            {
                connection.Open();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            }


            foreach (DataGridViewRow row in dataGridViewAMSTables.SelectedRows)
            {
                if (row.Index == -1)
                {
                    continue;
                }

                //DON'T DO ANYTHING WITH VIEWS
                string tableType = row.Cells["TABLE_TYPE"].Value.ToString();
                if (tableType == "VIEW")
                {
                    continue;
                }

                OleDbCommand command = new OleDbCommand(@"select * from " + row.Cells["TABLE_NAME"].Value.ToString());
                if (chkboxCountsOnly.Checked)
                {
                    command = new OleDbCommand(@"select '" + row.Cells["TABLE_NAME"].Value.ToString() + ": ', count(*) from " + row.Cells["TABLE_NAME"].Value.ToString());
                }
                //OleDbCommand command = new OleDbCommand(@"select * from conslog");
                command.Connection = connection;

                string fileName;

                if (!chkboxCountsOnly.Checked)
                {
                    string tableName = row.Cells["TABLE_NAME"].Value.ToString();
                    GetData(command, tableName);
                }
                else
                {
                    try
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                reader.Read();
                                fileName = row.Cells["TABLE_NAME"].Value.ToString() + ": " + reader[1].ToString();
                            }
                            else
                            {
                                fileName = row.Cells["TABLE_NAME"].Value.ToString() + ": 0" ;
                            }
                            //reader.Close(); not needed because inside using?
                            Console.WriteLine(fileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Failed to get table count. Table is: " + row.Cells["TABLE_NAME"].Value.ToString() + ".  Error is: " + ex.Message);
                    }
                }
            }


            connection.Close();

            if (chkboxCountsOnly.Checked)
            {
                MessageBox.Show("You requested counts only so check the output console for selected table counts");
            }

            //****************************************************************************************************************************************
            //*  END Good section get list of tables
            //***********************************************************************************************************************



            //example to limit results
            //string sql = "select * from lots where RECNO() > 1000 and RECNO() < 1002";
            //OleDbCommand cmd = new OleDbCommand(sql, connection);

            //try
            //{
            //    DataTable YourResultSet = new DataTable();
            //    OleDbDataAdapter DA = new OleDbDataAdapter(cmd);
            //    DA.Fill(YourResultSet);
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to get AMS count. Error is: " + ex.Message);
            //}


            //********************************************************************************************************
            //string sql = "select cust_no from lots where RECNO() > 1000 and RECNO() < 1002";
            //OleDbCommand command = new OleDbCommand(sql, connection);

            //try
            //{
            //    connection.Open();
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            //}

            //OleDbDataReader reader = command.ExecuteReader();

            //while (reader.Read())
            //{
            //    Console.WriteLine(reader[0].ToString());
            //}
            //reader.Close();

            //connection.Close();
            //*****************************************************************************************************


            ////****************************************************************************************************************************************
            ////*  Testing section try to get data
            ////****************************************************************************************************************************************
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=;Data Source=\\10.1.1.17\\Data;Mode=Read;Collating Sequence=machine"))
            ////Context switch error - using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=;Data Source=\\10.1.1.17\\Data\\au.dbc;Mode=Read;Collating Sequence=machine"))
            ////Context switch error - using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=;Data Source=\\10.1.1.17\\Data\\lots.dbf;Mode=Read;Collating Sequence=machine"))
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data\\lots.dbf;Mode=Read;Collating Sequence=machine"))
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data\\au.dbc;Mode=Read;Collating Sequence=machine"))
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data\\au.dbc;Mode=Read"))

            ////https://www.connectionstrings.com/visual-foxpro/
            ////free table dir
            //using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data; Mode=Read; Collating Sequence = general"))

            ////single dbf
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data\\lots.dbf;Mode=Read; Collating Sequence = machine"))

            ////single db container
            ////using (OleDbConnection cn = new OleDbConnection(@"Provider=VFPOLEDB;Data Source=\\10.1.1.17\\Data\\au.dbc;Mode=Read; Collating Sequence = machine; Exclusive=No"))


            ////using (OleDbCommand cmd = new OleDbCommand("select cust_no from lots where RECNO() > 1000 and RECNO() < 1005", cn))
            ////using (OleDbCommand cmd = new OleDbCommand("select count(*) as 'Hello' from lots", cn)) //works may - need to check number hasn't changed!!!
            //using (OleDbCommand cmd = new OleDbCommand("select count(*) as 'Hello' from lots where RECNO() > 1000 and RECNO() < 1005", cn)) //works may - need to check number hasn't changed!!!
            ////using (OleDbCommand cmd = new OleDbCommand("select 'fuck'  as 'Hello' from lots  where RECNO() > 1000 and RECNO() < 1005", cn)) 
            ////using (OleDbCommand cmd = new OleDbCommand("select distinct consignor  as 'Hello' from lots", cn))
            //{
            //    cn.Open();
            //    //var reader = cmd.ExecuteReader();

            //    cmd.CommandText = @"select count(*) as 'Hello' from lots where RECNO() > 1000 and RECNO() < 1005";
            //    cmd.CommandType = CommandType.Text;

            //    OleDbDataReader reader = cmd.ExecuteReader();
            //    // do something with the reader. ie:
            //    // someDataTable.Load(reader);
            //    while (reader.Read())
            //    {
            //        Console.WriteLine(reader[0].ToString());  //465961  -- s/b 579163
            //    }
            //    cn.Close();
            //    reader.Close();
            //}


        }

        private void GetData(OleDbCommand command, string tableName)
        {
            try
            {
                string path;
                string filename;
                string fullname;
                string fullpath;

                using (OleDbDataReader reader = command.ExecuteReader())
                {

                    //reader.Read();
                    //string t = reader.GetDataTypeName(3);
                    //int cnt = 0;
                    //while (cnt < 120)
                    //{
                    //    Console.WriteLine(cnt.ToString() + ": " + reader.GetName(cnt) + " : " + reader.GetDataTypeName(cnt));
                    //    cnt++;
                    //}
                    IXLWorkbook thisWB = new XLWorkbook();
                    IXLWorksheet thisWS = thisWB.Worksheets.Add(tableName);

                    //thisWS.Cell(1, 1).Value = "FirstCol";
                    //thisWS.Cell(1, 1).Value = "Item_Number";
                    //thisWS.Cell(1, 2).Value = "Country_Lookup";

                    int wsCol = 1;
                    for (int title = 0; title < reader.FieldCount; title++)
                    {
                        //Console.WriteLine("name: " + reader.GetName(title).ToString());         // Gets the column name
                        //Console.WriteLine("type: " + reader.GetFieldType(title).ToString());    // Gets the column type
                        //Console.WriteLine("typename: " + reader.GetDataTypeName(title).ToString()); // Gets the column database type

                        thisWS.Cell(1, wsCol).Value = reader.GetName(title).ToString();

                        wsCol++;
                    }

                    int currentRow = 2;
                    int nbrFields = 0;
                    string fileName;


                    while (reader.Read())
                    {
                        //load spreadsheet data here
                        nbrFields = reader.FieldCount;
                        wsCol = 1;

                        for (int readerRowFieldNo = 0; readerRowFieldNo < reader.FieldCount; readerRowFieldNo++)
                        {
                            //Console.WriteLine("name: " + reader.GetName(title).ToString());         // Gets the column name
                            //Console.WriteLine("type: " + reader.GetFieldType(title).ToString());    // Gets the column type
                            //Console.WriteLine("typename: " + reader.GetDataTypeName(title).ToString()); // Gets the column database type

                            thisWS.Cell(currentRow, wsCol).Value = reader[readerRowFieldNo].ToString();

                            string test2 = reader[readerRowFieldNo].ToString();

                            wsCol++;
                        }

                        currentRow++;

                        //Console.WriteLine(String.Format("{0}", reader[0]));
                        //Console.WriteLine(String.Format("{0}", reader[1]));
                        //Console.WriteLine(String.Format("{0}", reader[2]));
                        //Console.WriteLine(String.Format("{0}", reader[3]));
                        //Console.WriteLine(String.Format("{0}", reader[4]));
                        //Console.WriteLine(String.Format("{0}", reader[5]));
                        //Console.WriteLine(String.Format("{0}", reader[6]));
                        //Console.WriteLine(String.Format("{0}", reader[7]));
                        //Console.WriteLine(String.Format("{0}", reader[8]));
                        //Console.WriteLine(String.Format("{0}", reader[9]));
                        //Console.WriteLine(String.Format("{0}", reader[10]));
                        //Console.WriteLine(String.Format("{0}", reader[11]));
                        //Console.WriteLine(String.Format("{0}", reader[12]));

                    }

                    //reader.Close(); not needed because using (OleDbDataReader reader = command.ExecuteReader()) ?? 

                    fileName = tableName;

                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Title = "Save .csv file for " + tableName,
                        //Filter = ".xlsx Files (*.xlsx)|*.xlsx",
                        //FileName = row.Cells["TABLE_NAME"].Value.ToString(),
                        FileName = fileName,
                        Filter = ".csv Files (*.csv)|*.csv",
                        OverwritePrompt = true
                    };

                    if (sfd.ShowDialog() != DialogResult.OK)
                    {
                        MessageBox.Show("You did not select a destination so process is terminating");
                        return;
                    }


                    //MessageBox.Show("Please wait while we save the file ......");
                    Cursor.Current = Cursors.WaitCursor;

                    path = Path.GetDirectoryName(sfd.FileName);
                    filename = Path.GetFileNameWithoutExtension(sfd.FileName);
                    fullname = Path.GetFileName(sfd.FileName);
                    fullpath = Path.GetFullPath(sfd.FileName);


                    Type tp;

                    //.Select(r => string.Join("|", r.Cells(1, lastCellAddress.ColumnNumber)
                    var lastCellAddress = thisWS.RangeUsed().LastCell().Address;
                    File.WriteAllLines(fullpath, thisWS.Rows(1, lastCellAddress.RowNumber)
                        .Select(r => string.Join(",", r.Cells(1, lastCellAddress.ColumnNumber)
                                .Select(cell =>
                                {
                                    var cellValue = cell.GetValue<string>();
                                    cellValue = cellValue.Replace("\"", "\"\"");
                                    tp = cellValue.GetType();
                                    if (tp.Equals(typeof(string)))   //just quote all strings instead of only when have comma ,
                                            {
                                        cellValue = "\"" + cellValue + "\"";
                                    }

                                    return cellValue;
                                            //return cellValue.Contains(",") ? $"\"{cellValue}\"" : cellValue;
                                        }))));

                    Cursor.Current = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to get tables. Table is: " + tableName + ".  Error is: " + ex.Message);
            }
        }

        private void btnODBC_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            dataGridViewAMSTables.Rows.Clear();


            string connectionStringCM = "DRIVER={MySQL ODBC 8.0 UNICODE Driver}; Server = 10.1.1.15; Database = kelleher; User = nmancini; Password = shannon123; Option = 3";

            String sql = "SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS";

            OdbcConnection ODBCcon = new OdbcConnection(connectionStringCM);
            ODBCcon.Open();
            OdbcCommand DbCommand = ODBCcon.CreateCommand();
            DbCommand.CommandText = sql;
            OdbcDataReader DbReader = DbCommand.ExecuteReader();

            if (DbReader.HasRows)
            {
                //System.Windows.Forms.MessageBox.Show("its already there numknuck");
                while (DbReader.Read())
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridViewAMSTables.Rows[0].Clone();
                    row.Cells[0].Value = DbReader.GetString(0);
                    dataGridViewAMSTables.Rows.Add(row);

                   // selectedConstInNo = DbReader.GetString(0);
                   // selectedConstInNo = DbReader.GetInt32(0).ToString();

                }

            }

            DbReader.Close();

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;


            //string sql = "select count(*) from lots";

            ////https://learn.microsoft.com/bs-latn-ba/previous-versions/sql/odbc/microsoft/using-the-vfp-foxpro-odbc-driver-with-your-visual-basic-application?view=sql-server-ver16&viewFallbackFrom=aps-pdw-2016
            //// System.Data.Odbc.OdbcConnection conn = new System.Data.Odbc.OdbcConnection("Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" + System.IO.Path.GetFullPath(strFileName).Replace(System.IO.Path.GetFileName(strFileName), "") + ";Exclusive=No");

            //System.Data.Odbc.OdbcConnection conn = new System.Data.Odbc.OdbcConnection("Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=\\10.1.1.17\\Data\\au.dbc;Exclusive=No");

            ////from shipreceipts: connectionStringCM = "DRIVER={ODBC}; Server = 10.1.1.15; Database = kelleher; User = nmancini; Password = shannon123; Option = 3";

            ////System.Data.Odbc.OdbcConnection conn = new System.Data.Odbc.OdbcConnection("Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" + System.IO.Path.GetFullPath(strFileName).Replace(System.IO.Path.GetFileName(strFileName), "") + ";Exclusive=No");
            //conn.Open();
            ////string strQuery = "SELECT * FROM [" + System.IO.Path.GetFileName(strFileName) + "]";
            //string strQuery = "select count(*) from lots";
            //System.Data.Odbc.OdbcDataAdapter adapter = new System.Data.Odbc.OdbcDataAdapter(strQuery, conn);
            //System.Data.DataSet ds = new System.Data.DataSet();
            //adapter.Fill(ds);
            ////return ds.Tables[0];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Production: AuctionWindows\Data
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\aw$\\Data\\au.dbc");

            ////Test: Auction\Data
            ////OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");

            //try
            //{
            //    connection.Open();
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            //}

            //try
            //{
            //    DataTable tables = connection.GetSchema(
            //    System.Data.OleDb.OleDbMetaDataCollectionNames.Tables);

            //    //foreach (System.Data.DataRow rowTables in tables.Rows)
            //    //{
            //    //    Console.Out.WriteLine(rowTables["table_name"].ToString());
            //    //}

            //    dataGridViewAMSTables.DataSource = bindingSource1;
            //    bindingSource1.DataSource = tables;

            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Failed to get tables. Error is: " + ex.Message);
            //}
        }

        private void btnSelectAllRows_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewAMSTables.Rows)
            {
                row.Selected = true;
            }
        }

        private void btnShowAMS_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            dataGridViewAMSTables.Rows.Clear();

            OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\aw$\\Data\\au.dbc");

            //Test: Auction\Data
            //OleDbConnection connection = new OleDbConnection(@"Provider=VFPOLEDB.1;Data Source=\\10.1.1.17\\Data\\au.dbc");

            try
            {
                connection.Open();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to make AMS connection. Error is: " + ex.Message);
            }

            try
            {
                DataTable tables = connection.GetSchema(
                System.Data.OleDb.OleDbMetaDataCollectionNames.Tables);

                //foreach (System.Data.DataRow rowTables in tables.Rows)
                //{
                //    Console.Out.WriteLine(rowTables["table_name"].ToString());
                //}

                dataGridViewAMSTables.DataSource = bindingSource1;
                bindingSource1.DataSource = tables;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to get tables. Error is: " + ex.Message);
            }


            txtDataType.Text = "AMS Tables";
            btnODBC.Visible = false;
            btnODBC.Enabled = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

        }
    }
}
