using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveFrom_File
{
    public partial class Form1 : Form
    {
        //http://www.aspsnippets.com/Articles/Read-and-Import-Excel-File-into-DataSet-or-DataTable-using-C-and-VBNet-in-ASPNet.aspx
        //http://www.aspsnippets.com/Articles/Import-data-from-Excel-file-to-Windows-Forms-DataGridView-using-C-and-VBNet.aspx
        //https://msdn.microsoft.com/en-us/library/ex21zs8x(v=vs.110).aspx
        //http://www.aspsnippets.com/Articles/Using-SqlBulkCopy-to-import-Excel-SpreadSheet-data-into-SQL-Server-in-ASPNet-using-C-and-VBNet.aspx
        //http://www.aspsnippets.com/Articles/Import-data-from-Excel-file-to-Windows-Forms-DataGridView-using-C-and-VBNet.aspx
        //http://www.codeproject.com/Questions/445400/Read-Excel-Sheet-Data-into-DataTable

        public Form1()
        {
            InitializeComponent();
            comboBox1_Loaded();
        }
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        string sqlConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True";
        string databaseConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True";




        private string openDB()
        { //get data from Table1 using sql
            string sqlConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True";
            SqlConnection myConnection = new SqlConnection(sqlConnectionString);
            //myConnection.ConnectionString = sqlConnectionString;
            try
            {
                string sqlQuery = "select * from Table1";
                SqlCommand com = new SqlCommand(sqlQuery, myConnection);
                com.Connection.Open();
                string rez = " ";
                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            rez += reader.GetValue(i).ToString();
                        }
                    }
                }
                com.Connection.Close();
                MessageBox.Show("Read done");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Can not open connection ! ");
            }
            return sqlConnectionString;
        }
        /// <summary>get data from Table1 using sql
        /// get data from Table1 using sql
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            //get data from Table1 using sql
            string sqlConnectionString = openDB();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            OpenFileDialog dlg = new OpenFileDialog();
            //dlg.Filter = "Excel files | *.xls";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string filePath = dlg.FileName;
                string extension = Path.GetExtension(filePath);
                string conStr, sheetName;

                conStr = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03
                        conStr = string.Format(Excel03ConString, filePath);
                        break;
                    case ".xlsx": //Excel 07 to later

                        conStr = string.Format(Excel07ConString, filePath, "no");
                        break;
                }
                //Read Data from the Sheet.
                using (OleDbConnection con = new OleDbConnection(conStr))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        using (OleDbDataAdapter oda = new OleDbDataAdapter())
                        {
                            List<string> lstSheets1 = new List<string>();
                            //excelConnection.Open();
                            con.Open();
                            DataTable dt1 = con.GetSchema("Tables");
                            foreach (DataRow dr in dt1.Rows)
                            {
                                lstSheets1.Add(dr["TABLE_NAME"].ToString());
                            }
                            con.Close();
                            sheetName = lstSheets1[0].ToString();
                            dt = new DataTable();
                            dt1 = new DataTable();
                            //cmd.CommandText = "SELECT * From [Sheet1$]";
                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            oda.Fill(dt1);
                            con.Close();
                        }
                    }
                }
                string sqlConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True";
                if (dt != null)
                {
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnectionString))
                    {
                        string sqlConnectionString1 = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True";
                        SqlConnection myConnection = new SqlConnection();
                        myConnection.ConnectionString = sqlConnectionString1;

                        bulkCopy.DestinationTableName = "PRICETEST2";
                        //bulkCopy.ColumnMappings.Add("[Id]", "Id");
                        //bulkCopy.ColumnMappings.Add("Name", "Name");
                        //bulkCopy.ColumnMappings.Add("Salary", "Salary");

                        myConnection.Open();
                        bulkCopy.WriteToServer(dt);

                        myConnection.Close();
                        MessageBox.Show("Upload Successfull!");
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            string headerFile = radioButton1.Checked ? "yes" : "no";

            //file upload path            
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files (.xls)|*.xls; *.xlsx; *.csv |All Files (*.*)|*.*";
            dlg.Title = "Please select file";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string filePath = dlg.FileName;
                string extension = Path.GetExtension(filePath);
                string conStr, sheetName;
                conStr = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03
                        conStr = string.Format(Excel03ConString, filePath, headerFile);
                        break;
                    case ".xlsx": //Excel 07 to later
                        conStr = string.Format(Excel07ConString, filePath, headerFile);
                        break;
                }
                //Create connection string to Excel work book
                string excelConnectionString = conStr;
                //Create Connection to Excel work book
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                //Create OleDbCommand to fetch data from Excel
                List<string> lstSheets = new List<string>();
                excelConnection.Open();
                DataTable dt = excelConnection.GetSchema("Tables");
                foreach (DataRow dr in dt.Rows)
                {
                    lstSheets.Add(dr["TABLE_NAME"].ToString());
                }
                sheetName = lstSheets[0];
                dt.Columns.AddRange(new DataColumn[3] { 
                new DataColumn("Id", typeof(int)),
                new DataColumn("Name", typeof(string)),
                new DataColumn("Salary",typeof(int)) });
                string queryExcel = "SELECT  * FROM [" + lstSheets[0] + "]";
                using (OleDbDataAdapter oda = new OleDbDataAdapter(queryExcel, excelConnection))
                {
                    oda.Fill(dt);
                }
                //excelConnection.Close();
                OleDbCommand cmd = new OleDbCommand(queryExcel, excelConnection);
                OleDbDataAdapter adp = new OleDbDataAdapter(cmd);
                DataTable dataT = new DataTable();
                cmd.CommandText = queryExcel;
                adp.SelectCommand = cmd;
                adp.Fill(dataT);

                //excelConnection.Close();
                //excelConnection.Open();
                using (OleDbDataReader dReader = cmd.ExecuteReader())
                {
                    var aaa = dReader;
                    using (SqlConnection myConnection1 = new SqlConnection(sqlConnectionString))
                    {
                        using (SqlBulkCopy sqlBulk = new SqlBulkCopy(databaseConnectionString))
                        {
                            //Give your Destination table name 
                            sqlBulk.DestinationTableName = "PRICETEST2";
                            sqlBulk.ColumnMappings.Add("Id", "Id");
                            sqlBulk.ColumnMappings.Add("Name", "Name");
                            sqlBulk.ColumnMappings.Add("Salary", "Salary");
                            //sqlBulk.WriteToServer(dataT);
                            sqlBulk.WriteToServer(dReader);
                        }
                    }
                }
                excelConnection.Close();
                MessageBox.Show("Upload Successfull!");
            }
        }
        public void comboBox1_Loaded()
        {
            SqlConnection myConnection = new SqlConnection(sqlConnectionString);
            myConnection.Open();
            DataTable tableDB = myConnection.GetSchema("Tables");
            List<string> tables = new List<string>();
            tables.Add("");
            foreach (DataRow row in tableDB.Rows)
            {
                string tablename = (string)row[2];
                tables.Add(tablename);
            }

            myConnection.Close();
            comboBox1.DataSource = tables;
        }
    }
}
