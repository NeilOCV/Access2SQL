using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml.Serialization;

namespace Access2SQL
{
    

    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        public List<string> GetTables(OleDbConnection conn)
        {
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            conn.Close();

            List<string> lst = new List<string>();
            foreach (DataRow row in schemaTable.Rows)
            {
                lst.Add(row[2].ToString());
            }

            return lst;
        }
        private string TranslateDataType(string strDataType)
        {
            //https://www.w3schools.com/asp/ado_datatypes.asp
            string result = "Unknown";

            if (strDataType=="20")
                result="BigInt";
            if (strDataType == "128")
                result = "Binary";
            if (strDataType == "11")
                result = "Boolean";
            if (strDataType == "129")
                result = "Char";
            if (strDataType == "6")
                result = "Currency";
            if (strDataType == "7")
                result = "Date";
            if (strDataType == "135")
                result = "DBTimeStamp";
            if (strDataType == "14")
                result = "Decimal";
            if (strDataType == "5")
                result = "Double";
            if (strDataType == "72")
                result = "GUID";
            if (strDataType == "9")
                result = "IDispatch";
            if (strDataType == "3")
                result = "Integer";
            if (strDataType == "205")
                result = "LongVarBinary";
            if (strDataType == "201")
                result = "LongVarChar";
            if (strDataType == "203")
                result = "LongVarWChar";
            if (strDataType == "131")
                result = "Numeric";
            if (strDataType == "4")
                result = "Single";
            if (strDataType == "2")
                result = "SmallInt";
            if (strDataType == "17")
                result = "UnsignedTinyInt";
            if (strDataType == "204")
                result = "VarBinary";
            if (strDataType == "200")
                result = "VarChar";
            if (strDataType == "12")
                result = "Variant";
            if (strDataType == "202")
                result = "VarWChar";
            if (strDataType == "130")
                result = "WChar";

            return result;
        }

        public List<DBSchema> GetAll(OleDbConnection conn,string table)
        {
            List<DBSchema> result = new List<DBSchema>();
            List<string> types = new List<string>();

            conn.Open();

            DataTable dbSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, table, null });
            foreach (DataRow row in dbSchema.Rows)
            {
                DBSchema obj = new DBSchema();
                obj.ColumnName = row["COLUMN_NAME"].ToString();
                obj.DataType = TranslateDataType(row["DATA_TYPE"].ToString());

                obj.DataValue = row[obj.ColumnName.ToString()].ToString();

                types.Add(obj.DataType);
                obj.TableName = table;
                result.Add(obj);
            }

            conn.Close();

            var prog = (from tb in types
                        select tb).Distinct();
            types = prog.ToList();

            if (types.Contains("Unknown"))
            {
                MessageBox.Show(table + " contains \"Unknown\"");
            }

            return result;
        }
        private List<string> GetColumnNames(OleDbConnection conn, string table)
        {
            List<string> result = new List<string>();
            OleDbDataReader dr;
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT column_name tn FROM information_schema.columns WHERE table_name = '" + table + "'", conn);

            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string column = dr.GetValue(0).ToString();
                result.Add(column);

            }

            conn.Close();
            return result;
        }
        private void btnGo_Click(object sender, EventArgs e)
        {
            string strAccessConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFileName.Text;
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            List<string> tables = GetTables(myAccessConn);
            foreach(string tbl in tables)
            {
                List<DBSchema> obj = GetAll(myAccessConn, tbl);
            }

            return;
            
        }
    }
}
