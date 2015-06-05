using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace ExcelProject
{
    class ExcelLogic
    {
        private static string connectionStringDefault = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Excel_Automation_Training\Projektliste GiB.xlsx;Extended Properties=""Excel 12.0;HDR=YES;""";
        private static OleDbDataAdapter da;
        private static DataTable dt;
        private static string[] headers;
        private static string[] bearbeiterValues;

        private static OleDbConnection CreateConnection()
        {
            OleDbConnection conn = new OleDbConnection(connectionStringDefault);
            return conn;
        }


        public static DataTable FillDataTable()
        {
            dt = new DataTable();
            using (OleDbConnection conn = CreateConnection())
            {
                conn.Open();
                using (da = new OleDbDataAdapter(@"select * from [Tabelle1$]", conn))
                {

                    da.Fill(dt);
                    conn.Close();
                    //dt.PrimaryKey = new DataColumn[] { dt.Columns["Projektname"] };

                }
                headers = dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                bearbeiterValues = dt.AsEnumerable().Select(s => s.Field<string>("Bearbeiter")).Distinct().ToArray<string>();
            }

            return dt;
        }

        public static string[] GetExcelColumns()
        {
            return headers;
        }

        public static string[] GetBearbeiterValues()
        {
            return bearbeiterValues;
        }

        private static Dictionary<string, string> FilterChangedValues(Dictionary<string, string> original, Dictionary<string, string> current)
        {
            Dictionary<string, string> onlyChangedValues = new Dictionary<string, string>();
            foreach (string columnname in original.Keys)
            {
                if (!original[columnname].Equals(current[columnname]))
                {
                    onlyChangedValues.Add(columnname, current[columnname]);
                }
            }

            return onlyChangedValues;
        }

        public static DataTable UpdateExcelFile(Dictionary<string, string> original, Dictionary<string, string> current)
        {
            Dictionary<string, string> changedValues = ExcelLogic.FilterChangedValues(original,current);

            string updateCommand = @"Update [Tabelle1$] set";
            foreach (string columnname in changedValues.Keys)
            {
                updateCommand = updateCommand + " " + columnname + " = '" + changedValues[columnname] + "'" + " ,";
            }

            updateCommand = updateCommand.Substring(0, updateCommand.Length - 2);

            updateCommand = updateCommand + " where Projektname = '" + original["Projektname"] + "'";

            int columnsaffected = 0;
            using (OleDbConnection conn = CreateConnection())
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand(updateCommand, conn))
                {

                    columnsaffected = cmd.ExecuteNonQuery();                
                    conn.Close();

                }
  
            }
            dt = FillDataTable();
            System.Windows.Forms.MessageBox.Show(columnsaffected + " columns updated");
            return dt;
        }
     
        
    }
}
