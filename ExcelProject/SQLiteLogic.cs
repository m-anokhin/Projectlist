using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace ExcelProject
{
    class SQLiteLogic //: DBLogic
    {
        private static string connectionStringDefault = @"Data Source=C:\Users\Administrator\Desktop\Project List\Projektliste GiB.sqlite;Version=3;";
        private static SQLiteDataAdapter da;
        private static DataTable dt;
        private static string[] headers;
        private static string[] bearbeiterValues;

        private static SQLiteConnection CreateConnection()
        {
            try
            {
                SQLiteConnection conn = new SQLiteConnection(connectionStringDefault);
                return conn;
            }
            catch (Exception ex)
            {
                
                return null;
            }
           
        }

        public static DataTable GetProjects()
        {
            dt = new DataTable();
            using (SQLiteConnection conn = CreateConnection())
            {
                conn.Open();
                using (da = new SQLiteDataAdapter(@"select * from Projects", conn))
                {
                    try
                    {
                        da.Fill(dt);
                        conn.Close();

                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                        return null;
                    }
                    
                }
                headers = dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                bearbeiterValues = dt.AsEnumerable().Select(s => s.Field<string>("Bearbeiter")).Distinct().ToArray<string>();
            }
            return dt;
        }

        public static string[] GetColumnNames()
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

        public static DataTable UpdateDataBaseFile(Dictionary<string, string> original, Dictionary<string, string> current)
        {
            Dictionary<string, string> changedValues = FilterChangedValues(original, current);

            string updateCommand = @"Update Projects set";
            foreach (string columnname in changedValues.Keys)
            {
                updateCommand = updateCommand + " " + columnname + " = '" + changedValues[columnname] + "'" + " ,";
            }

            updateCommand = updateCommand.Substring(0, updateCommand.Length - 2);

            updateCommand = updateCommand + " where Projektname = '" + original["Projektname"] + "'";

            int columnsaffected = 0;
            using (SQLiteConnection conn = CreateConnection())
            {
                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(updateCommand, conn))
                {

                    columnsaffected = cmd.ExecuteNonQuery();
                    conn.Close();

                }

            }
            dt = GetProjects();
            System.Windows.Forms.MessageBox.Show(columnsaffected + " columns updated");
            return dt;
        }
    
        public static int AddBearbeiterToTheDataBase(Dictionary<string, string> bearbeiterData)
        {
            int columnsaffected = 0;
            using (SQLiteConnection conn = CreateConnection())
            {
                //String insertQuery = "INSERT INTO dbo.SMS_PW (id,username,password,email) VALUES(idvalue,usernamevalue,passwordvalue,emailvalue)";
                
                StringBuilder part1 = new StringBuilder("INSERT INTO Projects(");
                StringBuilder part2 = new StringBuilder("VALUES(");
                int colNumb= bearbeiterData.Keys.Count;
                int colCounter=0;
                
                foreach (string columnName in bearbeiterData.Keys)
                {
                    if (colCounter!=colNumb-1)
                    {
                        part1.Append(columnName + ", ");
                        part2.Append("'" + bearbeiterData[columnName] + "', ");
                    }
                    else
                    {
                        part1.Append(columnName + ") ");
                        part2.Append("'" + bearbeiterData[columnName] + "');");
                    }
                    colCounter++;
                }

                string insertQuery = part1.ToString() + part2.ToString();

                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(insertQuery, conn))
                {
                    columnsaffected = cmd.ExecuteNonQuery();
                    conn.Close();
                }

                return columnsaffected;
            }
        }

        public static DataTable GetLog()
        {
            dt = new DataTable();
            using (SQLiteConnection conn = CreateConnection())
            {
                conn.Open();
                using (da = new SQLiteDataAdapter(@"select * from LogData", conn))
                {
                    try
                    {
                        da.Fill(dt);
                        conn.Close();

                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                        return null;
                    }

                }
            }
            return dt;
        }

        public static int LogAddProject(String user, Dictionary<string, string> currentTextBoxValues)
        {
            string action = "Add Project\n";
            foreach (var key in currentTextBoxValues.Keys)
            {
                action += key + ": " + currentTextBoxValues[key] + "; ";
            }
            int rowsaffected = 0;
            string insertQuery = "insert into LogData(user,dateandtime,action) values('"+user+"', '"+(System.DateTime.Now).ToString()+"', '"+action+"')";
            using (SQLiteConnection conn = CreateConnection())
            {
                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(insertQuery, conn))
                {
                    rowsaffected = cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            return rowsaffected;
        }

        public static int LogUpdateProject(String user, Dictionary<string, string> previousTextBoxValues, Dictionary<string, string> currentTextBoxValues)
        {
            string action = "update Projectnumber "+ previousTextBoxValues["Projektnummer"]+ "\n";
            foreach (var key in previousTextBoxValues.Keys)
            {
                if (!previousTextBoxValues[key].Equals(currentTextBoxValues[key]))
                {
                    action += key + ": from " + previousTextBoxValues[key] + " to " + currentTextBoxValues[key];
                }
            }

            int rowsaffected = 0;
            string insertQuery = "insert into LogData(user,dateandtime,action) values('" + user + "', '" + (System.DateTime.Now).ToString() + "', '" + action + "')";
            using (SQLiteConnection conn = CreateConnection())
            {
                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(insertQuery, conn))
                {
                    rowsaffected = cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            return rowsaffected;
        }
    }
}
