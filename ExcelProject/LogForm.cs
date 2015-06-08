using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject
{
    public partial class LogForm : Form
    {
        DataTable dt = null;
        public LogForm()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void LogForm_Load(object sender, EventArgs e)
        {
            dt = SQLiteLogic.GetLog();
            LogDataGridView.DataSource = dt;
            this.Controls.Add(LogDataGridView);
            LogDataGridView.Columns[0].Width = 50;
            LogDataGridView.Columns[1].Width = 100;
            LogDataGridView.Columns[2].Width = 100;
            LogDataGridView.Columns[3].Width = 350;
            int amountOfRows = dt.Rows.Count;
            int rowSize = 40;
            this.Height = amountOfRows * rowSize;
        }
    }
}
