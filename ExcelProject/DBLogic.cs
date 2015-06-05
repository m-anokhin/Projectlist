using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ExcelProject
{
    interface DBLogic
    {
        DataTable FillDataTable();

        string[] GetColumnNames();

        string[] GetBearbeiterValues();

        int UpdateExcelFile(Dictionary<string, string> original, Dictionary<string, string> current);
    }
}
