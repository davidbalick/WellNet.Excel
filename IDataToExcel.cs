using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WellNet.Excel
{
    public interface IDataToExcel
    {
        void CreateExcel(SqlConnection connection, string sp, string fileName, bool suppressColumnNames, params string[] parms);
        void DataSetToExcel(DataSet dataSet, string excelFileName, bool suppressColumnNames, int? startTable = 0, int? numTables = null, bool allOnSameSheet = false);
    }
}
