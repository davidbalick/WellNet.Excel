using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

//using NPOI.HSSF.UserModel;
//using NPOI.SS.UserModel;

namespace WellNet.Excel
{
    public static class DataToExcel
    {
        private static XSSFWorkbook _workbook;

        private static XSSFWorkbook Workbook
        {
            get { return _workbook; }
            set
            {
                _workbook = value;
                _headerCellStyle = null;
                _dateCellStyle = null;
                _currCellStyle = null;
                _longStringCellStyle = null;
            }
        }

        private static ICellStyle _headerCellStyle;
        private static ICellStyle HeaderCellStyle
        {
            get
            {
                if (_headerCellStyle != null) return _headerCellStyle;
                _headerCellStyle = Workbook.CreateCellStyle();
                _headerCellStyle.BorderBottom = BorderStyle.Thin;
                var font = Workbook.CreateFont();
                font.Boldweight = (short) FontBoldWeight.Bold;
                _headerCellStyle.SetFont(font);
                return _headerCellStyle;
            }
        }

        private static ICellStyle _dateCellStyle;
        private static ICellStyle DateCellStyle
        {
            get
            {
                if (_dateCellStyle != null) return _dateCellStyle;
                _dateCellStyle = Workbook.CreateCellStyle();
                const string dateFormat = "mm/dd/yyyy";
                var formatId = HSSFDataFormat.GetBuiltinFormat(dateFormat);
                _dateCellStyle.DataFormat = formatId == -1 ? Workbook.CreateDataFormat().GetFormat(dateFormat) : formatId;
                return _dateCellStyle;
            }
        }
        private static ICellStyle _currCellStyle;
        private static ICellStyle CurrCellStyle
        {
            get
            {
                if (_currCellStyle != null) return _currCellStyle;
                _currCellStyle = Workbook.CreateCellStyle();
                const string currFormat = "$0.00";
                var formatId = HSSFDataFormat.GetBuiltinFormat(currFormat);
                _currCellStyle.DataFormat = formatId == -1 ? Workbook.CreateDataFormat().GetFormat(currFormat) : formatId;
                return _currCellStyle; 
            }
        }
        private static ICellStyle _longStringCellStyle;
        private static ICellStyle LongStringCellStyle
        {
            get
            {
                if (_longStringCellStyle != null) return _longStringCellStyle;
                _longStringCellStyle = Workbook.CreateCellStyle();
                _longStringCellStyle.WrapText = true;
                return _longStringCellStyle;
            }
        }
        
        public static void CreateExcel(SqlConnection connection, string sp, string fileName, bool suppressColumnNames, params string[] parms)
        {
            var dataSet = GetData(connection, sp, parms);
            if (!DataSetContainsData(dataSet))
                return;
            DataSetToExcel(dataSet, fileName, suppressColumnNames);
        }

        public static void DataSetToExcel(DataSet dataSet, string excelFileName, bool suppressColumnNames, int? startTable = 0, 
            int? numTables = null, bool allOnSameSheet = false)
        {
            var hasData = false;
            Workbook = new XSSFWorkbook();

            string worksheetName = null;

            var relativeTableNum = 0;
            var relativeRowNum = 0;
            var iNumTables = numTables ?? dataSet.Tables.Count;
            var iStartTable = startTable ?? 0;
            var lastTableNum = iStartTable + iNumTables;
            //var isFirstTable = true;
            ISheet ws = null;
            //var gotCustomWorksheetName = false;
            for (var tableNum = iStartTable; tableNum < lastTableNum; tableNum++)
            {
                var dataTable = dataSet.Tables[tableNum];
                worksheetName = worksheetName ?? dataTable.TableName ?? string.Format("Sheet{0}", ++relativeTableNum);
                if (!allOnSameSheet)
                {
                    if (dataTable.Rows.Count == 1 && dataTable.Columns.Count == 1) //&& !gotCustomWorksheetName)
                    {
                        worksheetName = dataTable.Rows[0][0].ToString();
                        //gotCustomWorksheetName = true;
                        continue;
                    }
                    ws = CreateWorkSheet(worksheetName);

                    worksheetName = null;
                    relativeRowNum = 0;
                }
                if (ws == null)
                    ws = CreateWorkSheet(worksheetName);
                if (allOnSameSheet)
                    relativeRowNum++;
                IRow row;
                if (!suppressColumnNames)
                {
                    if (ws == null)
                        throw new Exception("Worksheet is null!!");
                    row = ws.CreateRow(relativeRowNum++);
                    for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                    {
                        var cell = row.CreateCell(columnIndex);
                        cell.SetCellValue(dataTable.Columns[columnIndex].ColumnName);
                        cell.CellStyle = HeaderCellStyle;
                    }
                }
                //var longCols = new List<int>();
                for (var rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    hasData = true;
                    if (ws == null)
                        throw new Exception("Worksheet is null!!");
                    row = ws.CreateRow(relativeRowNum++);
                    for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                    {
                        var colDataType = dataTable.Columns[columnIndex].DataType;
                        var obj = dataTable.Rows[rowIndex][columnIndex];
                        if (obj == DBNull.Value)
                            continue;
                        var cell = row.CreateCell(columnIndex);
                        if (colDataType == typeof(decimal))
                        {
                            cell.SetCellValue(Convert.ToDouble(obj));
                            cell.CellStyle = CurrCellStyle;
                            continue;
                        }
                        if (colDataType == typeof(DateTime))
                        {
                            cell.SetCellValue(Convert.ToDateTime(obj));
                            cell.CellStyle = DateCellStyle;
                            continue;
                        }
                        if (colDataType == typeof(int))
                        {
                            cell.SetCellValue((int)obj);
                            continue;
                        }
                        var s = obj.ToString();
                        cell.SetCellValue(s);
                        if (s.Length > 50)
                            cell.CellStyle = LongStringCellStyle;
                        //if (s.Length > 30 && !longCols.Contains(columnIndex))
                        //    longCols.Add(columnIndex);
                    }
                }
                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                {
                    ws.AutoSizeColumn(columnIndex);
                    //if (ws == null)
                    //    throw new Exception("Worksheet is null!");
                    //if (longCols.Contains(columnIndex))
                    //    ws.SetColumnWidth(columnIndex, 30 * 256);
                    //else
                    //    ws.AutoSizeColumn(columnIndex);
                }
            }
            if (!hasData)
                return;
            //workbook.GetSheetAt(0).SetAutoFilter(new CellRangeAddress(0, 100, 0, 5));
            using (var fileData = new FileStream(excelFileName, FileMode.Create))
            {
                Workbook.Write(fileData);
            }
        }

        private static ISheet CreateWorkSheet(string worksheetName)
        {
            ISheet ws = Workbook.CreateSheet(worksheetName);
            ws.Autobreaks = true;
            ws.FitToPage = true;
            ws.SetMargin(MarginType.LeftMargin, .25);
            ws.SetMargin(MarginType.RightMargin, .25);
            ws.SetMargin(MarginType.TopMargin, .25);
            ws.SetMargin(MarginType.BottomMargin, .25);
            ws.SetMargin(MarginType.FooterMargin, .25);
            ws.SetMargin(MarginType.HeaderMargin, .25);
            ws.RepeatingRows = CellRangeAddress.ValueOf("1");
            ws.PrintSetup.Landscape = true;
            ws.PrintSetup.FitWidth = 1;
            ws.PrintSetup.FitHeight = 0;
            return ws;
        }

        public static DataSet GetData(SqlConnection conn, string sp, params string[] parms)
        {
            var cmd = new SqlCommand(sp, conn) { CommandTimeout = 0, CommandType = CommandType.StoredProcedure };
            if (parms != null && parms.Length > 0)
            {
                conn.Open();
                SqlCommandBuilder.DeriveParameters(cmd);
                conn.Close();
                for (var p = 1; p < cmd.Parameters.Count; p++)
                    cmd.Parameters[p].Value = parms[p - 1];
            }
            var da = new SqlDataAdapter(cmd);
            var result = new DataSet();
            da.Fill(result);
            return result;
        }

        private static bool DataSetContainsData(DataSet dataSet)
        {
            return dataSet.Tables.Cast<DataTable>().Any(dataTable => dataTable.Rows.Count > 0
                || (dataTable.Rows.Count == 1 && dataTable.Columns.Count == 1));
        }
    }
}
