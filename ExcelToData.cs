using System;
using System.Data;
using System.Globalization;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace WellNet.Excel
{
    public static class ExcelToData
    {
        public delegate void ProgressHandler(ProgressEventArgs e);
        public static event ProgressHandler ProgressEvent;

        //This assumes that the first row contains the column names
        public static DataSet XlsToDataSet(string excelFile, bool columnsOnly = false)
        {
            if (string.IsNullOrEmpty(excelFile))
                throw new Exception("No excel file specified");
            if (!File.Exists(excelFile))
                throw new Exception(string.Format("Cannot find {0}", excelFile));
            FileStream fs;
            try
            {
                fs = new FileStream(excelFile, FileMode.Open, FileAccess.Read);
            }
            catch
            {
                throw new Exception(string.Format("Cannot open {0}", excelFile));
            }
            var ext = Path.GetExtension(excelFile);
            if (string.IsNullOrEmpty(ext))
                return null;
            IWorkbook workBook;
            if (ext.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
                workBook = new HSSFWorkbook(fs);
            else
                workBook = new XSSFWorkbook(fs);
            var dataSet = new DataSet(Path.GetFileNameWithoutExtension(excelFile));
            for (var sheetNum = 0; sheetNum < workBook.NumberOfSheets; sheetNum++)
            {
                var workSheet = workBook.GetSheetAt(sheetNum);
                var excelRow = workSheet.GetRow(0);
                if (excelRow == null)
                    continue;
                var dataTable = new DataTable(workSheet.SheetName);
                var foundAColumn = false;
                for (var cellNum = 0; cellNum < excelRow.PhysicalNumberOfCells; cellNum++)
                {
                    var cell = excelRow.Cells[cellNum];
                    if (cell == null)
                        continue;
                    var value = cell.StringCellValue.Replace("\n", string.Empty).Trim();
                    if (string.IsNullOrEmpty(value))
                        continue;
                    for (var i = dataTable.Columns.Count; i <= cell.ColumnIndex - 1; i++)
                        dataTable.Columns.Add(string.Format("Column{0}", i+1), typeof(string));
                    dataTable.Columns.Add(value, typeof(string));
                    foundAColumn = true;
                }
                if (!foundAColumn)
                    continue;
                dataSet.Tables.Add(dataTable);
                if (columnsOnly)
                    continue;
                for (var rowNum = 1; rowNum < workSheet.PhysicalNumberOfRows; rowNum++)
                {
                    if (ProgressEvent != null && rowNum % 10 == 0)
                        ProgressEvent(new ProgressEventArgs(string.Format("Worksheet {0} of {1}, Row {2} of {3}", sheetNum, workBook.NumberOfSheets, rowNum, workSheet.PhysicalNumberOfRows)));
                    excelRow = workSheet.GetRow(rowNum);
                    if (excelRow == null)
                        continue;
                    var dataRow = dataTable.NewRow();
                    foreach (var cell in excelRow.Cells)
                    {
                        string value;
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                    value = cell.DateCellValue.ToString("yyyy-MM-dd");
                                else if (cell.CellStyle.DataFormat == 9)
                                    value = (cell.NumericCellValue * 100).ToString(CultureInfo.CurrentCulture) + "%";
                                else
                                    value = cell.NumericCellValue.ToString(CultureInfo.CurrentCulture);
                                break;
                            case CellType.String:
                                value = cell.StringCellValue;
                                break;
                            case CellType.Boolean:
                                value = cell.BooleanCellValue ? "Yes" : "No";
                                break;
                            default:
                                continue;
                        }
                        for (var i = dataTable.Columns.Count - 1; i < cell.ColumnIndex; i++)
                            dataTable.Columns.Add(string.Format("Column{0}", i+1), typeof (string));
                        dataRow[cell.ColumnIndex] = value;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            return dataSet;
        }
    }
    public class ProgressEventArgs : EventArgs
    {
        public string Progress { get; set; }
        public ProgressEventArgs(string progress)
        {
            Progress = progress;
        }
    }
}
