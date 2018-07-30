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
    public class DataToExcel : IDataToExcel
    {
		internal bool ShowChanges;
		
        public void CreateExcel(SqlConnection connection, string sp, string fileName, bool suppressColumnNames, params string[] parms)
        {
            var dataSet = GetData(connection, sp, parms);
            if (!DataSetContainsData(dataSet))
                return;
            DataSetToExcel(dataSet, fileName, suppressColumnNames);
        }

        public void DataSetToExcel(DataSet dataSet, string excelFileName, bool suppressColumnNames, int? startTable = 0,
            int? numTables = null, bool allOnSameSheet = false)
        {
            var hasData = false;
            var workbook = new WorkbookWithStyles();

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
                    ws = workbook.CreateWorkSheet(worksheetName);

                    worksheetName = null;
                    relativeRowNum = 0;
                }
                if (ws == null)
                    ws = workbook.CreateWorkSheet(worksheetName);
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
                        workbook.SetHeaderValueAndFormat(cell, dataTable.Columns[columnIndex].ColumnName);
                    }
                }

                RelatedColumnPairs relatedColumns = null;
                if (ShowChanges)
					relatedColumns = FindRelatedColumns(dataTable);

                //var longCols = new List<int>();
                for (var rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    hasData = true;
                    if (ws == null)
                        throw new Exception("Worksheet is null!!");

                    var dataRow = dataTable.Rows[rowIndex];
                    if (ShowChanges)
                        relatedColumns.DoCompare(dataRow);

                    row = ws.CreateRow(relativeRowNum++);
                    for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                    {
                        var colDataType = dataTable.Columns[columnIndex].DataType;
                        var obj = dataRow[columnIndex];
                        if (obj == DBNull.Value)
                            continue;

						var condFormatting = ConditionalStyle.None;
						if (ShowChanges && relatedColumns.IsDifferent(columnIndex))
							condFormatting = ConditionalStyle.Note;

                        var cell = row.CreateCell(columnIndex);

                        if (colDataType == typeof(decimal) || colDataType == typeof(double))
                            workbook.SetValueAndFormat(cell, Convert.ToDouble(obj), condFormatting);
                        else if (colDataType == typeof(DateTime))
                            workbook.SetValueAndFormat(cell, Convert.ToDateTime(obj), condFormatting);
                        else if (colDataType == typeof(int))
                            workbook.SetValueAndFormat(cell, (int)obj, condFormatting);
                        else
                            workbook.SetValueAndFormat(cell, obj.ToString(), condFormatting);
                    }
                }
                if (ShowChanges)
                {
                    ws.CreateFreezePane(0, 1, 0, 1);
                    ws.SetAutoFilter(new CellRangeAddress(0, 0, 0, dataTable.Columns.Count-1));
                }
                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                    ws.AutoSizeColumn(columnIndex);
            }
            if (!hasData)
                return;
            var tempFile = Path.GetTempFileName();
            using (var fileData = new FileStream(tempFile, FileMode.Create))
            {
                workbook.Write(fileData);
            }
            try
            {
                File.Delete(excelFileName);
                File.Move(tempFile, excelFileName);
            } catch (Exception ex)
            {
                throw new Exception(string.Format("Error moving {0}: {1}", tempFile, ex.Message));
            }
        }

        private static DataSet GetData(SqlConnection conn, string sp, params string[] parms)
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

        private static RelatedColumnPairs FindRelatedColumns(DataTable dataTable)
        {
            var result = new RelatedColumnPairs();
            var colNameInfos = new List<ColInfo>();

            for (var i = 0; i < dataTable.Columns.Count; i++)
                colNameInfos.Add(new ColInfo(dataTable.Columns[i].ColumnName, i));

            for (var i = 0; i < colNameInfos.Count; i++)
            {
                var thisRelatedPart = colNameInfos[i].RelatedPart;
                if (colNameInfos.Count(cni => cni.RelatedPart == thisRelatedPart ) != 2)
                    continue;
                result.Add(new RelatedColumnPair(new[] { i, colNameInfos.Single(cni => cni.ColIndex != i && cni.RelatedPart == thisRelatedPart).ColIndex }));
            }
            return result;
        }
    }
    class ColInfo
    {
        public string RelatedPart { get; set; }
        public int ColIndex { get; set; }

        public ColInfo(string columnName, int colIndex)
        {
            ColIndex = colIndex;
            RelatedPart = GetPart1(columnName);
        }

        private static string GetPart1(string s)
        {
            var i = s.ToUpper().Trim().Replace(" ", "_").IndexOf("_");
            if (i == -1)
                return Guid.NewGuid().ToString();
            return s.Substring(i + 1);
        }
    }
    class RelatedColumnPair
    {
        public readonly int[] Columns;
        public bool IsDifferent;

        public RelatedColumnPair(int[] pair)
        {
            Columns = pair;
        }

        public void DoCompare(DataRow dataRow)
        {
            IsDifferent = dataRow[Columns[0]].ToString().ToUpper() != dataRow[Columns[1]].ToString().ToUpper();
        }
    }
    class RelatedColumnPairs : List<RelatedColumnPair>
    {
        public void DoCompare(DataRow dataRow)
        {
            ForEach(rcp => rcp.DoCompare(dataRow));
        }

        public bool IsDifferent(int colIndex)
        {
            var relatedColumnPair = this.FirstOrDefault(rcp => rcp.Columns.Contains(colIndex));
            return relatedColumnPair == null ? false : relatedColumnPair.IsDifferent;
        }
    }


}
