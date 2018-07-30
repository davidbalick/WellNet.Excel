using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace WellNet.Excel
{
    internal enum ConditionalStyle
    {
        None,
        Note
    }
    internal enum CustomCellStyle
    {
        Header,
        DateTime,
        DateTime_Note,
        Double,
        Double_Note,
        Note,
        LongString,
        LongString_Note,
        Default
    }

    //https://github.com/tonyqus/npoi/tree/master/examples/xssf
    internal class WorkbookWithStyles : XSSFWorkbook
    {
        private CellStyleLib _cellStyleLib;

        public WorkbookWithStyles()
        {
            _cellStyleLib = new CellStyleLib(this);
        }

        internal void SetValueAndFormat(ICell cell, DateTime value, ConditionalStyle condStyle = ConditionalStyle.None)
        {
            if (value != null)
                cell.SetCellValue(value);
            cell.CellStyle = _cellStyleLib.Get(CustomCellStyle.DateTime, condStyle);
        }

        internal void SetValueAndFormat(ICell cell, double? value, ConditionalStyle condStyle = ConditionalStyle.None)
        {
            if (value.HasValue)
                cell.SetCellValue(value.Value);
            cell.CellStyle = _cellStyleLib.Get(CustomCellStyle.Double, condStyle);
        }

        internal void SetValueAndFormat(ICell cell, int? value, ConditionalStyle condStyle = ConditionalStyle.None)
        {
            if (value.HasValue)
                cell.SetCellValue(value.Value);
            cell.CellStyle = _cellStyleLib.Get(condStyle == ConditionalStyle.None ? CustomCellStyle.Default : CustomCellStyle.Note);
        }

        internal void SetValueAndFormat(ICell cell, string value, ConditionalStyle condStyle = ConditionalStyle.None)
        {
            const int WRAPTHRESHOLD = 50;
            if (!string.IsNullOrEmpty(value))
                cell.SetCellValue(value);
            else
                return;

            var customCellStyle = CustomCellStyle.Default;

            if (!(value.Length <= WRAPTHRESHOLD && condStyle == ConditionalStyle.None))
            {
                if (value.Length > WRAPTHRESHOLD)
                    customCellStyle = condStyle == ConditionalStyle.None ? CustomCellStyle.LongString : CustomCellStyle.LongString_Note;
                else
                    customCellStyle = CustomCellStyle.Note;
            }
            cell.CellStyle = _cellStyleLib.Get(customCellStyle);
        }

        internal void SetHeaderValueAndFormat(ICell cell, string value)
        {
            if (!string.IsNullOrEmpty(value))
                cell.SetCellValue(value);
            cell.CellStyle = _cellStyleLib.Get(CustomCellStyle.Header);
        }

        internal ISheet CreateWorkSheet(string worksheetName)
        {
            ISheet ws = CreateSheet(worksheetName);
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
    }

    class CellStyleLib 
    {
        private Dictionary<CustomCellStyle, ICellStyle> _cache = new Dictionary<CustomCellStyle, ICellStyle>();
        private WorkbookWithStyles _workbook;
        private IFont _headerFont;
        private IFont HeaderFont
        {
            get { return _headerFont ?? (_headerFont = CreateHeaderFont()); }
            set { _headerFont = value; }
        }
        private IFont _defaultFont;
        private IFont DefaultFont
        {
            get { return _defaultFont ?? (_defaultFont = _workbook.CreateFont()); }
            set { _defaultFont = value; }
        }
        private short? _dateFormat = null;
        private short DateFormat
        {
            get
            {
                if (_dateFormat.HasValue)
                    return _dateFormat.Value;
                _dateFormat = _workbook.CreateDataFormat().GetFormat("mm/dd/yyyy");
                return _dateFormat.Value;
            }
        }
        private short? _doubleFormat = null;
        private short DoubleFormat
        {
            get
            {
                if (_doubleFormat.HasValue)
                    return _doubleFormat.Value;
                _doubleFormat = _workbook.CreateDataFormat().GetFormat("$#,###.##");
                return _doubleFormat.Value;
            }
        }

        public CellStyleLib(WorkbookWithStyles workbook)
        {
            _workbook = workbook;
        }

        public ICellStyle Get(CustomCellStyle customCellStyle)
        {
            if (!_cache.ContainsKey(customCellStyle))
                _cache[customCellStyle] = Create(customCellStyle);
            return _cache[customCellStyle];
        }
        public ICellStyle Get(CustomCellStyle rootCustomCellStyle, ConditionalStyle condStyle)
        {
            var s = rootCustomCellStyle.ToString();
            if (condStyle != ConditionalStyle.None)
                s = string.Format("{0}_{1}", s, condStyle);
            var customCellStyle = CustomCellStyle.Default;
            Enum.TryParse<CustomCellStyle>(s, out customCellStyle);
            return Get(customCellStyle);
        }

        private ICellStyle Create(CustomCellStyle customCellStyle)
        {
            var cellStyle = _workbook.CreateCellStyle();
            switch (customCellStyle)
            {
                case (CustomCellStyle.Header):
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.SetFont(HeaderFont);
                    break;
                case (CustomCellStyle.DateTime):
                    cellStyle.DataFormat = DateFormat;
                    cellStyle.SetFont(DefaultFont);
                    break;
                case (CustomCellStyle.DateTime_Note):
                    cellStyle.DataFormat = DateFormat;
                    ApplyConditionalStyle(ConditionalStyle.Note, cellStyle);
                    cellStyle.SetFont(DefaultFont);
                    break;
                case (CustomCellStyle.Double):
                    cellStyle.DataFormat = DoubleFormat;
                    cellStyle.SetFont(DefaultFont);
                    break;
                case (CustomCellStyle.Double_Note):
                    cellStyle.DataFormat = DoubleFormat;
                    cellStyle.SetFont(DefaultFont);
                    ApplyConditionalStyle(ConditionalStyle.Note, cellStyle);
                    break;
                case (CustomCellStyle.Note):
                    cellStyle.SetFont(DefaultFont);
                    ApplyConditionalStyle(ConditionalStyle.Note, cellStyle);
                    break;
                case (CustomCellStyle.LongString):
                    cellStyle.SetFont(DefaultFont);
                    cellStyle.WrapText = true;
                    break;
                case (CustomCellStyle.LongString_Note):
                    cellStyle.SetFont(DefaultFont);
                    cellStyle.WrapText = true;
                    ApplyConditionalStyle(ConditionalStyle.Note, cellStyle);
                    break;
                case (CustomCellStyle.Default):
                    cellStyle.SetFont(DefaultFont);
                    break;

            }
            return cellStyle;
        }

        private void ApplyConditionalStyle(ConditionalStyle condStyle, ICellStyle cellStyle)
        {
            switch (condStyle)
            {
                case (ConditionalStyle.Note):
                    cellStyle.FillForegroundColor = IndexedColors.LightYellow.Index;
                    cellStyle.BorderBottom = BorderStyle.Hair;
                    //cellStyle.FillPattern = FillPattern.FineDots;
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    break;
            }
        }

        private IFont CreateHeaderFont()
        {
            var headerFont = _workbook.CreateFont();
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            return headerFont;
        }
    }
}
