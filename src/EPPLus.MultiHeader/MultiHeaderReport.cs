using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class MultiHeaderReport<T>
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _xls;

        private int FirstDataRow => _header?.Height + 1 ?? 2;
        private int row;
        protected HeaderManager<T>? _header;

        public const string HeaderStyleName = "Headers";

        protected Dictionary<string, PropertyInfo>? Properties { get; private set; }

        public MultiHeaderReport(ExcelPackage xls, ExcelWorksheet sheet)
        {
            _xls = xls;
            _sheet = sheet;
        }

        public MultiHeaderReport(ExcelPackage xls, string sheetName): this(xls, AddSheet(xls, sheetName)) { }

        public MultiHeaderReport<T> Configure(Action<ConfigurationBuilder<T>> options)
        {
            var builder = new ConfigurationBuilder<T>();
            options?.Invoke(builder);
            _header = new HeaderManager<T>(builder.Build());
            return this;
        }


        public void GenerateReport(IEnumerable<T> data)
        {
            //If no configuration is provided, use default simple headers
            if (_header == null)
            {
                _header = new HeaderManager<T>();
            } else
            {
                _header.BuildHeaders();
            }
            Properties = _header.Properties;
            WriteHeaders();

            row = FirstDataRow;
            foreach (T item in data)
            {
                ProcessRow(item);
            }
            DoFormatting();
        }

        private static ExcelWorksheet AddSheet(ExcelPackage xls, string sheetName)
        {
            if (!xls.Workbook.Worksheets.AsEnumerable().Any(x => x.Name == sheetName))
            {
                xls.Workbook.Worksheets.Add(sheetName);
            }
            return xls.Workbook.Worksheets[sheetName];
        }

        private void ProcessRow(T item)
        {
            foreach (var columnInfo in _header!.Columns)
            {
                if (columnInfo.HasChildren)
                {
                    ProcessRow(columnInfo.Header!, Properties![columnInfo.Name].GetValue(item));
                }
                else
                {
                    columnInfo.WriteCell(_sheet.Cells[row, columnInfo.Index], Properties!, item!);
                }
            }
            row++;
        }

        private void ProcessRow(HeaderManager header, object? item)
        {
            if (item == null)
                return;
            foreach(var columnInfo in header.Columns)
            {
                if (columnInfo.HasChildren)
                {
                    ProcessRow(columnInfo.Header!, header.Properties[columnInfo.Name].GetValue(item));
                }
                else
                {
                    columnInfo.WriteCell(_sheet.Cells[row, columnInfo.Index], header.Properties, item);
                }
            }

        }

        private void WriteHeaders(HeaderManager? header = null, int row = 1)
        {
            header = header ?? _header!;
            foreach (var columnInfo in header.Columns)
            {
                var cell = _sheet.Cells[row, columnInfo.Index];
                columnInfo.WriteHeader(cell);
                columnInfo.FormatHeader(cell, columnInfo.HasChildren ? 1 : header.Height - row + 1);
                if (columnInfo.HasChildren)
                {
                    WriteHeaders(columnInfo.Header!, row + 1);
                }
            }
        }

        private void DoFormatting()
        {
            foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden))
            {
                _sheet.Column(columnInfo.Index).Hidden = true;
            }

            var rangeHeader = _sheet.Cells[1, _header!.Columns.Min(x => x.Index), _header.Height, _header.Width];
            if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == HeaderStyleName) == null)
            {
                BuildDefaultStyle();

            }
            rangeHeader.StyleName = HeaderStyleName;

            bool NeedsCalculate = false;
            foreach(var columnInfo in _header!.Columns.OfType<ColumnFormula>())
            {
                var range = _sheet.Cells[FirstDataRow, columnInfo.Order!.Value, _sheet.Dimension.End.Row, columnInfo.Order!.Value];
                columnInfo.WriteCell(range, Properties!, null);
                NeedsCalculate = true;
            }
            if (NeedsCalculate)
                _sheet.Calculate();
        }

        private void BuildDefaultStyle()
        {
            var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(HeaderStyleName);
            namedStyle.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            namedStyle.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            namedStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            namedStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            namedStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            namedStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            namedStyle.Style.Fill.SetBackground(Color.LightGray, ExcelFillStyle.Solid);
            namedStyle.Style.Font.Bold = true;
        }

    }
}