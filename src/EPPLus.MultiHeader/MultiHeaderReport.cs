using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.CodeDom.Compiler;
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
                if (columnInfo.IsMultiValue)
                {
                    int c = columnInfo.Index;
                    foreach(string key in columnInfo.DisplayName.Split(','))
                    {
                        _sheet.Cells[row, c++].Value = key;
                    }
                }
                else
                {
                    _sheet.Cells[row, columnInfo.Index].Value = columnInfo.DisplayName;
                }
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

    }
}