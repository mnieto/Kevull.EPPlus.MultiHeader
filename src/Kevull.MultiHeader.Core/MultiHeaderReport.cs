using Kevull.MultiHeader.Core.Columns;
using Kevull.MultiHeader.Core.Abstractions;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace Kevull.MultiHeader.Core;

/// <summary>
/// Given an <see cref="IEnumerable{T}"/> list of objects it creates an in-memory Excel report
/// </summary>
/// <typeparam name="T">Type of objects</typeparam>
public abstract class MultiHeaderReport<T>
{

    private int FirstDataRow => (_header == null || !_header.AppendToExistingReport) ?
                                _header?.FirstRow + _header?.Height ?? 2 :
                                Worksheet.Dimension.EndRow + 1;
    private int row;
    
    /// <summary>
    /// Internal <see cref="HeaderManager{T}"/>
    /// </summary>
    protected HeaderManager<T>? _header;
    
    /// <summary>
    /// Gets the worksheet where the report is being generated
    /// </summary>
    protected abstract IExcelWorksheet Worksheet { get; }
    
    /// <summary>
    /// Gets the workbook that contains the report
    /// </summary>
    protected abstract IExcelWorkbook Workbook { get; }

    internal const string HeaderStyleName = "__Headers__";

    /// <summary>
    /// Object properties associated to the columns
    /// </summary>
    protected Dictionary<string, PropertyInfo>? Properties { get; private set; }



    /// <summary>
    /// Customize the columns and formats during the report generation. See <see cref="ConfigurationBuilder{T}"/>.
    /// </summary>
    /// <param name="options">Lambda expresion to configure the report</param>
    /// <returns><see cref="MultiHeaderReport{T}"/>This allows a fluent style to configure and generate the report</returns>
    public abstract MultiHeaderReport<T> Configure(Action<ConfigurationBuilder<T>> options);


    /// <summary>
    /// Generate the report in Excel
    /// </summary>
    /// <param name="data">Data of tyepe <typeparamref name="T"/></param>
    /// <remarks>If there is any configuration, it will generate the report using the default conventions</remarks>
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
        if (!_header.AppendToExistingReport)
            WriteHeaders();

        row = FirstDataRow;
        foreach (T item in data)
        {
            ProcessRow(item);
        }
        DoFormatting();
        CalulateFormulas();
    }

    /// <summary>
    /// Saves the generated report to a file
    /// </summary>
    /// <param name="fileName">The path and name of the file where the report will be saved</param>
    public void Save(string fileName)
    {
        Workbook.SaveAs(fileName);
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
                columnInfo.WriteCell(Worksheet.Cells[row, columnInfo.Index], Properties!, item!);
            }
        }
        row++;
    }

    private void ProcessRow(HeaderManager header, object? item)
    {
        if (item == null)
            return;
        if (header.Properties == null)
            throw new ArgumentNullException(nameof(header.Properties));
        foreach(var columnInfo in header.Columns)
        {
            if (columnInfo.HasChildren)
            {
                ProcessRow(columnInfo.Header!, header.Properties[columnInfo.Name].GetValue(item));
            }
            else
            {
                columnInfo.WriteCell(Worksheet.Cells[row, columnInfo.Index], header.Properties, item);
            }
        }

    }

    private void WriteHeaders(HeaderManager? header = null, int? topRow = null)
    {
        header = header ?? _header!;
        int row = topRow ?? _header!.FirstRow;
        foreach (var columnInfo in header.Columns)
        {
            var cell = Worksheet.Cells[row, columnInfo.Index];
            columnInfo.WriteHeader(cell);
            columnInfo.FormatHeader(cell, columnInfo.HasChildren ? 1 : header.Height - (row - _header!.FirstRow));
            if (columnInfo.HasChildren)
            {
                WriteHeaders(columnInfo.Header!, row + 1);
            }
        }
    }

    private void DoFormatting()
    {
        if (_header!.AutoFreezePanes)
            Worksheet.FreezePanes(_header.FirstRow + _header!.Height, _header.FirstColumn);

        //Hide columns if needed
        foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden || x.ColumnWidth.Type == WidthType.Hidden ))
        {
            Worksheet.Column(columnInfo.Index).Hidden = true;
        }

        //Autofilter
        int lastHeaderRow = _header.FirstRow + _header.Height - 1;
        int lastHeaderColumn = _header.FirstColumn + _header.Width - 1;
        Worksheet.Cells[lastHeaderRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn].AutoFilter = _header.AutoFilter;

        //Width
        foreach(var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Auto))
        {
            double minWidth = columnInfo.ColumnWidth.MinimumWidth == double.MinValue ? Worksheet.DefaultColWidth : columnInfo.ColumnWidth.MinimumWidth;
            double maxWidth = columnInfo.ColumnWidth.MaximunWidth;
            Worksheet.Column(columnInfo.Index).AutoFit(minWidth, maxWidth);
        }
        foreach (var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Custom))
        {
            Worksheet.Column(columnInfo.Index).Width = columnInfo.ColumnWidth.Width!.Value;
        }

        //Styles
        BuildDefaultHeaderStyle();
        BuildDateStyle();
        BuildTimeStyle();

        if (!_header!.AppendToExistingReport)
        {
            var rangeHeader = Worksheet.Cells[_header.FirstRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn];
            rangeHeader.StyleName = StyleNames.HeaderStyleName;
        }

        foreach (var columnInfo in _header!.Columns.Where(x => x.StyleName != null))
        {
            var range = Worksheet.Cells[FirstDataRow, columnInfo.Index, Worksheet.Dimension.EndRow, columnInfo.Index];
            range.StyleName = columnInfo.StyleName;
        }
    }

    private void CalulateFormulas()
    {
        bool NeedsCalculate = false;
        foreach (var columnInfo in _header!.Columns.OfType<ColumnFormula>())
        {
            var range = Worksheet.Cells[FirstDataRow, columnInfo.Index, Worksheet.Cells.EndRow, columnInfo.Index];
            columnInfo.WriteCell(range, Properties!, null);
            NeedsCalculate = true;
        }
        if (NeedsCalculate)
            Workbook.Calculate();
    }

    /// <summary>
    /// Builds and applies the default header style to the report headers
    /// </summary>
    /// <remarks>
    /// This method is called during the formatting phase of report generation.
    /// Override this method to customize the default appearance of report headers,
    /// including borders, alignment, background color, and font properties.
    /// </remarks>
    protected abstract void BuildDefaultHeaderStyle();

    /// <summary>
    /// Builds and applies the default date style for date columns
    /// </summary>
    /// <remarks>
    /// This method is called during the formatting phase of report generation.
    /// Override this method to define the number format for DateTime and DateOnly columns.
    /// The default format is typically "mm-dd-yy" but can be customized based on locale. See <see cref="StyleNames"/> for specific style names.
    /// </remarks>
    protected abstract void BuildDateStyle();

    /// <summary>
    /// Builds and applies the default time style for time columns
    /// </summary>
    /// <remarks>
    /// This method is called during the formatting phase of report generation.
    /// Override this method to define the number format for TimeOnly columns.
    /// The default format is typically "h:mm:ss AM/PM" but can be customized based on locale. See <see cref="StyleNames"/> for specific style names.
    /// </remarks>
    protected abstract void BuildTimeStyle();

}
