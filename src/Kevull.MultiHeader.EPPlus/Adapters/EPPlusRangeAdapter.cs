using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusRangeAdapter :  IExcelRange
{
    public ExcelRange Range {get; set;}
    public EPPlusRangeAdapter(ExcelRange range) { 
        Range = range;
    }
    public EPPlusRangeAdapter(ExcelRangeBase range)
    {
        //Create a new ExcelRange from the ExcelRangeBase to have full access to all members
        Range = range.Worksheet.Cells[range.Address];
    }

    public IExcelRange this[int Row, int Col] => new EPPlusRangeAdapter(Range[Row, Col]);

    public IExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol] =>
        new EPPlusRangeAdapter(Range[FromRow, FromCol, ToRow, ToCol]);

    public IExcelRange this[string Address] => new EPPlusRangeAdapter(Range[Address]);


    public object? Value
    {
        get => Range.Value;
        set => Range.Value = value;
    }

    public string? Formula
    {
        get => Range.Formula;
        set => Range.Formula = value;
    }

    public Uri? Hyperlink
    {
        get => Range.Hyperlink;
        set => Range.Hyperlink = value;
    }

    public IExcelStyle Style => new EPPlusStyleAdapter(Range.Style);

    public bool Merge
    {
        get => Range.Merge;
        set => Range.Merge = value;
    }

    public int EndRow => Range.End.Row;

    public int EndColumn => Range.End.Column;

    public int StartRow => Range.Start.Row;
    public int StartColumn => Range.Start.Column;

    public bool AutoFilter
    {
        get => Range.AutoFilter;
        set => Range.AutoFilter = value;
    }

    public string StyleName
    {
        get => Range.StyleName;
        set => Range.StyleName = value;
    }

    public void MergeCells(int fromRow, int fromColumn, int toRow, int toColumn)
    {
        Range.Worksheet.Cells[fromRow, fromColumn, toRow, toColumn].Merge = true;
    }


    public IExcelRange Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns)
    {
        //Can't use Offset because it returns ExcelRangeBase, have to calculate the range manually
        var r = Range.Worksheet.Cells[Range.Start.Row + RowOffset, Range.Start.Column + ColumnOffset, Range.Start.Row + RowOffset + NumberOfRows - 1, Range.Start.Column + ColumnOffset + NumberOfColumns - 1];
        var adapter = new EPPlusRangeAdapter(r);
        return adapter;
    }

    public IExcelRange Offset(int RowOffset, int ColumnOffset)
    {
        //Can't use Offset because it returns ExcelRangeBase, have to calculate the range manually
        var r = Range.Worksheet.Cells[Range.Start.Row + RowOffset, Range.Start.Column + ColumnOffset, Range.End.Row + RowOffset, Range.End.Column + ColumnOffset];
        return new EPPlusRangeAdapter(r);
    }

}
