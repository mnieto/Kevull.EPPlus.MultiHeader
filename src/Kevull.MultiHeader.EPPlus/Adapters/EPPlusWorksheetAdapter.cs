using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusWorksheetAdapter : IExcelWorksheet
{
    private readonly ExcelWorksheet _worksheet;
    
    public EPPlusWorksheetAdapter(ExcelWorksheet worksheet)
    {
        _worksheet = worksheet;
    }
    
    public IExcelRange Cells => new EPPlusRangeAdapter(_worksheet.Cells);

    public double DefaultColWidth
    {
        get => _worksheet.DefaultColWidth;
        set => _worksheet.DefaultColWidth = value;
    }

    public IExcelRange GetCell(int row, int column)
    {
        return new EPPlusRangeAdapter(_worksheet.Cells[row, column]);
    }

    public IExcelRange GetRange(string address)
    {
        return new EPPlusRangeAdapter(_worksheet.Cells[address]);
    }

    public IExcelColumn Column(int column)
    {
        return new EPPlusColumnAdapter(_worksheet, _worksheet.Column(column));
    }

    public IExcelRange Dimension => new EPPlusRangeAdapter(_worksheet.Dimension);
    public void AutoFitColumn(int column, double minWidth, double maxWidth)
    {
        _worksheet.Column(column).AutoFit(minWidth, maxWidth);
    }

    public void SetColumnWidth(int column, double width)
    {
        _worksheet.Column(column).Width = width;
    }

    public void SetColumnHidden(int column, bool hidden)
    {
        _worksheet.Column(column).Hidden = hidden;
    }

    public void FreezePanes(int row, int column)
    {
        _worksheet.View.FreezePanes(row, column);
    }

    public void SetAutoFilter(int fromRow, int fromColumn, int toRow, int toColumn)
    {
        _worksheet.Cells[fromRow, fromColumn, toRow, toColumn].AutoFilter = true;
    }

    public T? GetValue<T>(int row, int column)
    {
        return _worksheet.Cells[row, column].GetValue<T>();
    }
}