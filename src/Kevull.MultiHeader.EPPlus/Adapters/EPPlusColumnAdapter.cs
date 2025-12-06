using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusColumnAdapter : IExcelColumn
{
    private readonly ExcelColumn _column;
    private readonly ExcelWorksheet _worksheet;
    private readonly int _columnNumber;

    public EPPlusColumnAdapter(ExcelWorksheet worksheet, ExcelColumn column)
    {
        _column = column;
        _worksheet = worksheet;
        _columnNumber = column.ColumnMin;
    }

    public bool Hidden
    {
        get => _column.Hidden;
        set => _column.Hidden = value;
    }

    public double Width
    {
        get => _column.Width;
        set => _column.Width = value;
    }

    public void AutoFit(double MinimumWidth, double MaximumWidth)
    {
        _column.AutoFit(MinimumWidth, MaximumWidth);
    }

}
