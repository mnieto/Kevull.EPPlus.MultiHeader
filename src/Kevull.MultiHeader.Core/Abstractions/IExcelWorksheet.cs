namespace Kevull.MultiHeader.Core.Abstractions;
public interface IExcelWorksheet
{
    IExcelRange Cells { get; }
    IExcelRange GetCell(int row, int column);
    IExcelRange GetRange(string address);
    IExcelColumn Column(int column);
    void AutoFitColumn(int column, double minWidth, double maxWidth);
    void SetColumnWidth(int column, double width);
    void SetColumnHidden(int column, bool hidden);
    void FreezePanes(int row, int column);
    void SetAutoFilter(int fromRow, int fromColumn, int toRow, int toColumn);
    T? GetValue<T>(int row, int column);
    double DefaultColWidth { get; set; }
    IExcelRange Dimension { get; }
}