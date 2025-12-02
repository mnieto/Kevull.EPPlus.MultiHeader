// Kevull.MultiHeader.Core/Abstractions/IExcelWorksheet.cs
public interface IExcelWorksheet
{
    IExcelRange Cells { get; }
    IExcelRange GetCell(int row, int column);
    IExcelRange GetRange(string address);
    void AutoFitColumn(int column, double minWidth, double maxWidth);
    void SetColumnWidth(int column, double width);
    void SetColumnHidden(int column, bool hidden);
    void FreezePanes(int row, int column);
    void SetAutoFilter(int fromRow, int fromColumn, int toRow, int toColumn);
    T? GetValue<T>(int row, int column);
}