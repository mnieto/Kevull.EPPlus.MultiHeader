namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelRange
{
    object? Value { get; set; }
    string? Formula { get; set; }
    Uri? Hyperlink { get; set; }
    IExcelStyle Style { get; }
    bool Merge { get; set; }
    IExcelRange Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns);
    IExcelRange Offset(int RowOffset, int ColumnOffset);
    void MergeCells(int fromRow, int fromColumn, int toRow, int toColumn);
    int EndRow { get; }
    int EndColumn { get; }
    bool AutoFilter { get; set; }
    string StyleName { get; set; }

    IExcelRange this[int Row, int Col] { get; }
    IExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol] { get; }
    IExcelRange this[string Address] { get; }
}