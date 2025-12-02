// Kevull.MultiHeader.Core/Abstractions/IExcelRange.cs
public interface IExcelRange
{
    object? Value { get; set; }
    string? Formula { get; set; }
    Uri? Hyperlink { get; set; }
    IExcelStyle Style { get; }
    bool Merge { get; set; }
    
    void MergeCells(int fromRow, int fromColumn, int toRow, int toColumn);
}