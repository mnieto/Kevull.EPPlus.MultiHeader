// Kevull.MultiHeader.Core/Abstractions/IExcelWorkbook.cs
public interface IExcelWorkbook
{
    IExcelStyleCollection Styles { get; }
    IExcelWorksheet AddWorksheet(string name);
    void Calculate();
}